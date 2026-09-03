# ============================================================
# QA COMMENT UPDATER - FAST STREAMLIT VERSION
# Input Excel + Reference Excel -> Updated Excel + Summary
# ============================================================

import io
import zipfile
from pathlib import Path

import streamlit as st
import polars as pl


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="QA Comment Updater",
    page_icon="✅",
    layout="wide"
)


# ============================================================
# PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
REFERENCE_FILE = BASE_DIR / "reference.xlsx"


# ============================================================
# CONSTANTS
# ============================================================

REQUIRED_INPUT_COLUMNS = [
    "FunctionName",
    "IsMultiValue",
    "HasBlankValue",
    "QAComment",
]

REQUIRED_REFERENCE_COLUMNS = [
    "FunctionName",
    "Action",
    "Area",
    "Group",
    "Team leader",
]


# ============================================================
# HELPERS
# ============================================================

def clean_column_names(df: pl.DataFrame) -> pl.DataFrame:
    """
    Clean Excel column names.
    """
    new_names = {}

    for col in df.columns:
        new_names[col] = str(col).strip()

    return df.rename(new_names)


def normalize_columns(df: pl.DataFrame, required_columns: list[str]) -> pl.DataFrame:
    """
    Rename columns case-insensitively so small differences
    in Excel headers do not break the app.
    """

    lookup = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    rename_map = {}

    for required in required_columns:
        key = required.strip().lower()

        if key not in lookup:
            raise ValueError(
                f"Missing required column: '{required}'"
            )

        actual_name = lookup[key]

        if actual_name != required:
            rename_map[actual_name] = required

    if rename_map:
        df = df.rename(rename_map)

    return df


def normalize_bool_expression(column: str) -> pl.Expr:
    """
    Convert TRUE/FALSE-like Excel values into normalized strings.
    """

    return (
        pl.col(column)
        .cast(pl.String, strict=False)
        .str.strip_chars()
        .str.to_lowercase()
    )


# ============================================================
# LOAD REFERENCE
# ============================================================

@st.cache_data(show_spinner=False)
def load_reference() -> pl.DataFrame:

    if not REFERENCE_FILE.exists():
        raise FileNotFoundError(
            f"reference.xlsx was not found in:\n{REFERENCE_FILE}"
        )

    try:
        reference = pl.read_excel(
            REFERENCE_FILE,
            engine="calamine",
            infer_schema_length=1000,
        )

    except Exception as e:
        raise RuntimeError(
            f"Could not read reference.xlsx:\n{e}"
        )

    reference = clean_column_names(reference)

    reference = normalize_columns(
        reference,
        REQUIRED_REFERENCE_COLUMNS
    )

    # Keep only the columns we actually need
    reference = reference.select(
        REQUIRED_REFERENCE_COLUMNS
    )

    # Convert everything to string
    for col in REQUIRED_REFERENCE_COLUMNS:
        reference = reference.with_columns(
            pl.col(col)
            .cast(pl.String, strict=False)
            .fill_null("-")
            .str.strip_chars()
            .alias(col)
        )

    # Create lookup key
    reference = reference.with_columns(
        pl.col("FunctionName")
        .str.to_lowercase()
        .alias("_lookup")
    )

    # Remove duplicate FunctionName entries
    reference = reference.unique(
        subset=["_lookup"],
        keep="first"
    )

    return reference


# ============================================================
# PROCESS ONE FILE
# ============================================================

def process_file(
    uploaded_file,
    reference: pl.DataFrame
):
    """
    Process one uploaded Excel file.
    """

    file_bytes = uploaded_file.getvalue()

    # --------------------------------------------------------
    # Read Excel using Polars + Calamine
    # --------------------------------------------------------

    try:
        df = pl.read_excel(
            io.BytesIO(file_bytes),
            engine="calamine",
            infer_schema_length=1000,
        )

    except Exception as e:
        raise RuntimeError(
            f"Could not read '{uploaded_file.name}':\n{e}"
        )

    # --------------------------------------------------------
    # Clean columns
    # --------------------------------------------------------

    df = clean_column_names(df)

    df = normalize_columns(
        df,
        REQUIRED_INPUT_COLUMNS
    )

    # --------------------------------------------------------
    # Make required columns strings
    # --------------------------------------------------------

    df = df.with_columns([
        pl.col("FunctionName")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("QAComment")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars(),
    ])

    # --------------------------------------------------------
    # Normalize boolean columns
    # --------------------------------------------------------

    df = df.with_columns([
        normalize_bool_expression("IsMultiValue")
        .alias("_multi"),

        normalize_bool_expression("HasBlankValue")
        .alias("_blank"),
    ])

    # ========================================================
    # QA COMMENT LOGIC
    # ========================================================

    # Packing:
    #
    # TRUE  + FALSE -> ok
    # TRUE  + TRUE  -> null
    # FALSE + FALSE -> conflict
    # FALSE + TRUE  -> conflict
    #
    # Other:
    #
    # TRUE  + FALSE -> conflict
    # TRUE  + TRUE  -> conflict
    # FALSE + FALSE -> ok
    # FALSE + TRUE  -> null

    is_packing = (
        pl.col("FunctionName")
        .str.to_lowercase()
        == "packing"
    )

    new_comment = (
        pl.when(is_packing)
        .then(
            pl.when(
                (pl.col("_multi") == "true")
                & (pl.col("_blank") == "false")
            )
            .then(pl.lit("ok"))

            .when(
                (pl.col("_multi") == "true")
                & (pl.col("_blank") == "true")
            )
            .then(pl.lit("null"))

            .when(
                (pl.col("_multi") == "false")
                & (pl.col("_blank") == "false")
            )
            .then(pl.lit("conflict"))

            .when(
                (pl.col("_multi") == "false")
                & (pl.col("_blank") == "true")
            )
            .then(pl.lit("conflict"))

            .otherwise(
                pl.col("QAComment")
            )
        )

        .otherwise(
            pl.when(
                (pl.col("_multi") == "true")
                & (pl.col("_blank") == "false")
            )
            .then(pl.lit("conflict"))

            .when(
                (pl.col("_multi") == "true")
                & (pl.col("_blank") == "true")
            )
            .then(pl.lit("conflict"))

            .when(
                (pl.col("_multi") == "false")
                & (pl.col("_blank") == "false")
            )
            .then(pl.lit("ok"))

            .when(
                (pl.col("_multi") == "false")
                & (pl.col("_blank") == "true")
            )
            .then(pl.lit("null"))

            .otherwise(
                pl.col("QAComment")
            )
        )
    )

    df = df.with_columns(
        new_comment.alias("QAComment")
    )

    # ========================================================
    # CREATE LOOKUP
    # ========================================================

    df = df.with_columns(
        pl.col("FunctionName")
        .str.to_lowercase()
        .str.strip_chars()
        .alias("_lookup")
    )

    # ========================================================
    # JOIN WITH REFERENCE
    # ========================================================

    df = df.join(
        reference.select([
            "_lookup",
            "Action",
            "Area",
            "Group",
            "Team leader",
        ]),
        on="_lookup",
        how="left",
    )

    # --------------------------------------------------------
    # Fill missing reference values
    # --------------------------------------------------------

    df = df.with_columns([
        pl.col("Action")
        .cast(pl.String, strict=False)
        .fill_null("-")
        .alias("Action"),

        pl.col("Area")
        .cast(pl.String, strict=False)
        .fill_null("-")
        .alias("Area"),

        pl.col("Group")
        .cast(pl.String, strict=False)
        .fill_null("-")
        .alias("Group"),

        pl.col("Team leader")
        .cast(pl.String, strict=False)
        .fill_null("-")
        .alias("Team leader"),
    ])

    # ========================================================
    # REMOVE ANALYSIS
    # ========================================================

    df = df.filter(
        pl.col("Action")
        .str.to_lowercase()
        != "analysis"
    )

    # ========================================================
    # REMOVE OK
    # ========================================================

    df = df.filter(
        pl.col("QAComment")
        .str.to_lowercase()
        != "ok"
    )

    # ========================================================
    # FIRST COLUMNS
    # ========================================================

    first_columns = [
        "FunctionName",
        "Area",
        "Group",
        "Team leader",
    ]

    remaining_columns = [
        col
        for col in df.columns
        if col not in first_columns
        and col not in [
            "_lookup",
            "_multi",
            "_blank",
            "Action",
        ]
    ]

    final_columns = first_columns + remaining_columns

    df = df.select(final_columns)

    # ========================================================
    # SUMMARY
    # ========================================================

    summary = (
        df
        .group_by([
            "Area",
            "Team leader",
            "QAComment",
        ])
        .agg(
            pl.len().alias("Count")
        )
        .sort([
            "Area",
            "Team leader",
            "QAComment",
        ])
    )

    return df, summary


# ============================================================
# WRITE EXCEL
# ============================================================

def dataframe_to_excel(
    df: pl.DataFrame,
    sheet_name: str = "Output"
) -> bytes:

    output = io.BytesIO()

    df.write_excel(
        output,
        worksheet=sheet_name,
        autofit=False,
    )

    return output.getvalue()


# ============================================================
# CREATE ZIP
# ============================================================

def create_zip(files: list[tuple[str, bytes]]) -> bytes:

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as z:

        for filename, content in files:
            z.writestr(
                filename,
                content
            )

    return zip_buffer.getvalue()


# ============================================================
# UI
# ============================================================

st.title("✅ QA Comment Updater")

st.write(
    "Upload one or more Excel files to update QA Comments "
    "using the Reference mapping."
)


# ============================================================
# CHECK REFERENCE
# ============================================================

try:

    reference_df = load_reference()

    st.success(
        f"Reference loaded successfully — "
        f"{reference_df.height:,} functions"
    )

except Exception as e:

    st.error(str(e))

    st.stop()


# ============================================================
# UPLOAD FILES
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Input Excel File(s)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)


# ============================================================
# PREVIEW OPTION
# ============================================================

show_preview = st.checkbox(
    "Show preview after processing",
    value=False
)


# ============================================================
# PROCESS BUTTON
# ============================================================

if uploaded_files:

    st.info(
        f"{len(uploaded_files)} file(s) selected."
    )

    if st.button(
        "🚀 Process Files",
        type="primary",
        use_container_width=True
    ):

        output_files = []
        all_summaries = []

        progress = st.progress(0)

        status = st.empty()

        total_files = len(uploaded_files)

        for index, uploaded_file in enumerate(
            uploaded_files,
            start=1
        ):

            status.write(
                f"Processing {index}/{total_files}: "
                f"{uploaded_file.name}"
            )

            try:

                # --------------------------------------------
                # PROCESS
                # --------------------------------------------

                result_df, summary_df = process_file(
                    uploaded_file,
                    reference_df
                )

                # --------------------------------------------
                # OUTPUT FILE NAME
                # --------------------------------------------

                original_name = Path(
                    uploaded_file.name
                ).stem

                output_name = (
                    f"{original_name}_Updated.xlsx"
                )

                # --------------------------------------------
                # WRITE EXCEL
                # --------------------------------------------

                excel_bytes = dataframe_to_excel(
                    result_df,
                    "Output"
                )

                output_files.append(
                    (
                        output_name,
                        excel_bytes
                    )
                )

                # --------------------------------------------
                # ADD FILE NAME TO SUMMARY
                # --------------------------------------------

                summary_df = summary_df.with_columns(
                    pl.lit(uploaded_file.name)
                    .alias("Source File")
                )

                all_summaries.append(
                    summary_df
                )

                # --------------------------------------------
                # PREVIEW ONLY IF REQUESTED
                # --------------------------------------------

                if show_preview:

                    with st.expander(
                        f"Preview: {uploaded_file.name}"
                    ):

                        st.dataframe(
                            result_df
                            .head(100)
                            .to_dicts(),
                            use_container_width=True
                        )

                        st.write(
                            f"Rows in output: "
                            f"{result_df.height:,}"
                        )

            except Exception as e:

                st.error(
                    f"❌ Error processing "
                    f"'{uploaded_file.name}':\n\n{e}"
                )

            progress.progress(
                index / total_files
            )

        status.empty()

        # ====================================================
        # SUMMARY
        # ====================================================

        if all_summaries:

            final_summary = pl.concat(
                all_summaries,
                how="diagonal"
            )

            # Put Source File first
            summary_columns = [
                "Source File",
                "Area",
                "Team leader",
                "QAComment",
                "Count",
            ]

            final_summary = final_summary.select(
                [
                    col
                    for col in summary_columns
                    if col in final_summary.columns
                ]
            )

            summary_bytes = dataframe_to_excel(
                final_summary,
                "Summary"
            )

            st.success(
                f"✅ Finished processing "
                f"{len(output_files)} file(s)"
            )

            # =================================================
            # DOWNLOAD OUTPUTS
            # =================================================

            if len(output_files) == 1:

                filename, content = output_files[0]

                st.download_button(
                    label=f"⬇️ Download {filename}",
                    data=content,
                    file_name=filename,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

            else:

                zip_bytes = create_zip(
                    output_files
                )

                st.download_button(
                    label="🗜️ Download All Updated Files (ZIP)",
                    data=zip_bytes,
                    file_name="QA_Updated_Files.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

                st.subheader("Individual Files")

                for filename, content in output_files:

                    st.download_button(
                        label=f"⬇️ {filename}",
                        data=content,
                        file_name=filename,
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                        key=f"download_{filename}",
                    )

            # =================================================
            # SUMMARY DOWNLOAD
            # =================================================

            st.subheader("📊 QA Summary")

            st.download_button(
                label="⬇️ Download QA Summary",
                data=summary_bytes,
                file_name="QA_Summary.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

            # Show summary only
            st.dataframe(
                final_summary.to_dicts(),
                use_container_width=True
            )

        else:

            st.warning(
                "No files were processed successfully."
            )
