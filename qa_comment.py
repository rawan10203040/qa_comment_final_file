
# ============================================================
# QA COMMENT UPDATER
# Stable + Fast Streamlit Version
#
# Input Excel + Reference Excel
# -> Updated Excel
# -> QA Summary
# ============================================================

import io
import time
from pathlib import Path

import polars as pl
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="QA Comment Updater",
    page_icon="✅",
    layout="wide",
)


# ============================================================
# PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
REFERENCE_FILE = BASE_DIR / "reference.xlsx"


# ============================================================
# CONSTANTS
# ============================================================

INPUT_COLUMNS = [
    "FunctionName",
    "IsMultiValue",
    "HasBlankValue",
    "QAComment",
]

REFERENCE_COLUMNS = [
    "FunctionName",
    "Action",
    "Area",
    "Group",
    "Team leader",
]


# ============================================================
# HELPERS
# ============================================================

def clean_headers(df: pl.DataFrame) -> pl.DataFrame:
    """
    Clean Excel column names.
    """

    return df.rename({
        col: str(col).strip()
        for col in df.columns
    })


def normalize_required_columns(
    df: pl.DataFrame,
    required_columns: list[str],
) -> pl.DataFrame:
    """
    Match required columns ignoring case and surrounding spaces.
    """

    available = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    rename_map = {}

    missing = []

    for required in required_columns:

        key = required.strip().lower()

        if key not in available:
            missing.append(required)
            continue

        actual = available[key]

        if actual != required:
            rename_map[actual] = required

    if missing:
        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
        )

    if rename_map:
        df = df.rename(rename_map)

    return df


def normalize_string_column(
    column_name: str,
) -> pl.Expr:

    return (
        pl.col(column_name)
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars()
    )


def normalize_bool_column(
    column_name: str,
) -> pl.Expr:

    return (
        pl.col(column_name)
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars()
        .str.to_lowercase()
    )


# ============================================================
# LOAD REFERENCE
# ============================================================

@st.cache_data(
    show_spinner=False,
    ttl=3600,
)
def load_reference() -> pl.DataFrame:

    if not REFERENCE_FILE.exists():

        raise FileNotFoundError(
            "reference.xlsx was not found.\n\n"
            "Make sure reference.xlsx exists in the "
            "same GitHub folder as qa_comment.py."
        )

    try:

        reference = pl.read_excel(
            REFERENCE_FILE,
            engine="calamine",
            infer_schema_length=1000,
        )

    except Exception as exc:

        raise RuntimeError(
            "Could not read reference.xlsx.\n\n"
            f"Details: {exc}"
        )

    reference = clean_headers(reference)

    reference = normalize_required_columns(
        reference,
        REFERENCE_COLUMNS,
    )

    reference = reference.select(
        REFERENCE_COLUMNS
    )

    # --------------------------------------------------------
    # Normalize reference values
    # --------------------------------------------------------

    reference = reference.with_columns([

        normalize_string_column(
            "FunctionName"
        ).alias("FunctionName"),

        normalize_string_column(
            "Action"
        ).alias("Action"),

        normalize_string_column(
            "Area"
        ).alias("Area"),

        normalize_string_column(
            "Group"
        ).alias("Group"),

        normalize_string_column(
            "Team leader"
        ).alias("Team leader"),
    ])

    # --------------------------------------------------------
    # Replace empty reference values
    # --------------------------------------------------------

    reference = reference.with_columns([

        pl.when(
            pl.col("Action") == ""
        )
        .then(pl.lit("-"))
        .otherwise(pl.col("Action"))
        .alias("Action"),

        pl.when(
            pl.col("Area") == ""
        )
        .then(pl.lit("-"))
        .otherwise(pl.col("Area"))
        .alias("Area"),

        pl.when(
            pl.col("Group") == ""
        )
        .then(pl.lit("-"))
        .otherwise(pl.col("Group"))
        .alias("Group"),

        pl.when(
            pl.col("Team leader") == ""
        )
        .then(pl.lit("-"))
        .otherwise(pl.col("Team leader"))
        .alias("Team leader"),
    ])

    # --------------------------------------------------------
    # Lookup key
    # --------------------------------------------------------

    reference = reference.with_columns(
        pl.col("FunctionName")
        .str.to_lowercase()
        .str.strip_chars()
        .alias("_lookup")
    )

    # --------------------------------------------------------
    # Remove duplicate FunctionName
    # --------------------------------------------------------

    reference = reference.unique(
        subset=["_lookup"],
        keep="first",
    )

    return reference


# ============================================================
# PROCESS ONE FILE
# ============================================================

def process_file(
    uploaded_file,
    reference: pl.DataFrame,
):

    start_time = time.perf_counter()

    # --------------------------------------------------------
    # Read bytes
    # --------------------------------------------------------

    file_bytes = uploaded_file.getvalue()

    if not file_bytes:
        raise ValueError(
            f"'{uploaded_file.name}' is empty."
        )

    # --------------------------------------------------------
    # Read Excel
    # --------------------------------------------------------

    try:

        df = pl.read_excel(
            io.BytesIO(file_bytes),
            engine="calamine",
            infer_schema_length=1000,
        )

    except Exception as exc:

        raise RuntimeError(
            f"Could not read '{uploaded_file.name}'.\n\n"
            f"Details: {exc}"
        )

    input_rows = df.height

    # --------------------------------------------------------
    # Clean headers
    # --------------------------------------------------------

    df = clean_headers(df)

    # --------------------------------------------------------
    # Validate columns
    # --------------------------------------------------------

    df = normalize_required_columns(
        df,
        INPUT_COLUMNS,
    )

    # --------------------------------------------------------
    # Normalize required columns
    # --------------------------------------------------------

    df = df.with_columns([

        normalize_string_column(
            "FunctionName"
        ).alias("FunctionName"),

        normalize_string_column(
            "QAComment"
        ).alias("QAComment"),

        normalize_bool_column(
            "IsMultiValue"
        ).alias("_multi"),

        normalize_bool_column(
            "HasBlankValue"
        ).alias("_blank"),
    ])

    # ========================================================
    # QA COMMENT LOGIC
    # ========================================================

    is_packing = (
        pl.col("FunctionName")
        .str.to_lowercase()
        .str.strip_chars()
        == "packing"
    )

    packing_logic = (

        pl.when(
            (pl.col("_multi") == "true")
            &
            (pl.col("_blank") == "false")
        )
        .then(pl.lit("ok"))

        .when(
            (pl.col("_multi") == "true")
            &
            (pl.col("_blank") == "true")
        )
        .then(pl.lit("null"))

        .when(
            (pl.col("_multi") == "false")
            &
            (pl.col("_blank") == "false")
        )
        .then(pl.lit("conflict"))

        .when(
            (pl.col("_multi") == "false")
            &
            (pl.col("_blank") == "true")
        )
        .then(pl.lit("conflict"))

        .otherwise(
            pl.col("QAComment")
        )
    )

    normal_logic = (

        pl.when(
            (pl.col("_multi") == "true")
            &
            (pl.col("_blank") == "false")
        )
        .then(pl.lit("conflict"))

        .when(
            (pl.col("_multi") == "true")
            &
            (pl.col("_blank") == "true")
        )
        .then(pl.lit("conflict"))

        .when(
            (pl.col("_multi") == "false")
            &
            (pl.col("_blank") == "false")
        )
        .then(pl.lit("ok"))

        .when(
            (pl.col("_multi") == "false")
            &
            (pl.col("_blank") == "true")
        )
        .then(pl.lit("null"))

        .otherwise(
            pl.col("QAComment")
        )
    )

    df = df.with_columns(
        pl.when(is_packing)
        .then(packing_logic)
        .otherwise(normal_logic)
        .alias("QAComment")
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
    # REFERENCE JOIN
    # ========================================================

    reference_small = reference.select([
        "_lookup",
        "Action",
        "Area",
        "Group",
        "Team leader",
    ])

    df = df.join(
        reference_small,
        on="_lookup",
        how="left",
    )

    # ========================================================
    # FILL REFERENCE VALUES
    # ========================================================

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

    output_rows = df.height

    # ========================================================
    # COLUMN ORDER
    # ========================================================

    first_columns = [
        "FunctionName",
        "Area",
        "Group",
        "Team leader",
    ]

    helper_columns = [
        "_lookup",
        "_multi",
        "_blank",
        "Action",
    ]

    remaining_columns = [
        col
        for col in df.columns
        if col not in first_columns
        and col not in helper_columns
    ]

    df = df.select(
        first_columns + remaining_columns
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    if output_rows > 0:

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

    else:

        summary = pl.DataFrame({
            "Area": [],
            "Team leader": [],
            "QAComment": [],
            "Count": [],
        })

    processing_time = (
        time.perf_counter() - start_time
    )

    return (
        df,
        summary,
        input_rows,
        output_rows,
        processing_time,
    )


# ============================================================
# WRITE EXCEL
# ============================================================

def dataframe_to_excel(
    df: pl.DataFrame,
    sheet_name: str = "Output",
) -> bytes:

    buffer = io.BytesIO()

    df.write_excel(
        buffer,
        worksheet=sheet_name,
        autofit=False,
    )

    return buffer.getvalue()


# ============================================================
# PAGE TITLE
# ============================================================

st.title("✅ QA Comment Updater")

st.caption(
    "Fast and stable Excel processing using Polars"
)


# ============================================================
# LOAD REFERENCE
# ============================================================

try:

    reference_df = load_reference()

    st.success(
        f"Reference loaded successfully — "
        f"{reference_df.height:,} functions"
    )

except Exception as exc:

    st.error(
        "❌ Reference loading failed"
    )

    st.exception(exc)

    st.stop()


# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Input Excel File(s)",
    type=["xlsx"],
    accept_multiple_files=True,
)


# ============================================================
# OPTIONS
# ============================================================

show_preview = st.checkbox(
    "Show preview",
    value=False,
)


# ============================================================
# PROCESS BUTTON
# ============================================================

if uploaded_files:

    st.info(
        f"{len(uploaded_files)} file(s) selected."
    )

    process_button = st.button(
        "🚀 Process Files",
        type="primary",
        use_container_width=True,
    )

    if process_button:

        total_files = len(uploaded_files)

        progress = st.progress(0)

        status = st.empty()

        successful_files = 0
        failed_files = 0

        # ====================================================
        # PROCESS FILES ONE BY ONE
        # ====================================================

        for index, uploaded_file in enumerate(
            uploaded_files,
            start=1,
        ):

            status.write(
                f"Processing "
                f"{index}/{total_files}: "
                f"{uploaded_file.name}"
            )

            try:

                (
                    result_df,
                    summary_df,
                    input_rows,
                    output_rows,
                    processing_time,
                ) = process_file(
                    uploaded_file,
                    reference_df,
                )

                successful_files += 1

                # ------------------------------------------------
                # File name
                # ------------------------------------------------

                original_name = Path(
                    uploaded_file.name
                ).stem

                output_name = (
                    f"{original_name}_Updated.xlsx"
                )

                # ------------------------------------------------
                # Create Excel
                # ------------------------------------------------

                excel_bytes = dataframe_to_excel(
                    result_df,
                    "Output",
                )

                # ------------------------------------------------
                # File information
                # ------------------------------------------------

                st.success(
                    f"✅ {uploaded_file.name} "
                    f"processed successfully"
                )

                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    st.metric(
                        "Input Rows",
                        f"{input_rows:,}",
                    )

                with col2:
                    st.metric(
                        "Output Rows",
                        f"{output_rows:,}",
                    )

                with col3:
                    st.metric(
                        "Removed",
                        f"{input_rows - output_rows:,}",
                    )

                with col4:
                    st.metric(
                        "Time",
                        f"{processing_time:.2f}s",
                    )

                # ------------------------------------------------
                # Download
                # ------------------------------------------------

                st.download_button(
                    label=(
                        f"⬇️ Download "
                        f"{output_name}"
                    ),
                    data=excel_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    key=(
                        f"download_{index}_"
                        f"{uploaded_file.name}"
                    ),
                    use_container_width=True,
                )

                # ------------------------------------------------
                # Preview
                # ------------------------------------------------

                if show_preview and output_rows > 0:

                    with st.expander(
                        f"Preview — "
                        f"{uploaded_file.name}"
                    ):

                        st.dataframe(
                            result_df
                            .head(100)
                            .to_dicts(),
                            use_container_width=True,
                        )

                # ------------------------------------------------
                # Summary
                # ------------------------------------------------

                if summary_df.height > 0:

                    st.subheader(
                        f"📊 Summary — "
                        f"{uploaded_file.name}"
                    )

                    st.dataframe(
                        summary_df.to_dicts(),
                        use_container_width=True,
                    )

                    summary_bytes = (
                        dataframe_to_excel(
                            summary_df,
                            "Summary",
                        )
                    )

                    st.download_button(
                        label=(
                            "⬇️ Download Summary"
                        ),
                        data=summary_bytes,
                        file_name=(
                            f"{original_name}"
                            "_QA_Summary.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                        key=(
                            f"summary_{index}_"
                            f"{uploaded_file.name}"
                        ),
                        use_container_width=True,
                    )

                else:

                    st.info(
                        "No rows remained after "
                        "the QA filtering."
                    )

                # ------------------------------------------------
                # Release large objects
                # ------------------------------------------------

                del result_df
                del summary_df
                del excel_bytes

            except Exception as exc:

                failed_files += 1

                st.error(
                    f"❌ Failed to process "
                    f"'{uploaded_file.name}'"
                )

                st.exception(exc)

            progress.progress(
                index / total_files
            )

        # ====================================================
        # FINAL STATUS
        # ====================================================

        status.empty()

        if failed_files == 0:

            st.success(
                f"🎉 All {successful_files} "
                f"file(s) processed successfully."
            )

        else:

            st.warning(
                f"Finished with "
                f"{successful_files} successful "
                f"and {failed_files} failed file(s)."
            )

else:

    st.info(
        "Please upload one or more .xlsx files "
        "to start."
    )
