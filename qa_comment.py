
# ============================================================
# QA COMMENT UPDATER
# Stable Streamlit Version
#
# Input Excel + Reference Excel
# -> Updated Excel
# -> QA Summary
# ============================================================

import io
import time
from pathlib import Path

import pandas as pd
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
# REQUIRED COLUMNS
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
# HELPER FUNCTIONS
# ============================================================

def clean_headers(df):
    """
    Remove spaces from Excel column names.
    """

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    return df


def normalize_required_columns(
    df,
    required_columns,
):
    """
    Match required columns ignoring:
    - upper/lower case
    - surrounding spaces
    """

    lookup = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    rename_map = {}

    missing = []

    for required in required_columns:

        key = required.strip().lower()

        if key not in lookup:

            missing.append(required)

        else:

            actual = lookup[key]

            if actual != required:

                rename_map[actual] = required

    if missing:

        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
        )

    if rename_map:

        df = df.rename(
            columns=rename_map
        )

    return df


def normalize_string_column(
    series
):
    """
    Convert values to strings safely.
    """

    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
    )


def normalize_boolean(
    series
):
    """
    Normalize TRUE/FALSE values.
    """

    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )


# ============================================================
# LOAD REFERENCE
# ============================================================

@st.cache_data(
    show_spinner=False,
    ttl=3600,
)
def load_reference():

    if not REFERENCE_FILE.exists():

        raise FileNotFoundError(
            "reference.xlsx was not found.\n\n"
            "Make sure reference.xlsx is in the "
            "same folder as qa_comment.py."
        )

    try:

        reference = pd.read_excel(
            REFERENCE_FILE,
            engine="openpyxl",
        )

    except Exception as exc:

        raise RuntimeError(
            "Could not read reference.xlsx.\n\n"
            f"Details: {exc}"
        )

    # --------------------------------------------------------
    # Clean headers
    # --------------------------------------------------------

    reference = clean_headers(
        reference
    )

    # --------------------------------------------------------
    # Validate columns
    # --------------------------------------------------------

    reference = normalize_required_columns(
        reference,
        REFERENCE_COLUMNS,
    )

    # --------------------------------------------------------
    # Keep required columns only
    # --------------------------------------------------------

    reference = reference[
        REFERENCE_COLUMNS
    ].copy()

    # --------------------------------------------------------
    # Normalize values
    # --------------------------------------------------------

    for column in REFERENCE_COLUMNS:

        reference[column] = (
            reference[column]
            .fillna("-")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Lookup key
    # --------------------------------------------------------

    reference["_lookup"] = (
        reference["FunctionName"]
        .str.lower()
        .str.strip()
    )

    # --------------------------------------------------------
    # Remove duplicate FunctionName
    # --------------------------------------------------------

    reference = reference.drop_duplicates(
        subset=["_lookup"],
        keep="first",
    )

    return reference


# ============================================================
# PROCESS ONE FILE
# ============================================================

def process_file(
    uploaded_file,
    reference,
):

    start_time = time.perf_counter()

    # ========================================================
    # READ INPUT
    # ========================================================

    file_bytes = uploaded_file.getvalue()

    if not file_bytes:

        raise ValueError(
            f"'{uploaded_file.name}' is empty."
        )

    try:

        df = pd.read_excel(
            io.BytesIO(file_bytes),
            engine="openpyxl",
        )

    except Exception as exc:

        raise RuntimeError(
            f"Could not read '{uploaded_file.name}'.\n\n"
            f"Details: {exc}"
        )

    input_rows = len(df)

    # ========================================================
    # CLEAN HEADERS
    # ========================================================

    df = clean_headers(df)

    # ========================================================
    # VALIDATE INPUT COLUMNS
    # ========================================================

    df = normalize_required_columns(
        df,
        INPUT_COLUMNS,
    )

    # ========================================================
    # NORMALIZE INPUT VALUES
    # ========================================================

    df["FunctionName"] = normalize_string_column(
        df["FunctionName"]
    )

    df["QAComment"] = normalize_string_column(
        df["QAComment"]
    )

    df["_multi"] = normalize_boolean(
        df["IsMultiValue"]
    )

    df["_blank"] = normalize_boolean(
        df["HasBlankValue"]
    )

    # ========================================================
    # PACKING LOGIC
    # ========================================================

    is_packing = (
        df["FunctionName"]
        .str.lower()
        .str.strip()
        .eq("packing")
    )

    # --------------------------------------------------------
    # Default / normal logic
    # --------------------------------------------------------

    normal_conditions = [

        (
            (df["_multi"] == "true")
            &
            (df["_blank"] == "false")
        ),

        (
            (df["_multi"] == "true")
            &
            (df["_blank"] == "true")
        ),

        (
            (df["_multi"] == "false")
            &
            (df["_blank"] == "false")
        ),

        (
            (df["_multi"] == "false")
            &
            (df["_blank"] == "true")
        ),
    ]

    normal_values = [
        "conflict",
        "conflict",
        "ok",
        "null",
    ]

    # --------------------------------------------------------
    # Apply normal logic
    # --------------------------------------------------------

    for condition, value in zip(
        normal_conditions,
        normal_values,
    ):

        df.loc[
            (~is_packing) & condition,
            "QAComment"
        ] = value

    # ========================================================
    # PACKING LOGIC
    # ========================================================

    packing_conditions = [

        (
            (df["_multi"] == "true")
            &
            (df["_blank"] == "false")
        ),

        (
            (df["_multi"] == "true")
            &
            (df["_blank"] == "true")
        ),

        (
            (df["_multi"] == "false")
            &
            (df["_blank"] == "false")
        ),

        (
            (df["_multi"] == "false")
            &
            (df["_blank"] == "true")
        ),
    ]

    packing_values = [
        "ok",
        "null",
        "conflict",
        "conflict",
    ]

    for condition, value in zip(
        packing_conditions,
        packing_values,
    ):

        df.loc[
            is_packing & condition,
            "QAComment"
        ] = value

    # ========================================================
    # CREATE LOOKUP
    # ========================================================

    df["_lookup"] = (
        df["FunctionName"]
        .str.lower()
        .str.strip()
    )

    # ========================================================
    # REFERENCE JOIN
    # ========================================================

    reference_small = reference[
        [
            "_lookup",
            "Action",
            "Area",
            "Group",
            "Team leader",
        ]
    ].copy()

    df = df.merge(
        reference_small,
        on="_lookup",
        how="left",
    )

    # ========================================================
    # FILL REFERENCE VALUES
    # ========================================================

    for column in [
        "Action",
        "Area",
        "Group",
        "Team leader",
    ]:

        df[column] = (
            df[column]
            .fillna("-")
            .astype(str)
            .str.strip()
        )

        df.loc[
            df[column].eq(""),
            column
        ] = "-"

    # ========================================================
    # REMOVE ACTION = ANALYSIS
    # ========================================================

    df = df[
        df["Action"]
        .str.lower()
        .ne("analysis")
    ].copy()

    # ========================================================
    # REMOVE QA COMMENT = OK
    # ========================================================

    df = df[
        df["QAComment"]
        .str.lower()
        .ne("ok")
    ].copy()

    output_rows = len(df)

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
        column
        for column in df.columns
        if column not in first_columns
        and column not in helper_columns
    ]

    df = df[
        first_columns + remaining_columns
    ]

    # ========================================================
    # SUMMARY
    # ========================================================

    if output_rows > 0:

        summary = (
            df.groupby(
                [
                    "Area",
                    "Team leader",
                    "QAComment",
                ],
                dropna=False,
            )
            .size()
            .reset_index(
                name="Count"
            )
            .sort_values(
                [
                    "Area",
                    "Team leader",
                    "QAComment",
                ]
            )
        )

    else:

        summary = pd.DataFrame(
            columns=[
                "Area",
                "Team leader",
                "QAComment",
                "Count",
            ]
        )

    # ========================================================
    # PROCESSING TIME
    # ========================================================

    processing_time = (
        time.perf_counter()
        - start_time
    )

    return (
        df,
        summary,
        input_rows,
        output_rows,
        processing_time,
    )


# ============================================================
# WRITE DATAFRAME TO EXCEL
# ============================================================

def dataframe_to_excel(
    df,
    sheet_name="Output",
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
    ) as writer:

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
        )

        workbook = writer.book

        worksheet = writer.sheets[
            sheet_name
        ]

        # ----------------------------------------------------
        # Header format
        # ----------------------------------------------------

        header_format = workbook.add_format({
            "bold": True,
            "border": 1,
        })

        for column_number, column_name in enumerate(
            df.columns
        ):

            worksheet.write(
                0,
                column_number,
                column_name,
                header_format,
            )

        # ----------------------------------------------------
        # Freeze header
        # ----------------------------------------------------

        worksheet.freeze_panes(
            1,
            0,
        )

        # ----------------------------------------------------
        # Autofilter
        # ----------------------------------------------------

        if len(df.columns) > 0:

            worksheet.autofilter(
                0,
                0,
                max(len(df), 1),
                len(df.columns) - 1,
            )

        # ----------------------------------------------------
        # Reasonable column widths
        # ----------------------------------------------------

        for index, column in enumerate(
            df.columns
        ):

            try:

                max_length = max(
                    len(str(column)),
                    df[column]
                    .astype(str)
                    .head(1000)
                    .map(len)
                    .max(),
                )

                width = min(
                    max_length + 2,
                    40,
                )

            except Exception:

                width = 15

            worksheet.set_column(
                index,
                index,
                width,
            )

    return output.getvalue()


# ============================================================
# UI
# ============================================================

st.title(
    "✅ QA Comment Updater"
)

st.write(
    "Upload one or more Excel files "
    "to update QA Comments using "
    "the Reference mapping."
)


# ============================================================
# LOAD REFERENCE
# ============================================================

try:

    reference_df = load_reference()

    st.success(
        "Reference loaded successfully — "
        f"{len(reference_df):,} functions"
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
# PREVIEW
# ============================================================

show_preview = st.checkbox(
    "Show preview after processing",
    value=False,
)


# ============================================================
# PROCESS
# ============================================================

if uploaded_files:

    st.info(
        f"{len(uploaded_files)} file(s) selected."
    )

    if st.button(
        "🚀 Process Files",
        type="primary",
        use_container_width=True,
    ):

        progress = st.progress(0)

        status = st.empty()

        successful = 0
        failed = 0

        total = len(uploaded_files)

        # ====================================================
        # PROCESS FILES ONE BY ONE
        # ====================================================

        for index, uploaded_file in enumerate(
            uploaded_files,
            start=1,
        ):

            status.write(
                f"Processing "
                f"{index}/{total}: "
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

                successful += 1

                # ------------------------------------------------
                # Output filename
                # ------------------------------------------------

                original_name = Path(
                    uploaded_file.name
                ).stem

                output_name = (
                    f"{original_name}_Updated.xlsx"
                )

                # ------------------------------------------------
                # Create output
                # ------------------------------------------------

                output_bytes = dataframe_to_excel(
                    result_df,
                    "Output",
                )

                # =================================================
                # METRICS
                # =================================================

                st.success(
                    f"✅ {uploaded_file.name} "
                    "processed successfully"
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
                        "Processing Time",
                        f"{processing_time:.2f}s",
                    )

                # =================================================
                # DOWNLOAD OUTPUT
                # =================================================

                st.download_button(
                    label=(
                        f"⬇️ Download "
                        f"{output_name}"
                    ),
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    key=(
                        f"output_{index}_"
                        f"{uploaded_file.name}"
                    ),
                    use_container_width=True,
                )

                # =================================================
                # PREVIEW
                # =================================================

                if show_preview:

                    with st.expander(
                        f"Preview — "
                        f"{uploaded_file.name}"
                    ):

                        if output_rows > 0:

                            st.dataframe(
                                result_df.head(100),
                                use_container_width=True,
                            )

                        else:

                            st.info(
                                "No output rows."
                            )

                # =================================================
                # SUMMARY
                # =================================================

                st.subheader(
                    f"📊 QA Summary — "
                    f"{uploaded_file.name}"
                )

                if len(summary_df) > 0:

                    st.dataframe(
                        summary_df,
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
                            "⬇️ Download QA Summary"
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
                        "QA filtering."
                    )

                # ------------------------------------------------
                # Release references
                # ------------------------------------------------

                del result_df
                del summary_df
                del output_bytes

            except Exception as exc:

                failed += 1

                st.error(
                    f"❌ Failed to process "
                    f"'{uploaded_file.name}'"
                )

                st.exception(exc)

            progress.progress(
                index / total
            )

        # ====================================================
        # FINAL RESULT
        # ====================================================

        status.empty()

        if failed == 0:

            st.success(
                f"🎉 All {successful} "
                "file(s) processed successfully."
            )

        else:

            st.warning(
                f"Finished: "
                f"{successful} successful, "
                f"{failed} failed."
            )

else:

    st.info(
        "Please upload one or more .xlsx "
        "files to start."
    )
