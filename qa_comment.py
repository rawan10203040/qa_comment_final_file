import streamlit as st
import pandas as pd
import io
import zipfile
from pathlib import Path


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="QA Comment Updater",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("📊 QA Comment Updater")

st.write(
    "Upload one or more Excel files to update QAComment, "
    "add Area / Group / Team leader from reference.xlsx, "
    "remove Analysis and OK rows, and generate a summary."
)


# ============================================================
# FILE PATH
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
REFERENCE_FILE = BASE_DIR / "reference.xlsx"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_column_name(column):
    """
    Clean Excel column names.
    """
    return str(column).strip()


def normalize_text(value):
    """
    Convert value to clean lowercase text.
    """
    if pd.isna(value):
        return ""

    return str(value).strip().lower()


def normalize_header(value):
    """
    Normalize header for comparison.
    """
    return (
        str(value)
        .strip()
        .lower()
        .replace("_", "")
        .replace(" ", "")
    )


# ============================================================
# LOAD REFERENCE
# ============================================================

@st.cache_data
def load_reference(reference_path):
    """
    Read reference.xlsx using pandas/openpyxl.
    """

    if not reference_path.exists():
        raise FileNotFoundError(
            f"Reference file not found:\n{reference_path}"
        )

    # --------------------------------------------------------
    # Read first sheet
    # --------------------------------------------------------

    reference = pd.read_excel(
        reference_path,
        engine="openpyxl"
    )

    # --------------------------------------------------------
    # Clean headers
    # --------------------------------------------------------

    reference.columns = [
        clean_column_name(col)
        for col in reference.columns
    ]

    # --------------------------------------------------------
    # Case-insensitive header mapping
    # --------------------------------------------------------

    header_mapping = {}

    for col in reference.columns:

        normalized = normalize_header(col)

        if normalized == "functionname":
            header_mapping[col] = "FunctionName"

        elif normalized == "action":
            header_mapping[col] = "Action"

        elif normalized == "area":
            header_mapping[col] = "Area"

        elif normalized == "group":
            header_mapping[col] = "Group"

        elif normalized in [
            "teamleader",
            "team_leader"
        ]:
            header_mapping[col] = "Team leader"

    reference = reference.rename(
        columns=header_mapping
    )

    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

    required_columns = [
        "FunctionName",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ]

    missing_columns = [
        col
        for col in required_columns
        if col not in reference.columns
    ]

    if missing_columns:
        raise ValueError(
            "reference.xlsx is missing these columns:\n"
            + "\n".join(
                f"- {col}"
                for col in missing_columns
            )
            + "\n\nColumns found in reference.xlsx:\n"
            + "\n".join(
                f"- {col}"
                for col in reference.columns
            )
        )

    # --------------------------------------------------------
    # Keep only required columns
    # --------------------------------------------------------

    reference = reference[
        required_columns
    ].copy()

    # --------------------------------------------------------
    # Clean values
    # --------------------------------------------------------

    for col in required_columns:

        reference[col] = (
            reference[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Create lookup key
    # --------------------------------------------------------

    reference["_lookup"] = (
        reference["FunctionName"]
        .str.strip()
        .str.lower()
    )

    # --------------------------------------------------------
    # Remove empty FunctionName
    # --------------------------------------------------------

    reference = reference[
        reference["_lookup"] != ""
    ].copy()

    # --------------------------------------------------------
    # Remove duplicate FunctionName
    # Keep first occurrence
    # --------------------------------------------------------

    reference = reference.drop_duplicates(
        subset=["_lookup"],
        keep="first"
    )

    # --------------------------------------------------------
    # Replace empty reference values
    # --------------------------------------------------------

    for col in [
        "Action",
        "Area",
        "Group",
        "Team leader"
    ]:

        reference[col] = reference[col].replace(
            "",
            "-"
        )

    return reference


# ============================================================
# LOAD REFERENCE
# ============================================================

try:

    reference_df = load_reference(
        REFERENCE_FILE
    )

    st.success(
        f"✅ Reference loaded successfully: "
        f"{len(reference_df):,} rows"
    )

except Exception as e:

    st.error(
        "❌ Cannot read reference.xlsx"
    )

    st.code(
        str(e)
    )

    st.info(
        "Make sure reference.xlsx is in the same GitHub "
        "folder as qa_comment.py."
    )

    st.stop()


# ============================================================
# REFERENCE INFORMATION
# ============================================================

with st.expander("📚 Reference Information"):

    st.write(
        f"**Reference file:** `{REFERENCE_FILE.name}`"
    )

    st.write(
        "**Reference columns:**"
    )

    st.write([
        "FunctionName",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ])

    preview_reference = reference_df[
        [
            "FunctionName",
            "Action",
            "Area",
            "Group",
            "Team leader"
        ]
    ].head(20)

    st.dataframe(
        preview_reference,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_files = st.file_uploader(
    "📤 Upload Excel file(s)",
    type=["xlsx"],
    accept_multiple_files=True
)


# ============================================================
# PROCESS FILES
# ============================================================

if uploaded_files:

    generated_files = []

    summary_rows = []

    st.info(
        f"📁 {len(uploaded_files)} file(s) uploaded."
    )

    # ========================================================
    # PROCESS EACH FILE
    # ========================================================

    for uploaded_file in uploaded_files:

        filename = uploaded_file.name

        st.markdown(
            f"### 🔄 Processing: `{filename}`"
        )

        try:

            # ==================================================
            # READ INPUT EXCEL
            # ==================================================

            file_bytes = uploaded_file.getvalue()

            df = pd.read_excel(
                io.BytesIO(file_bytes),
                engine="openpyxl"
            )

            # ==================================================
            # CLEAN COLUMN NAMES
            # ==================================================

            df.columns = [
                clean_column_name(col)
                for col in df.columns
            ]

            # ==================================================
            # NORMALIZE INPUT HEADERS
            # ==================================================

            input_header_mapping = {}

            for col in df.columns:

                normalized = normalize_header(col)

                if normalized == "functionname":
                    input_header_mapping[col] = "FunctionName"

                elif normalized == "ismultivalue":
                    input_header_mapping[col] = "IsMultiValue"

                elif normalized == "hasblankvalue":
                    input_header_mapping[col] = "HasBlankValue"

                elif normalized == "qacomment":
                    input_header_mapping[col] = "QAComment"

                elif normalized == "action":
                    input_header_mapping[col] = "Action"

                elif normalized == "area":
                    input_header_mapping[col] = "Area"

                elif normalized == "group":
                    input_header_mapping[col] = "Group"

                elif normalized in [
                    "teamleader",
                    "team_leader"
                ]:
                    input_header_mapping[col] = "Team leader"

            df = df.rename(
                columns=input_header_mapping
            )

            # ==================================================
            # REQUIRED INPUT COLUMNS
            # ==================================================

            required_input_columns = [
                "FunctionName",
                "IsMultiValue",
                "HasBlankValue",
                "QAComment"
            ]

            missing_input_columns = [
                col
                for col in required_input_columns
                if col not in df.columns
            ]

            if missing_input_columns:

                st.error(
                    f"❌ `{filename}` skipped.\n\n"
                    f"Missing columns: "
                    f"{', '.join(missing_input_columns)}"
                )

                continue

            # ==================================================
            # ORIGINAL ROW COUNT
            # ==================================================

            original_rows = len(df)

            # ==================================================
            # CLEAN INPUT VALUES
            # ==================================================

            df["FunctionName"] = (
                df["FunctionName"]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df["IsMultiValue"] = (
                df["IsMultiValue"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            df["HasBlankValue"] = (
                df["HasBlankValue"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            df["QAComment"] = (
                df["QAComment"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            )

            # ==================================================
            # UPDATE QACOMMENT
            # ==================================================

            function_name = (
                df["FunctionName"]
                .str.lower()
                .str.strip()
            )

            is_packing = (
                function_name == "packing"
            )

            is_multi = df["IsMultiValue"]
            has_blank = df["HasBlankValue"]

            # --------------------------------------------------
            # PACKING
            # --------------------------------------------------

            # TRUE + FALSE = OK

            condition = (
                is_packing
                & (is_multi == "TRUE")
                & (has_blank == "FALSE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "ok"

            # TRUE + TRUE = NULL

            condition = (
                is_packing
                & (is_multi == "TRUE")
                & (has_blank == "TRUE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "null"

            # FALSE + FALSE = CONFLICT

            condition = (
                is_packing
                & (is_multi == "FALSE")
                & (has_blank == "FALSE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "conflict"

            # FALSE + TRUE = CONFLICT

            condition = (
                is_packing
                & (is_multi == "FALSE")
                & (has_blank == "TRUE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "conflict"

            # --------------------------------------------------
            # OTHER FUNCTIONS
            # --------------------------------------------------

            not_packing = (
                function_name != "packing"
            )

            # TRUE + FALSE = CONFLICT

            condition = (
                not_packing
                & (is_multi == "TRUE")
                & (has_blank == "FALSE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "conflict"

            # TRUE + TRUE = CONFLICT

            condition = (
                not_packing
                & (is_multi == "TRUE")
                & (has_blank == "TRUE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "conflict"

            # FALSE + FALSE = OK

            condition = (
                not_packing
                & (is_multi == "FALSE")
                & (has_blank == "FALSE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "ok"

            # FALSE + TRUE = NULL

            condition = (
                not_packing
                & (is_multi == "FALSE")
                & (has_blank == "TRUE")
            )

            df.loc[
                condition,
                "QAComment"
            ] = "null"

            # ==================================================
            # CREATE LOOKUP KEY
            # ==================================================

            df["_lookup"] = (
                df["FunctionName"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            )

            # ==================================================
            # REMOVE OLD REFERENCE COLUMNS
            # ==================================================

            old_reference_columns = [
                "Action",
                "Area",
                "Group",
                "Team leader"
            ]

            columns_to_drop = [
                col
                for col in old_reference_columns
                if col in df.columns
            ]

            if columns_to_drop:

                df = df.drop(
                    columns=columns_to_drop
                )

            # ==================================================
            # PREPARE REFERENCE FOR MERGE
            # ==================================================

            reference_for_merge = reference_df[
                [
                    "_lookup",
                    "Action",
                    "Area",
                    "Group",
                    "Team leader"
                ]
            ].copy()

            # ==================================================
            # MERGE INPUT WITH REFERENCE
            # ==================================================

            df = df.merge(
                reference_for_merge,
                on="_lookup",
                how="left"
            )

            # ==================================================
            # REMOVE LOOKUP COLUMN
            # ==================================================

            df = df.drop(
                columns=["_lookup"]
            )

            # ==================================================
            # FILL MISSING REFERENCE VALUES
            # ==================================================

            for col in [
                "Action",
                "Area",
                "Group",
                "Team leader"
            ]:

                if col not in df.columns:

                    df[col] = "-"

                else:

                    df[col] = (
                        df[col]
                        .fillna("-")
                        .astype(str)
                        .str.strip()
                    )

                    df.loc[
                        df[col].isin(
                            ["", "nan", "None"]
                        ),
                        col
                    ] = "-"

            # ==================================================
            # REORDER COLUMNS
            # ==================================================

            first_columns = [
                "FunctionName",
                "Area",
                "Group",
                "Team leader"
            ]

            remaining_columns = [
                col
                for col in df.columns
                if col not in first_columns
            ]

            df = df[
                first_columns
                + remaining_columns
            ]

            # ==================================================
            # REMOVE ACTION = ANALYSIS
            # ==================================================

            before_analysis = len(df)

            action_normalized = (
                df["Action"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            )

            df = df[
                action_normalized != "analysis"
            ].copy()

            removed_analysis = (
                before_analysis
                - len(df)
            )

            # ==================================================
            # REMOVE QACOMMENT = OK
            # ==================================================

            before_ok = len(df)

            qa_normalized = (
                df["QAComment"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            )

            df = df[
                qa_normalized != "ok"
            ].copy()

            removed_ok = (
                before_ok
                - len(df)
            )

            # ==================================================
            # BUILD SUMMARY
            # ==================================================

            if len(df) > 0:

                summary_df = df[
                    df["QAComment"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    .isin(
                        [
                            "conflict",
                            "null"
                        ]
                    )
                ].copy()

                if len(summary_df) > 0:

                    temp_summary = (
                        summary_df
                        .groupby(
                            [
                                "Area",
                                "Team leader"
                            ],
                            dropna=False
                        )
                        .size()
                        .reset_index(
                            name="Counts"
                        )
                    )

                    for _, row in temp_summary.iterrows():

                        summary_rows.append(
                            {
                                "Area": str(
                                    row["Area"]
                                ),
                                "Counts": int(
                                    row["Counts"]
                                ),
                                "Team Leader": str(
                                    row["Team leader"]
                                )
                            }
                        )

            # ==================================================
            # OUTPUT FILE NAME
            # ==================================================

            if filename.lower().endswith(".xlsx"):

                output_name = (
                    filename[:-5]
                    + "_Updated.xlsx"
                )

            else:

                output_name = (
                    filename
                    + "_Updated.xlsx"
                )

            # ==================================================
            # WRITE OUTPUT EXCEL
            # ==================================================

            output_buffer = io.BytesIO()

            with pd.ExcelWriter(
                output_buffer,
                engine="xlsxwriter"
            ) as writer:

                df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Output"
                )

            output_buffer.seek(0)

            # ==================================================
            # STORE OUTPUT
            # ==================================================

            generated_files.append(
                {
                    "name": output_name,
                    "data": output_buffer.getvalue()
                }
            )

            # ==================================================
            # SUCCESS
            # ==================================================

            st.success(
                f"✅ Created `{output_name}`"
            )

            # ==================================================
            # STATISTICS
            # ==================================================

            c1, c2, c3, c4 = st.columns(4)

            c1.metric(
                "Original Rows",
                f"{original_rows:,}"
            )

            c2.metric(
                "Analysis Removed",
                f"{removed_analysis:,}"
            )

            c3.metric(
                "OK Removed",
                f"{removed_ok:,}"
            )

            c4.metric(
                "Final Rows",
                f"{len(df):,}"
            )

            # ==================================================
            # PREVIEW
            # ==================================================

            with st.expander(
                f"👀 Preview - {output_name}"
            ):

                st.dataframe(
                    df.head(100),
                    use_container_width=True,
                    hide_index=True
                )

        # ======================================================
        # ERROR FOR CURRENT FILE
        # ======================================================

        except Exception as e:

            st.error(
                f"❌ Error processing `{filename}`"
            )

            st.exception(e)


    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    st.divider()

    st.header(
        "📋 FINAL SUMMARY"
    )

    if summary_rows:

        final_summary = pd.DataFrame(
            summary_rows
        )

        final_summary = (
            final_summary
            .groupby(
                [
                    "Area",
                    "Team Leader"
                ],
                as_index=False
            )["Counts"]
            .sum()
        )

        final_summary = final_summary[
            [
                "Area",
                "Counts",
                "Team Leader"
            ]
        ]

        final_summary = (
            final_summary
            .sort_values(
                "Counts",
                ascending=False
            )
            .reset_index(drop=True)
        )

    else:

        final_summary = pd.DataFrame(
            columns=[
                "Area",
                "Counts",
                "Team Leader"
            ]
        )

    # ========================================================
    # DISPLAY SUMMARY
    # ========================================================

    if len(final_summary) > 0:

        st.dataframe(
            final_summary,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "ℹ️ No Conflict or Null rows were found."
        )

    # ========================================================
    # DOWNLOAD SUMMARY
    # ========================================================

    summary_buffer = io.BytesIO()

    with pd.ExcelWriter(
        summary_buffer,
        engine="xlsxwriter"
    ) as writer:

        final_summary.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

    summary_buffer.seek(0)

    st.download_button(
        label="📥 Download Summary Excel",
        data=summary_buffer.getvalue(),
        file_name="QA_Summary.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        key="download_summary"
    )

    # ========================================================
    # DOWNLOAD OUTPUT FILES
    # ========================================================

    st.divider()

    st.header(
        "📥 DOWNLOAD OUTPUT FILES"
    )

    # --------------------------------------------------------
    # One file
    # --------------------------------------------------------

    if len(generated_files) == 1:

        file_info = generated_files[0]

        st.download_button(
            label="⬇️ Download Updated Excel",
            data=file_info["data"],
            file_name=file_info["name"],
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            key="download_single"
        )

    # --------------------------------------------------------
    # Multiple files
    # --------------------------------------------------------

    elif len(generated_files) > 1:

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_file:

            for file_info in generated_files:

                zip_file.writestr(
                    file_info["name"],
                    file_info["data"]
                )

        zip_buffer.seek(0)

        st.download_button(
            label="📦 Download All Files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Updated_Output_Files.zip",
            mime="application/zip",
            key="download_zip"
        )

        st.write(
            "### Individual Files"
        )

        for index, file_info in enumerate(
            generated_files
        ):

            st.download_button(
                label=(
                    f"⬇️ {file_info['name']}"
                ),
                data=file_info["data"],
                file_name=file_info["name"],
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                key=f"download_file_{index}"
            )

    # --------------------------------------------------------
    # No files
    # --------------------------------------------------------

    else:

        st.warning(
            "⚠️ No files were generated."
        )

    # ========================================================
    # FINISHED
    # ========================================================

    if generated_files:

        st.success(
            "🎉 Finished Successfully!"
        )
