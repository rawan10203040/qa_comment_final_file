# ============================================================
# STREAMLIT APP
# QA COMMENT UPDATER + REFERENCE MAPPING + SUMMARY
# ============================================================

import streamlit as st
import polars as pl
import io
import zipfile


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
    "add Area / Group / Team leader, remove Analysis and OK rows, "
    "and generate a consolidated summary."
)


# ============================================================
# LOAD REFERENCE.CSV
# ============================================================

@st.cache_data
def load_reference():

    reference_df = pl.read_csv(
        "reference.csv",
        infer_schema_length=None,
        ignore_errors=False
    )

    required_reference_columns = [
        "FunctionName",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ]

    missing_reference_columns = [
        c
        for c in required_reference_columns
        if c not in reference_df.columns
    ]

    if missing_reference_columns:

        raise ValueError(
            "reference.csv is missing required columns: "
            + ", ".join(missing_reference_columns)
        )

    # --------------------------------------------------------
    # Clean Reference
    # --------------------------------------------------------

    reference_df = reference_df.with_columns([

        pl.col("FunctionName")
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars()
        .str.to_lowercase()
        .alias("FunctionName_lookup"),

        pl.col("Action")
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("Area")
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("Group")
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("Team leader")
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars()
    ])

    # --------------------------------------------------------
    # Keep only required mapping columns
    # --------------------------------------------------------

    mapping_df = reference_df.select([
        "FunctionName_lookup",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ])

    # --------------------------------------------------------
    # Remove duplicate FunctionName mappings
    #
    # This prevents duplicate rows after JOIN if the same
    # FunctionName appears more than once in reference.csv.
    # --------------------------------------------------------

    mapping_df = mapping_df.unique(
        subset=["FunctionName_lookup"],
        keep="first"
    )

    return mapping_df


# ============================================================
# LOAD REFERENCE
# ============================================================

try:

    mapping_df = load_reference()

except Exception as e:

    st.error(
        "❌ Could not load reference.csv"
    )

    st.exception(e)

    st.stop()


# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_files = st.file_uploader(
    "📤 Upload Excel file(s)",
    type=["xlsx"],
    accept_multiple_files=True
)


# ============================================================
# PROCESS
# ============================================================

if uploaded_files:

    generated_files = []

    all_summary_rows = []

    st.info(
        f"{len(uploaded_files)} file(s) uploaded."
    )


    # ========================================================
    # PROCESS EACH FILE
    # ========================================================

    for uploaded_file in uploaded_files:

        filename = uploaded_file.name

        st.write(
            f"### 🔄 Processing: `{filename}`"
        )

        try:

            # ==================================================
            # READ EXCEL
            # ==================================================

            file_bytes = uploaded_file.getvalue()

            df = pl.read_excel(
                io.BytesIO(file_bytes),
                infer_schema_length=None
            )

            # ==================================================
            # REQUIRED COLUMNS
            # ==================================================

            required_columns = [
                "FunctionName",
                "IsMultiValue",
                "HasBlankValue",
                "QAComment"
            ]

            missing_columns = [
                c
                for c in required_columns
                if c not in df.columns
            ]

            if missing_columns:

                st.error(
                    f"❌ Skipping `{filename}`. "
                    f"Missing columns: "
                    f"{', '.join(missing_columns)}"
                )

                continue


            # ==================================================
            # ORIGINAL ROW COUNT
            # ==================================================

            original_count = df.height


            # ==================================================
            # CLEAN INPUT DATA
            # ==================================================

            df = df.with_columns([

                pl.col("FunctionName")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars(),

                pl.col("IsMultiValue")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_uppercase(),

                pl.col("HasBlankValue")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_uppercase(),

                pl.col("QAComment")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_lowercase()
            ])


            # ==================================================
            # UPDATE QACOMMENT
            # ==================================================

            function_name = (
                pl.col("FunctionName")
                .str.to_lowercase()
            )

            df = df.with_columns(

                # ------------------------------------------------
                # PACKING
                # TRUE + FALSE = OK
                # ------------------------------------------------

                pl.when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "TRUE")
                    &
                    (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("ok"))

                # ------------------------------------------------
                # PACKING
                # TRUE + TRUE = NULL
                # ------------------------------------------------

                .when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "TRUE")
                    &
                    (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("null"))

                # ------------------------------------------------
                # PACKING
                # FALSE + FALSE = CONFLICT
                # ------------------------------------------------

                .when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "FALSE")
                    &
                    (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("conflict"))

                # ------------------------------------------------
                # PACKING
                # FALSE + TRUE = CONFLICT
                # ------------------------------------------------

                .when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "FALSE")
                    &
                    (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("conflict"))

                # =================================================
                # OTHER FUNCTIONS
                # =================================================

                # TRUE + FALSE = CONFLICT

                .when(
                    (function_name != "packing")
                    &
                    (pl.col("IsMultiValue") == "TRUE")
                    &
                    (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("conflict"))

                # TRUE + TRUE = CONFLICT

                .when(
                    (function_name != "packing")
                    &
                    (pl.col("IsMultiValue") == "TRUE")
                    &
                    (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("conflict"))

                # FALSE + FALSE = OK

                .when(
                    (function_name != "packing")
                    &
                    (pl.col("IsMultiValue") == "FALSE")
                    &
                    (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("ok"))

                # FALSE + TRUE = NULL

                .when(
                    (function_name != "packing")
                    &
                    (pl.col("IsMultiValue") == "FALSE")
                    &
                    (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("null"))

                # KEEP ORIGINAL

                .otherwise(
                    pl.col("QAComment")
                )

                .alias("QAComment")
            )


            # ==================================================
            # CREATE LOOKUP COLUMN
            # ==================================================

            df = df.with_columns(

                pl.col("FunctionName")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_lowercase()
                .alias("FunctionName_lookup")
            )


            # ==================================================
            # JOIN WITH REFERENCE
            # ==================================================

            df = df.join(
                mapping_df,
                on="FunctionName_lookup",
                how="left"
            )


            # ==================================================
            # REMOVE TEMP COLUMN
            # ==================================================

            df = df.drop(
                "FunctionName_lookup"
            )


            # ==================================================
            # FILL MISSING REFERENCE VALUES
            # ==================================================

            df = df.with_columns([

                pl.col("Action")
                .cast(pl.Utf8, strict=False)
                .fill_null("-"),

                pl.col("Area")
                .cast(pl.Utf8, strict=False)
                .fill_null("-"),

                pl.col("Group")
                .cast(pl.Utf8, strict=False)
                .fill_null("-"),

                pl.col("Team leader")
                .cast(pl.Utf8, strict=False)
                .fill_null("-")
            ])


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
                c
                for c in df.columns
                if c not in first_columns
            ]

            df = df.select(
                first_columns + remaining_columns
            )


            # ==================================================
            # REMOVE ACTION = ANALYSIS
            # ==================================================

            before_analysis = df.height

            df = df.filter(

                pl.col("Action")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_lowercase()
                != "analysis"
            )

            removed_analysis = (
                before_analysis - df.height
            )


            # ==================================================
            # REMOVE QACOMMENT = OK
            # ==================================================

            before_ok = df.height

            df = df.filter(

                pl.col("QAComment")
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_lowercase()
                != "ok"
            )

            removed_ok = (
                before_ok - df.height
            )


            # ==================================================
            # SUMMARY DATA
            # ==================================================

            if df.height > 0:

                summary_df = (

                    df

                    .filter(
                        pl.col("QAComment")
                        .cast(pl.Utf8, strict=False)
                        .fill_null("")
                        .str.to_lowercase()
                        .is_in([
                            "conflict",
                            "null"
                        ])
                    )

                    .group_by([
                        "Area",
                        "Team leader"
                    ])

                    .agg(
                        pl.len().alias("Counts")
                    )
                )


                for row in summary_df.iter_rows(
                    named=True
                ):

                    all_summary_rows.append({

                        "Area":
                            row["Area"],

                        "Counts":
                            row["Counts"],

                        "Team Leader":
                            row["Team leader"]
                    })


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
            # WRITE EXCEL TO MEMORY
            # ==================================================

            output_buffer = io.BytesIO()

            df.write_excel(
                output_buffer
            )

            output_buffer.seek(0)

            output_bytes = (
                output_buffer.getvalue()
            )


            # ==================================================
            # SAVE OUTPUT IN MEMORY
            # ==================================================

            generated_files.append({

                "name":
                    output_name,

                "data":
                    output_bytes
            })


            # ==================================================
            # SUCCESS
            # ==================================================

            st.success(
                f"✅ Created: `{output_name}`"
            )


            # ==================================================
            # STATISTICS
            # ==================================================

            col1, col2, col3, col4 = st.columns(4)

            with col1:

                st.metric(
                    "Original Rows",
                    original_count
                )

            with col2:

                st.metric(
                    "Analysis Removed",
                    removed_analysis
                )

            with col3:

                st.metric(
                    "OK Removed",
                    removed_ok
                )

            with col4:

                st.metric(
                    "Final Rows",
                    df.height
                )


            # ==================================================
            # PREVIEW
            # ==================================================

            with st.expander(
                f"👀 Preview - {output_name}"
            ):

                st.dataframe(
                    df.head(100).to_pandas(),
                    use_container_width=True,
                    hide_index=True
                )


        except Exception as e:

            st.error(
                f"❌ Error processing `{filename}`: {str(e)}"
            )

            with st.expander(
                "🔎 Show technical error"
            ):

                st.exception(e)


    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    st.divider()

    st.header(
        "📋 FINAL SUMMARY"
    )


    if all_summary_rows:

        final_summary = pl.DataFrame(
            all_summary_rows,
            orient="row"
        )


        # ----------------------------------------------------
        # Combine same Area + Team Leader
        # ----------------------------------------------------

        final_summary = (

            final_summary

            .group_by([
                "Area",
                "Team Leader"
            ])

            .agg(
                pl.col("Counts").sum()
            )

            .sort(
                "Counts",
                descending=True
            )
        )


        final_summary = final_summary.select([

            "Area",
            "Counts",
            "Team Leader"
        ])


    else:

        final_summary = pl.DataFrame({

            "Area":
                pl.Series([], dtype=pl.Utf8),

            "Counts":
                pl.Series([], dtype=pl.Int64),

            "Team Leader":
                pl.Series([], dtype=pl.Utf8)
        })


    # ========================================================
    # DISPLAY SUMMARY
    # ========================================================

    st.dataframe(
        final_summary.to_pandas(),
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # SUMMARY EXCEL
    # ========================================================

    summary_buffer = io.BytesIO()

    final_summary.write_excel(
        summary_buffer
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

            key="download_single_file"
        )


    elif len(generated_files) > 1:

        # ====================================================
        # CREATE ZIP
        # ====================================================

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


        # ====================================================
        # DOWNLOAD ZIP
        # ====================================================

        st.download_button(

            label="📦 Download All Files (ZIP)",

            data=zip_buffer.getvalue(),

            file_name="Updated_Output_Files.zip",

            mime="application/zip",

            key="download_zip"
        )


        # ====================================================
        # INDIVIDUAL FILES
        # ====================================================

        st.write(
            "### Individual Files"
        )


        for index, file_info in enumerate(
            generated_files
        ):

            st.download_button(

                label=
                    f"⬇️ {file_info['name']}",

                data=
                    file_info["data"],

                file_name=
                    file_info["name"],

                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),

                key=
                    f"download_file_{index}"
            )


    else:

        st.warning(
            "No files were generated."
        )


    # ========================================================
    # FINISHED
    # ========================================================

    if generated_files:

        st.success(
            "🎉 Finished Successfully!"
        )
