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
    "add Area / Group / Team leader from reference.xlsx, "
    "remove Analysis and OK rows, and generate a summary."
)


# ============================================================
# LOAD REFERENCE EXCEL
# ============================================================

@st.cache_data
def load_reference():

    reference_file = "reference.xlsx"

    # --------------------------------------------------------
    # Read Excel
    # --------------------------------------------------------

    reference = pl.read_excel(
        reference_file,
        infer_schema_length=None
    )

    # --------------------------------------------------------
    # Clean column names
    # --------------------------------------------------------

    reference = reference.rename({
        col: col.strip()
        for col in reference.columns
    })

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
            "reference.xlsx is missing these columns: "
            + ", ".join(missing_columns)
        )

    # --------------------------------------------------------
    # Clean reference data
    # --------------------------------------------------------

    reference = reference.with_columns([

        pl.col("FunctionName")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars()
        .str.to_lowercase()
        .alias("_lookup"),

        pl.col("Action")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("Area")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("Group")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars(),

        pl.col("Team leader")
        .cast(pl.String, strict=False)
        .fill_null("")
        .str.strip_chars()
    ])

    # --------------------------------------------------------
    # Keep only mapping columns
    # --------------------------------------------------------

    reference = reference.select([
        "_lookup",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ])

    # --------------------------------------------------------
    # Remove empty FunctionName
    # --------------------------------------------------------

    reference = reference.filter(
        pl.col("_lookup") != ""
    )

    # --------------------------------------------------------
    # Remove duplicate FunctionName
    # --------------------------------------------------------

    reference = reference.unique(
        subset=["_lookup"],
        keep="first"
    )

    return reference


# ============================================================
# LOAD REFERENCE
# ============================================================

try:

    reference_df = load_reference()

    st.success(
        f"✅ Reference loaded successfully: "
        f"{reference_df.height:,} rows"
    )

except Exception as e:

    st.error(
        "❌ Cannot read reference.xlsx"
    )

    st.code(str(e))

    st.stop()


# ============================================================
# REFERENCE PREVIEW
# ============================================================

with st.expander("📚 Reference Information"):

    st.write(
        "Reference columns:"
    )

    st.write(
        reference_df.columns
    )

    st.dataframe(
        reference_df.head(20).to_pandas(),
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

            df = pl.read_excel(
                io.BytesIO(file_bytes),
                infer_schema_length=None
            )


            # ==================================================
            # CLEAN COLUMN NAMES
            # ==================================================

            df = df.rename({
                col: col.strip()
                for col in df.columns
            })


            # ==================================================
            # REQUIRED INPUT COLUMNS
            # ==================================================

            required_columns = [
                "FunctionName",
                "IsMultiValue",
                "HasBlankValue",
                "QAComment"
            ]

            missing_columns = [
                col
                for col in required_columns
                if col not in df.columns
            ]

            if missing_columns:

                st.error(
                    f"❌ `{filename}` skipped. "
                    f"Missing columns: "
                    f"{', '.join(missing_columns)}"
                )

                continue


            # ==================================================
            # ORIGINAL ROWS
            # ==================================================

            original_rows = df.height


            # ==================================================
            # CLEAN INPUT DATA
            # ==================================================

            df = df.with_columns([

                pl.col("FunctionName")
                .cast(pl.String, strict=False)
                .fill_null("")
                .str.strip_chars(),

                pl.col("IsMultiValue")
                .cast(pl.String, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_uppercase(),

                pl.col("HasBlankValue")
                .cast(pl.String, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_uppercase(),

                pl.col("QAComment")
                .cast(pl.String, strict=False)
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

            is_multi = pl.col("IsMultiValue")

            has_blank = pl.col("HasBlankValue")


            df = df.with_columns(

                # PACKING
                # TRUE + FALSE = OK

                pl.when(
                    (function_name == "packing")
                    & (is_multi == "TRUE")
                    & (has_blank == "FALSE")
                )
                .then(pl.lit("ok"))

                # PACKING
                # TRUE + TRUE = NULL

                .when(
                    (function_name == "packing")
                    & (is_multi == "TRUE")
                    & (has_blank == "TRUE")
                )
                .then(pl.lit("null"))

                # PACKING
                # FALSE + FALSE = CONFLICT

                .when(
                    (function_name == "packing")
                    & (is_multi == "FALSE")
                    & (has_blank == "FALSE")
                )
                .then(pl.lit("conflict"))

                # PACKING
                # FALSE + TRUE = CONFLICT

                .when(
                    (function_name == "packing")
                    & (is_multi == "FALSE")
                    & (has_blank == "TRUE")
                )
                .then(pl.lit("conflict"))

                # OTHER FUNCTIONS
                # TRUE + FALSE = CONFLICT

                .when(
                    (function_name != "packing")
                    & (is_multi == "TRUE")
                    & (has_blank == "FALSE")
                )
                .then(pl.lit("conflict"))

                # OTHER FUNCTIONS
                # TRUE + TRUE = CONFLICT

                .when(
                    (function_name != "packing")
                    & (is_multi == "TRUE")
                    & (has_blank == "TRUE")
                )
                .then(pl.lit("conflict"))

                # OTHER FUNCTIONS
                # FALSE + FALSE = OK

                .when(
                    (function_name != "packing")
                    & (is_multi == "FALSE")
                    & (has_blank == "FALSE")
                )
                .then(pl.lit("ok"))

                # OTHER FUNCTIONS
                # FALSE + TRUE = NULL

                .when(
                    (function_name != "packing")
                    & (is_multi == "FALSE")
                    & (has_blank == "TRUE")
                )
                .then(pl.lit("null"))

                # Keep original

                .otherwise(
                    pl.col("QAComment")
                )

                .alias("QAComment")
            )


            # ==================================================
            # CREATE LOOKUP
            # ==================================================

            df = df.with_columns(

                pl.col("FunctionName")
                .cast(pl.String, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_lowercase()
                .alias("_lookup")
            )


            # ==================================================
            # REMOVE OLD REFERENCE COLUMNS
            # ==================================================

            old_reference_columns = [
                col
                for col in [
                    "Action",
                    "Area",
                    "Group",
                    "Team leader"
                ]
                if col in df.columns
            ]

            if old_reference_columns:

                df = df.drop(
                    old_reference_columns
                )


            # ==================================================
            # JOIN REFERENCE
            # ==================================================

            df = df.join(
                reference_df,
                on="_lookup",
                how="left"
            )


            # ==================================================
            # REMOVE LOOKUP
            # ==================================================

            df = df.drop(
                "_lookup"
            )


            # ==================================================
            # FILL EMPTY REFERENCE VALUES
            # ==================================================

            df = df.with_columns([

                pl.col("Action")
                .fill_null("-"),

                pl.col("Area")
                .fill_null("-"),

                pl.col("Group")
                .fill_null("-"),

                pl.col("Team leader")
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
                col
                for col in df.columns
                if col not in first_columns
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
                .cast(pl.String, strict=False)
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
                .cast(pl.String, strict=False)
                .fill_null("")
                .str.strip_chars()
                .str.to_lowercase()
                != "ok"
            )

            removed_ok = (
                before_ok - df.height
            )


            # ==================================================
            # BUILD SUMMARY
            # ==================================================

            if df.height > 0:

                temp_summary = (

                    df

                    .filter(
                        pl.col("QAComment")
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


                for row in temp_summary.iter_rows(
                    named=True
                ):

                    summary_rows.append({

                        "Area":
                            str(row["Area"]),

                        "Counts":
                            int(row["Counts"]),

                        "Team Leader":
                            str(row["Team leader"])
                    })


            # ==================================================
            # OUTPUT NAME
            # ==================================================

            output_name = (
                filename.rsplit(".", 1)[0]
                + "_Updated.xlsx"
            )


            # ==================================================
            # WRITE OUTPUT EXCEL
            # ==================================================

            output_buffer = io.BytesIO()

            df.write_excel(
                output_buffer
            )

            output_buffer.seek(0)


            # ==================================================
            # STORE FILE
            # ==================================================

            generated_files.append({

                "name":
                    output_name,

                "data":
                    output_buffer.getvalue()
            })


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
                original_rows
            )

            c2.metric(
                "Analysis Removed",
                removed_analysis
            )

            c3.metric(
                "OK Removed",
                removed_ok
            )

            c4.metric(
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
                f"❌ Error processing `{filename}`"
            )

            st.code(
                str(e)
            )


    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    st.divider()

    st.header(
        "📋 FINAL SUMMARY"
    )


    if summary_rows:

        final_summary = pl.DataFrame(
            summary_rows,
            orient="row"
        )

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

            .select([
                "Area",
                "Counts",
                "Team Leader"
            ])
        )

    else:

        final_summary = pl.DataFrame({
            "Area": pl.Series(
                "Area",
                [],
                dtype=pl.String
            ),

            "Counts": pl.Series(
                "Counts",
                [],
                dtype=pl.Int64
            ),

            "Team Leader": pl.Series(
                "Team Leader",
                [],
                dtype=pl.String
            )
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
    # DOWNLOAD SUMMARY
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

            key="download_single"
        )


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
            "⚠️ No files were generated."
        )


    # ========================================================
    # FINISHED
    # ========================================================

    if generated_files:

        st.success(
            "🎉 Finished Successfully!"
        )
