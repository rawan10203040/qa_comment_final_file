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
    "add Area / Group / Team leader from reference.csv, "
    "remove Analysis and OK rows, and generate a summary."
)


# ============================================================
# LOAD REFERENCE CSV
# ============================================================

@st.cache_data
def load_reference():

    reference = pl.read_csv(
        "reference.csv",
        infer_schema_length=None
    )

    # Clean column names
    reference = reference.rename({
        col: col.strip()
        for col in reference.columns
    })

    required = [
        "FunctionName",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ]

    missing = [
        col
        for col in required
        if col not in reference.columns
    ]

    if missing:
        raise ValueError(
            "Missing columns in reference.csv: "
            + ", ".join(missing)
        )

    # Convert everything used in mapping to strings
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
        .str.strip_chars(),
    ])

    # Only keep mapping columns
    reference = reference.select([
        "_lookup",
        "Action",
        "Area",
        "Group",
        "Team leader"
    ])

    # Remove empty FunctionName
    reference = reference.filter(
        pl.col("_lookup") != ""
    )

    # Make sure one FunctionName has one mapping
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

except Exception as e:

    st.error("❌ Error loading reference.csv")

    st.code(str(e))

    st.stop()


# ============================================================
# SHOW REFERENCE STATUS
# ============================================================

with st.expander("📚 Reference Status"):

    st.write(
        f"Reference rows: **{reference_df.height:,}**"
    )

    st.write(
        "Reference columns:"
    )

    st.write(reference_df.columns)


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

    summary_rows = []

    st.info(
        f"📁 {len(uploaded_files)} file(s) uploaded."
    )


    # ========================================================
    # EACH FILE
    # ========================================================

    for uploaded_file in uploaded_files:

        filename = uploaded_file.name

        st.markdown(
            f"### 🔄 Processing: `{filename}`"
        )

        try:

            # ==================================================
            # READ EXCEL
            # ==================================================

            data = uploaded_file.getvalue()

            df = pl.read_excel(
                io.BytesIO(data),
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

            required_input = [
                "FunctionName",
                "IsMultiValue",
                "HasBlankValue",
                "QAComment"
            ]

            missing_input = [
                col
                for col in required_input
                if col not in df.columns
            ]

            if missing_input:

                st.error(
                    f"❌ `{filename}` skipped. "
                    f"Missing columns: "
                    f"{', '.join(missing_input)}"
                )

                continue


            # ==================================================
            # ORIGINAL ROWS
            # ==================================================

            original_rows = df.height


            # ==================================================
            # CLEAN INPUT COLUMNS
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

            fn = (
                pl.col("FunctionName")
                .str.to_lowercase()
            )

            multi = pl.col("IsMultiValue")

            blank = pl.col("HasBlankValue")


            df = df.with_columns(

                pl.when(
                    (fn == "packing")
                    & (multi == "TRUE")
                    & (blank == "FALSE")
                )
                .then(pl.lit("ok"))

                .when(
                    (fn == "packing")
                    & (multi == "TRUE")
                    & (blank == "TRUE")
                )
                .then(pl.lit("null"))

                .when(
                    (fn == "packing")
                    & (multi == "FALSE")
                    & (blank == "FALSE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (fn == "packing")
                    & (multi == "FALSE")
                    & (blank == "TRUE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (fn != "packing")
                    & (multi == "TRUE")
                    & (blank == "FALSE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (fn != "packing")
                    & (multi == "TRUE")
                    & (blank == "TRUE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (fn != "packing")
                    & (multi == "FALSE")
                    & (blank == "FALSE")
                )
                .then(pl.lit("ok"))

                .when(
                    (fn != "packing")
                    & (multi == "FALSE")
                    & (blank == "TRUE")
                )
                .then(pl.lit("null"))

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
            # REMOVE OLD REFERENCE COLUMNS IF THEY EXIST
            # ==================================================

            columns_to_remove = [
                col
                for col in [
                    "Action",
                    "Area",
                    "Group",
                    "Team leader"
                ]
                if col in df.columns
            ]

            if columns_to_remove:

                df = df.drop(
                    columns_to_remove
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

            df = df.drop("_lookup")


            # ==================================================
            # FILL REFERENCE NULLS
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

            preferred = [
                "FunctionName",
                "Area",
                "Group",
                "Team leader"
            ]

            remaining = [
                col
                for col in df.columns
                if col not in preferred
            ]

            df = df.select(
                preferred + remaining
            )


            # ==================================================
            # REMOVE ANALYSIS
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
            # REMOVE OK
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
            # SUMMARY
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
            # OUTPUT FILE
            # ==================================================

            output_name = (
                filename.rsplit(".", 1)[0]
                + "_Updated.xlsx"
            )


            # ==================================================
            # WRITE EXCEL
            # ==================================================

            output = io.BytesIO()

            df.write_excel(
                output
            )

            output.seek(0)


            # ==================================================
            # STORE OUTPUT
            # ==================================================

            generated_files.append({

                "name":
                    output_name,

                "data":
                    output.getvalue()
            })


            # ==================================================
            # SUCCESS
            # ==================================================

            st.success(
                f"✅ `{output_name}` created successfully."
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
    # SHOW SUMMARY
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

        file = generated_files[0]

        st.download_button(

            label="⬇️ Download Updated Excel",

            data=file["data"],

            file_name=file["name"],

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
        ) as z:

            for file in generated_files:

                z.writestr(
                    file["name"],
                    file["data"]
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


        for i, file in enumerate(
            generated_files
        ):

            st.download_button(

                label=f"⬇️ {file['name']}",

                data=file["data"],

                file_name=file["name"],

                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),

                key=f"download_{i}"
            )


    else:

        st.warning(
            "⚠️ No output files were generated."
        )


    # ========================================================
    # FINISHED
    # ========================================================

    if generated_files:

        st.success(
            "🎉 Finished Successfully!"
        )
