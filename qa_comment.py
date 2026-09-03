import streamlit as st
import polars as pl
import pandas as pd
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
    "and generate a summary."
)


# ============================================================
# REFERENCE DATA
# ============================================================

REFERENCE_DATA = [
    ["Part Options", "analysis", "-", "-", "-"],
    ["ZPCN", "analysis", "-", "-", "-"],
    ["OBS code", "analysis", "-", "-", "-"],
    ["Tags", "analysis", "-", "-", "-"],
    ["Gidep", "analysis", "-", "-", "-"],
    ["Packing", "analysis", "-", "-", "-"],
    ["Other Qualification Features", "ok", "Qualification", "Qualification", "Mohamed Farouk"],
    ["Reach version", "analysis", "-", "-", "-"],
    ["InroductionDate", "analysis", "Inroduction Date", "Qualification", "Mohamed Farouk"],
    ["Lifecycle", "ok", "Risk", "Lifecycle", "Hamed mohamed"],
    ["Die Family", "ok", "Parts", "Parts", "Mohamed Sayed Farouk"],
    ["New Market code", "ok", "Market", "Market", "Hatem Hussein"],
    ["Forecast PLP", "ok", "Forecast", "Forecast", "Ahmed Abotaleb"],
    ["TradeCodes", "ok", "Trade Codes", "Trade Codes", "Ahmed Abotaleb"],
    ["New OBS code", "analysis", "-", "-", "-"],
    ["Exemption", "ok", "Compliance - Main", "Compliance", "Mohamed Nabil"],
    ["cross flag", "analysis", "-", "-", "-"],
    ["Rohs", "ok", "Compliance - Main", "Compliance", "Mohamed Nabil"],
    ["Part Marking Image", "ok", "Part Image", "Parts", "Abdelrahman Ata"],
    ["Reach", "ok", "Compliance - Main", "Compliance", "Mohamed Nabil"],
    ["Reliability", "ok", "Qualification", "Qualification", "Mohamed Farouk"],
    ["Halogen Free", "ok", "Compliance - Main", "Compliance", "Mohamed Nabil"],
    ["China Rohs", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["EMRT", "ok", "Part conflict Mineral", "Part conflict Mineral", "Ahmed Abotaleb"],
    ["CMRT", "ok", "Part conflict Mineral", "Part conflict Mineral", "Ahmed Abotaleb"],
    ["QualificationGrad", "ok", "Qualification", "Qualification", "Mohamed Farouk"],
    ["Automotive", "ok", "Qualification", "Qualification", "Mohamed Farouk"],
    ["Scip", "analysis", "-", "-", "-"],
    ["Chemical", "ok", "Chemical", "Compliance", "Mohamed Nabil"],
    ["Compliance level", "analysis", "Qualification .. Others", "Qualification", "Mohamed Farouk"],
    ["New Parametric", "ok", "Parametric", "Parts", "Mohamed Sayed Farouk"],
    ["Country of Origin", "ok", "SC location", "SC sites", "Mostafa Ahmed"],
    ["SiteIDs", "ok", "SC location", "SC sites", "Mostafa Ahmed"],
    ["Family", "ok", "Parts", "Parts", "Mohamed Sayed Farouk"],
    ["Part Marking", "ok", "Part Marking", "Parts", "Abdelrahman Ata"],
    ["Datasheet", "ok", "Parts", "Parts", "Mohamed Sayed Farouk"],
    ["PFAS", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["LeadFinishPlating", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["Lead Free Process Capability", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["ReflowSolderTime", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["California Proposition 65", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["NumberOfReflowCycle", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["Facility Type", "ok", "SC location", "SC sites", "Mostafa Ahmed"],
    ["NetWeight", "ok", "MFG - Others", "MFG", "Ahmed ibraheem"],
    ["Automotive date", "ok", "Qualification .. Others", "Qualification", "Mohamed Farouk"],
    ["MSL", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["MaximumReflowTemperature", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["Shelf Life Condition", "ok", "MFG - Others", "MFG", "Ahmed ibraheem"],
    ["MaximumWaveTemperature", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["WaveSolderTime", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["TSCA", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["BaseMaterial", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["Military", "ok", "Qualification", "Qualification", "Mohamed Farouk"],
    ["Shelf Life Months", "ok", "MFG - Others", "MFG", "Ahmed ibraheem"],
    ["PinOut", "ok", "PinOut", "Parts", "Abdelrahman Ata"],
    ["Underplating", "ok", "MFG - Main", "MFG", "Ahmed ibraheem"],
    ["Package", "ok", "Package", "Parts", "Abdelrahman Ata"],
    ["POPs", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["ELV Directive", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["EU REACH Restricted", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["WEEE", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["Asbestos", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["EU Batteries Regulation", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["GADSL", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["Korea RoHS", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["Ozone Layer Depletion 2009", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["Stockholm POPs", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["Lead Free", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["Rohs version", "ok", "Compliance - Others", "Compliance", "Mohamed Nabil"],
    ["original introduction date", "ok", "InroductionDate", "InroductionDate", "Mohamed Farouk"],
    ["Shelf Life Start", "ok", "MFG - Others", "MFG", "Ahmed ibraheem"],
    ["Baking Time", "ok", "MFG - Others", "MFG", "Ahmed ibraheem"],
    ["Part Mating", "ok", "MFG - Others", "MFG", "Ahmed ibraheem"],
    ["Product", "ok", "Parts", "Parts", "Mohamed Sayed Farouk"],
    ["Countrfiet", "analysis", "-", "-", "-"],
    ["ProcessTechnology", "ok", "Technology", "Technology", "Mostafa Ahmed"],
    ["DieArea", "ok", "Technology", "Technology", "Mostafa Ahmed"],
    ["WaferSize", "ok", "Technology", "Technology", "Mostafa Ahmed"],
    ["Fabrication Technology", "ok", "Technology", "Technology", "Mostafa Ahmed"],
    ["Fabrication", "ok", "SC location", "SC sites", "Mostafa Ahmed"],
]


REFERENCE_COLUMNS = [
    "FunctionName",
    "Action",
    "Area",
    "Group",
    "Team leader"
]


reference_df = pl.DataFrame(
    REFERENCE_DATA,
    schema=REFERENCE_COLUMNS
)


# ============================================================
# CREATE LOOKUP DICTIONARY
# ============================================================

reference_lookup = {
    row["FunctionName"].strip().lower(): {
        "Area": row["Area"],
        "Group": row["Group"],
        "Team leader": row["Team leader"]
    }
    for row in REFERENCE_DATA
}


# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
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

    for uploaded_file in uploaded_files:

        filename = uploaded_file.name

        st.write(
            f"### Processing: `{filename}`"
        )

        try:

            # ==================================================
            # READ EXCEL
            # ==================================================

            file_bytes = uploaded_file.getvalue()

            df = pl.read_excel(
                io.BytesIO(file_bytes)
            )

            # ==================================================
            # REQUIRED COLUMNS
            # ==================================================

            required = [
                "FunctionName",
                "IsMultiValue",
                "HasBlankValue",
                "QAComment"
            ]

            missing_columns = [
                c for c in required
                if c not in df.columns
            ]

            if missing_columns:

                st.error(
                    f"Skipping `{filename}` - "
                    f"Required columns are missing: "
                    f"{', '.join(missing_columns)}"
                )

                continue

            original_count = df.height

            # ==================================================
            # CLEAN DATA
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

            df = df.with_columns(

                pl.when(
                    (pl.col("FunctionName").str.to_lowercase() == "packing")
                    & (pl.col("IsMultiValue") == "TRUE")
                    & (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("ok"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() == "packing")
                    & (pl.col("IsMultiValue") == "TRUE")
                    & (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("null"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() == "packing")
                    & (pl.col("IsMultiValue") == "FALSE")
                    & (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() == "packing")
                    & (pl.col("IsMultiValue") == "FALSE")
                    & (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() != "packing")
                    & (pl.col("IsMultiValue") == "TRUE")
                    & (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() != "packing")
                    & (pl.col("IsMultiValue") == "TRUE")
                    & (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("conflict"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() != "packing")
                    & (pl.col("IsMultiValue") == "FALSE")
                    & (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("ok"))

                .when(
                    (pl.col("FunctionName").str.to_lowercase() != "packing")
                    & (pl.col("IsMultiValue") == "FALSE")
                    & (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("null"))

                .otherwise(
                    pl.col("QAComment")
                )

                .alias("QAComment")
            )

            # ==================================================
            # ADD AREA / GROUP / TEAM LEADER
            # ==================================================

            function_lower = (
                pl.col("FunctionName")
                .str.to_lowercase()
            )

            # Build mapping DataFrame
            mapping_df = reference_df.select([
                pl.col("FunctionName")
                .str.to_lowercase()
                .alias("FunctionName_lookup"),

                pl.col("Action"),

                pl.col("Area"),

                pl.col("Group"),

                pl.col("Team leader")
            ])

            # Add lookup column
            df = df.with_columns(
                function_lower.alias(
                    "FunctionName_lookup"
                )
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
            # REMOVE TEMP LOOKUP COLUMN
            # ==================================================

            df = df.drop(
                "FunctionName_lookup"
            )

            # ==================================================
            # REPLACE NULL REFERENCE VALUES
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
            #
            # FunctionName
            # Area
            # Group
            # Team leader
            # ==================================================

            columns = df.columns

            other_columns = [
                c for c in columns
                if c not in [
                    "FunctionName",
                    "Area",
                    "Group",
                    "Team leader"
                ]
            ]

            df = df.select([
                "FunctionName",
                "Area",
                "Group",
                "Team leader"
            ] + other_columns)

            # ==================================================
            # REMOVE ACTION = ANALYSIS
            # ==================================================

            before_analysis_filter = df.height

            df = df.filter(
                pl.col("Action")
                .str.to_lowercase()
                != "analysis"
            )

            removed_analysis = (
                before_analysis_filter - df.height
            )

            # ==================================================
            # REMOVE QACOMMENT = OK
            # ==================================================

            before_ok_filter = df.height

            df = df.filter(
                pl.col("QAComment")
                .str.to_lowercase()
                != "ok"
            )

            removed_ok = (
                before_ok_filter - df.height
            )

            # ==================================================
            # SUMMARY DATA
            #
            # Only conflict / null rows remain
            # ==================================================

            if df.height > 0:

                summary_df = (
                    df
                    .filter(
                        pl.col("QAComment")
                        .str.to_lowercase()
                        .is_in(["conflict", "null"])
                    )
                    .group_by([
                        "Area",
                        "Team leader"
                    ])
                    .agg(
                        pl.len().alias("Counts")
                    )
                    .sort(
                        "Counts",
                        descending=True
                    )
                )

                # Convert to Python rows
                for row in summary_df.iter_rows(
                    named=True
                ):

                    all_summary_rows.append({
                        "Area": row["Area"],
                        "Counts": row["Counts"],
                        "Team Leader": row["Team leader"]
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

            generated_files.append({
                "name": output_name,
                "data": output_bytes
            })

            # ==================================================
            # STATISTICS
            # ==================================================

            st.success(
                f"Created: {output_name}"
            )

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
                f"Preview - {output_name}"
            ):

                st.dataframe(
                    df.head(100).to_pandas(),
                    use_container_width=True
                )

        except Exception as e:

            st.error(
                f"Error processing `{filename}`: {str(e)}"
            )

    # ========================================================
    # COMBINE SUMMARY
    # ========================================================

    if all_summary_rows:

        summary_pl = pl.DataFrame(
            all_summary_rows
        )

        # Combine same Area + Team Leader
        summary_pl = (
            summary_pl
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

        # Rename exactly as requested
        summary_pl = summary_pl.select([
            "Area",
            "Counts",
            "Team Leader"
        ])

        # ====================================================
        # SUMMARY DISPLAY
        # ====================================================

        st.divider()

        st.header("📋 SUMMARY")

        st.dataframe(
            summary_pl.to_pandas(),
            use_container_width=True,
            hide_index=True
        )

        # ====================================================
        # SUMMARY EXCEL
        # ====================================================

        summary_buffer = io.BytesIO()

        summary_pl.write_excel(
            summary_buffer
        )

        summary_buffer.seek(0)

        st.download_button(
            label="⬇️ Download Summary Excel",
            data=summary_buffer.getvalue(),
            file_name="QA_Summary.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    # ========================================================
    # DOWNLOAD UPDATED FILES
    # ========================================================

    if generated_files:

        st.divider()

        st.header("📥 Download Results")

        # ====================================================
        # ONE FILE
        # ====================================================

        if len(generated_files) == 1:

            file_info = generated_files[0]

            st.download_button(
                label="⬇️ Download Updated Excel",
                data=file_info["data"],
                file_name=file_info["name"],
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                )
            )

        # ====================================================
        # MULTIPLE FILES
        # ====================================================

        else:

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
                label="⬇️ Download All Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Updated_Output_Files.zip",
                mime="application/zip"
            )

            # =================================================
            # INDIVIDUAL DOWNLOADS
            # =================================================

            st.write("### Individual Files")

            for index, file_info in enumerate(
                generated_files
            ):

                st.download_button(
                    label=f"⬇️ {file_info['name']}",
                    data=file_info["data"],
                    file_name=file_info["name"],
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    key=f"download_{index}"
                )

        st.success(
            "✅ Finished Successfully!"
        )

    else:

        st.warning(
            "No files were generated."
        )
