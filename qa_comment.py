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

            file_bytes = (
                uploaded_file.getvalue()
            )

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
                    f"Skipping `{filename}`. "
                    f"Missing columns: "
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

            function_name = (
                pl.col("FunctionName")
                .str.to_lowercase()
            )

            df = df.with_columns(

                # ----------------------------------------------
                # PACKING
                # TRUE + FALSE = OK
                # ----------------------------------------------

                pl.when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "TRUE")
                    &
                    (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("ok"))

                # ----------------------------------------------
                # PACKING
                # TRUE + TRUE = NULL
                # ----------------------------------------------

                .when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "TRUE")
                    &
                    (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("null"))

                # ----------------------------------------------
                # PACKING
                # FALSE + FALSE = CONFLICT
                # ----------------------------------------------

                .when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "FALSE")
                    &
                    (pl.col("HasBlankValue") == "FALSE")
                )
                .then(pl.lit("conflict"))

                # ----------------------------------------------
                # PACKING
                # FALSE + TRUE = CONFLICT
                # ----------------------------------------------

                .when(
                    (function_name == "packing")
                    &
                    (pl.col("IsMultiValue") == "FALSE")
                    &
                    (pl.col("HasBlankValue") == "TRUE")
                )
                .then(pl.lit("conflict"))

                # ==============================================
                # OTHER FUNCTIONS
                # ==============================================

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
            # PREPARE REFERENCE
            # ==================================================

            mapping_df = reference_df.select([

                pl.col("FunctionName")
                .str.to_lowercase()
                .alias("FunctionName_lookup"),

                pl.col("Action"),

                pl.col("Area"),

                pl.col("Group"),

                pl.col("Team leader")

            ])

            # ==================================================
            # CREATE LOOKUP COLUMN
            # ==================================================

            df = df.with_columns(

                pl.col("FunctionName")
                .str.to_lowercase()
                .alias("FunctionName_lookup")

            )

            # ==================================================
            # JOIN
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
            # FILL MISSING VALUES
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

            remaining_columns = [
                c for c in df.columns
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

            ] + remaining_columns)

            # ==================================================
            # REMOVE ACTION = ANALYSIS
            # ==================================================

            before_analysis = df.height

            df = df.filter(

                pl.col("Action")
                .str.to_lowercase()
                != "analysis"

            )

            removed_analysis = (
                before_analysis - df.height
            )

            # ==================================================
            # REMOVE QAComment = OK
            # ==================================================

            before_ok = df.height

            df = df.filter(

                pl.col("QAComment")
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

            output_name = (
                filename[:-5]
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
            # SAVE OUTPUT
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
                f"✅ Created: {output_name}"
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


    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    st.divider()

    st.header(
        "📋 FINAL SUMMARY"
    )

    if all_summary_rows:

        final_summary = pl.DataFrame(
            all_summary_rows
        )

        # Combine same Area + Team Leader
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

            "Area": [],

            "Counts": [],

            "Team Leader": []

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
