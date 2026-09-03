
import streamlit as st
import polars as pl
from pathlib import Path

st.set_page_config(
    page_title="QA Test",
    page_icon="🧪",
)

st.title("🧪 QA Comment App - Diagnostic Test")

st.success("✅ Streamlit started successfully")

st.write("Testing Polars...")

try:
    test_df = pl.DataFrame({
        "Test": ["OK"],
        "Value": [1],
    })

    st.success("✅ Polars is working")
    st.dataframe(test_df)

except Exception as e:
    st.error("❌ Polars test failed")
    st.exception(e)
    st.stop()


st.write("Testing reference.xlsx...")

reference_file = Path(__file__).resolve().parent / "reference.xlsx"

if not reference_file.exists():

    st.error(
        "❌ reference.xlsx NOT FOUND"
    )

    st.write(
        "Expected location:"
    )

    st.code(str(reference_file))

    st.stop()


st.success(
    "✅ reference.xlsx found"
)


try:

    reference = pl.read_excel(
        reference_file,
        engine="calamine",
        infer_schema_length=1000,
    )

    st.success(
        f"✅ Reference loaded successfully — "
        f"{reference.height:,} rows"
    )

    st.write("Columns:")

    st.write(reference.columns)

    st.dataframe(
        reference.head(20),
        use_container_width=True,
    )

except Exception as e:

    st.error(
        "❌ Failed to read reference.xlsx"
    )

    st.exception(e)

    st.stop()


st.success(
    "🎉 ALL DIAGNOSTIC TESTS PASSED"
)

st.info(
    "If you can see this message, "
    "Streamlit + Polars + reference.xlsx "
    "are working correctly."
)
