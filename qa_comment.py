
import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="QA Comment Test",
    page_icon="🧪"
)

st.title("🧪 QA Comment Diagnostic")

st.success("✅ Streamlit is working")

st.write("Testing Pandas...")

try:
    test = pd.DataFrame({
        "Test": ["OK"],
        "Value": [1]
    })

    st.success("✅ Pandas is working")
    st.dataframe(test)

except Exception as e:
    st.error("❌ Pandas failed")
    st.exception(e)
    st.stop()


st.write("Testing reference.xlsx...")

reference_file = (
    Path(__file__).resolve().parent
    / "reference.xlsx"
)


if not reference_file.exists():

    st.error("❌ reference.xlsx NOT FOUND")

    st.code(
        str(reference_file)
    )

    st.stop()


st.success(
    "✅ reference.xlsx found"
)


try:

    reference = pd.read_excel(
        reference_file,
        engine="openpyxl"
    )

    st.success(
        f"✅ Reference loaded: "
        f"{len(reference):,} rows"
    )

    st.write(
        "Columns:"
    )

    st.write(
        list(reference.columns)
    )

    st.dataframe(
        reference.head(20),
        use_container_width=True
    )

except Exception as e:

    st.error(
        "❌ Could not read reference.xlsx"
    )

    st.exception(e)

    st.stop()


st.success(
    "🎉 DIAGNOSTIC TEST PASSED"
)
