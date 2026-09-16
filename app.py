from io import BytesIO
import hashlib

import streamlit as st

from renderer import render

st.set_page_config(page_title="Organization Chart builder", page_icon=":material/account_tree:", layout="wide")

st.markdown("""
<style>
  .stApp { background: #f6f8f5; }
  [data-testid="stHeader"] { background: transparent; }
  .hero { background: #173328; color: white; border-radius: 8px; padding: 28px 34px; margin-bottom: 24px; }
  .hero h1 { margin: 0; font-size: 30px; letter-spacing: 0; }
  .hero p { margin: 8px 0 0; color: #d6e4db; font-size: 16px; }
  [data-testid="stFileUploader"] { background: white; border: 1px solid #d6e0d9; border-radius: 8px; padding: 14px; }
</style>
""", unsafe_allow_html=True)
st.markdown("""<div class="hero"><h1>HR organization chart</h1><p>Upload the latest HR export to generate the approved single-page organization chart.</p></div>""", unsafe_allow_html=True)

left, right = st.columns((1, 2), gap="large", vertical_alignment="top")
with left:
    st.subheader("HR workbook")
    uploaded = st.file_uploader("Excel export", type=("xlsx", "xlsm"), label_visibility="collapsed")
    st.caption("Required columns: Employee, Department, Manager, and Title. Active rows only are included when Employment status is present.")
    st.subheader("Approved template")
    st.caption("The PDF uses the same approved layout, department colours, team boxes, and compact alignment rules as the current organization chart.")

if uploaded:
    upload_key = hashlib.sha256(uploaded.getvalue()).hexdigest()
    if st.session_state.get("chart_upload_key") != upload_key:
        st.session_state.pop("chart_result", None)
        st.session_state["chart_upload_key"] = upload_key
    if st.button("Build single-page chart", type="primary", icon=":material/picture_as_pdf:"):
        with st.spinner("Reading the uploaded workbook and building the chart..."):
            result = render(BytesIO(uploaded.getvalue()))
            st.session_state["chart_result"] = result
    result = st.session_state.get("chart_result")
    if result and result.errors:
        for error in result.errors:
            st.error(error, icon=":material/error:")
    elif result:
        for warning in result.warnings:
            st.warning(warning, icon=":material/warning:")
        if result.pdf and result.png:
            with right:
                st.subheader("Chart preview")
                st.image(result.png, width="stretch")
            st.success(f"Chart generated from this upload for {len(result.people)} people.", icon=":material/check_circle:")
            st.download_button("Download single-page PDF", data=result.pdf,
                               file_name="Organization_Chart_SinglePage.pdf", mime="application/pdf",
                               icon=":material/download:", type="primary")
else:
    with right:
        st.subheader("Chart preview")
        st.info("Upload an HR workbook to create the chart preview and PDF.", icon=":material/upload_file:")