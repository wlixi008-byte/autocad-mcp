import streamlit as st
import os
import tempfile
import converter

st.set_page_config(page_title="AutoCAD Skill Generator", layout="wide")

st.title("🏗️ AutoCAD DWG to Skill Converter")
st.markdown("Upload a DWG file to extract its geometry and generate a Python capability (Skill) to reproduce it.")

uploaded_file = st.file_uploader("Choose a DWG file", type=["dwg"])

if uploaded_file is not None:
    # Save uploaded file to temp
    # Note: AutoCAD requires a real file path on disk
    with tempfile.NamedTemporaryFile(delete=False, suffix=".dwg") as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        tmp_path = tmp_file.name

    st.success(f"File uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Generate Skill"):
        with st.spinner("Connecting to AutoCAD and scanning drawing..."):
            entities = converter.extract_entities_from_file(tmp_path)
        
        if entities:
            st.success(f"Successfully extracted {len(entities)} entities!")
            
            # Generate Code
            skill_code = converter.generate_python_skill(entities)
            
            # Display Code
            st.subheader("Generated Python Skill")
            st.code(skill_code, language="python")
            
            # Download Button
            st.download_button(
                label="💾 Download Skill Script",
                data=skill_code,
                file_name="generated_skill.py",
                mime="text/x-python"
            )
        else:
            st.error("No entities found or failed to read file.")
            st.info("Troubleshooting:\n1. Ensure AutoCAD is running.\n2. Ensure no modal dialogs (like 'Open File') are blocking AutoCAD.\n3. Try opening the file manually in AutoCAD first, then uploading.")
            
    # Clean up (Optional, usually temp files persist slightly or handled by OS)
