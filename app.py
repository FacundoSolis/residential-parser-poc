"""
Residential Energy Certificate Parser - Web Interface
"""

import streamlit as st
import zipfile
import tempfile
import shutil
from pathlib import Path
import sys

# Add both root and src to path for imports
root_path = str(Path(__file__).parent)
src_path = str(Path(__file__).parent / 'src')
sys.path.insert(0, root_path)
sys.path.insert(0, src_path)

from src.matrix_generator import MatrixGenerator

st.set_page_config(
    page_title="Residential Parser",
    page_icon="🏠",
    layout="wide"
)

st.title("🏠 Residential Energy Certificate Parser")
st.markdown("Upload a project folder (ZIP) to extract data and generate Excel report")

# File uploader
uploaded_file = st.file_uploader(
    "Upload Project ZIP File",
    type=['zip'],
    help="Upload a ZIP file containing residential energy documents"
)

if uploaded_file is not None:
    st.success(f"File uploaded: {uploaded_file.name}")
    
    # Create temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Extract ZIP
        with st.spinner("Extracting files..."):
            zip_path = temp_path / uploaded_file.name
            with open(zip_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            
            extract_path = temp_path / "project"
            extract_path.mkdir()
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
        
        st.success("Files extracted successfully")
        
        # Parse documents
        with st.spinner("Parsing documents..."):
            # Find the actual project folder (might be nested)
            project_folders = list(extract_path.glob("*"))
            if len(project_folders) == 1 and project_folders[0].is_dir():
                project_path = project_folders[0]
            else:
                project_path = extract_path
            
            # Generate matrix
            generator = MatrixGenerator(str(project_path))
            generator.parse_all_documents()
            
            # Generate Excel
            output_path = temp_path / f"{project_path.name}_Checks.xlsx"
            generator.generate_excel(str(output_path), project_path.name)
        
        st.success("✅ Processing complete!")
        
        # Show extracted data summary
        st.subheader("Extracted Data Summary")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Documents Parsed", len(generator.parsed_data))
        
        with col2:
            # Try CONTRATO first, then CALCULO as fallback
            homeowner = generator.parsed_data.get('CONTRATO', {}).get('homeowner_name')
            if not homeowner or homeowner == 'NOT FOUND':
                homeowner = generator.parsed_data.get('CALCULO', {}).get('client_name')
            if not homeowner or homeowner == 'NOT FOUND':
                homeowner = generator.parsed_data.get('FACTURA', {}).get('homeowner_name')
            st.metric("Homeowner", homeowner or 'NOT FOUND')
        
        with col3:
            # Priority: CALCULO ae > CERTIFICADO energy_savings > CONTRATO energy_savings
            energy = generator.parsed_data.get('CALCULO', {}).get('ae')
            if not energy or energy == 'NOT FOUND':
                energy = generator.parsed_data.get('CERTIFICADO', {}).get('energy_savings')
            if not energy or energy == 'NOT FOUND':
                energy = generator.parsed_data.get('CONTRATO', {}).get('energy_savings')
            st.metric("Energy Savings (kWh)", energy or 'NOT FOUND')
        
        # Show parsed documents
        with st.expander("View Parsed Documents"):
            for doc_type, data in generator.parsed_data.items():
                st.write(f"**{doc_type}:**")
                st.json(data)
        
        # Download button
        with open(output_path, 'rb') as f:
            st.download_button(
                label="⬇️ Download Excel Report",
                data=f.read(),
                file_name=f"{project_path.name}_Checks.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Sidebar with instructions
with st.sidebar:
    st.header("📖 Instructions")
    st.markdown("""
    1. **Prepare your project folder** with these documents:
       - CONTRATO CESION AHORROS
       - CERTIFICADO INSTALADOR
       - FACTURA
       - DECLARACION RESPONSABLE
       - CEE FINAL
       - REGISTRO CEE
       - DNI
       - CALCULO.xlsx
    
    2. **ZIP the folder**
    
    3. **Upload the ZIP** using the button above
    
    4. **Download** the generated Excel report
    """)
    
    st.header("📊 Sample Projects")
    st.markdown("""
    - DE PAZ FRANCO QUINTILIANA
    - sample_1 (RABANAL ALONSO)
    """)
