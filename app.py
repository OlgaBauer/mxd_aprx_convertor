import streamlit as st
import arcpy
import os
import pandas as pd
import time

st.title('MXD to APRX Converter\n**by Olga Bauer TWDB BRACS**:balloon:')
logo = r"C:\Users\obauer\Downloads\bracs_logo_landsat.jpg"
st.sidebar.image(logo)

# Upload Excel file with MXD paths (optional)
uploaded_file = st.file_uploader("Upload Excel file with MXD paths", type=["xlsx", "xls"])

# OR manually enter folder path
mxd_folder = st.text_input("Or enter folder path containing MXDs (if not using Excel):")

# Template APRX file (required)
#template_aprx = st.text_input("Enter full path to the template APRX file:") #upload via button 
template_aprx = r"C:\gis_pm\innovative_water\bracs\study\mxd_aprx_conversion_test\template\template.aprx"

# Output folder for new APRX files
#output_folder = st.text_input("Enter output folder for new APRX files:") 

# Start processing when button clicked
if st.button("Convert MXDs to APRX", type = 'primary'):
    try:
        if not template_aprx:
            st.error("Please provide both the template APRX path and output folder.")
        else:
            if uploaded_file:
                # Read MXD paths from uploaded Excel
                df = pd.read_excel(uploaded_file)
                filtered_df = df[(df.iloc[:,4].str.strip().str.lower()=='migrate')]
                mxd_paths = filtered_df.iloc[:, 0].dropna().tolist()
            elif mxd_folder:
                # Use all MXD files in the folder
                mxd_paths = [os.path.join(mxd_folder, f) for f in os.listdir(mxd_folder) if f.endswith(".mxd")]
            else:
                st.error("Please upload an Excel file or specify an MXD folder.")
                st.stop()

            success_count = 0
            for mxd_path in mxd_paths:
                try:
                    mxd_name = os.path.basename(mxd_path)
                    #new_aprx_name = mxd_name.replace(".mxd", ".aprx")
                    #new_aprx_path = os.path.join(output_folder, new_aprx_name) 
                    new_aprx_path = os.path.splitext(mxd_path)[0]+'.aprx'
                    
                    if os.path.exists(new_aprx_path):
                        st.info(f'Skipped: {new_aprx_path} already exists.')
                        continue #skip to the next mxd

                    # Load template project
                    aprx = arcpy.mp.ArcGISProject(template_aprx)
                    aprx.saveACopy(new_aprx_path)

                    # Reopen and import the MXD
                    new_aprx = arcpy.mp.ArcGISProject(new_aprx_path)
                    new_aprx.importDocument(mxd_path)
                    new_aprx.save()

                    success_count += 1
                    st.success(f"Converted: {mxd_name} to aprx", icon="✅")
                except Exception as e:
                    st.warning(f"Failed to convert {mxd_path}: {e}", icon="⚠️")

            st.info(f"Conversion complete. {success_count} files converted to aprx.")
    except Exception as e:
        st.error(f"Error during processing: {e}")
    with st.spinner("Wait for it...", show_time=True):
        time.sleep(3)
    st.success("Done!")
    st.balloons()