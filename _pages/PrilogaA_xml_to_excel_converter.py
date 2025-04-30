import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

st.title("XML to Excel Converter")

uploaded_file = st.file_uploader("Upload XML file", type=["xml"])

def parse_xml_to_dataframes(xml_content):
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    namespaces = {
        'gl': 'http://www.gesmes.org/xml/2002-08-01',
        'ns': 'http://schemas.lis.energy/billing/1.1',
    }

    data = []
    measurement_rows = []

    for reading in root.findall('.//ns:Readings/ns:Reading', namespaces):
        for point in reading.findall('.//ns:IntervalReading', namespaces):
            quantity = point.find('ns:Quantity', namespaces)
            time = point.find('ns:Time', namespaces)

            if quantity is not None and time is not None:
                measurement_rows.append({
                    'Time': time.text,
                    'Quantity': quantity.text
                })

    measurement_df = pd.DataFrame(measurement_rows)

    return measurement_df

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Measurements')
    output.seek(0)
    return output

if uploaded_file:
    xml_content = uploaded_file.read()

    try:
        measurement_df = parse_xml_to_dataframes(xml_content)

        if not measurement_df.empty:
            st.success("XML parsed successfully.")
            st.dataframe(measurement_df)

            excel_data = convert_df_to_excel(measurement_df)
            st.download_button(
                label="Download Excel file",
                data=excel_data,
                file_name="measurement_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No measurement data found in the XML.")
    except Exception as e:
        st.error(f"Error processing XML file: {e}")
