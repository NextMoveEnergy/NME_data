import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

def parse_full_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data = {}

    # Recursive function to flatten all elements
    def flatten(element, prefix=''):
        for child in element:
            tag = child.tag.split('}')[-1]
            key = f"{prefix}{tag}" if prefix == '' else f"{prefix}_{tag}"
            if list(child):
                flatten(child, key)
            else:
                text = child.text.strip() if child.text else ''
                if key in data:
                    i = 2
                    while f"{key}_{i}" in data:
                        i += 1
                    key = f"{key}_{i}"
                data[key] = text

    flatten(root)
    return data

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Parsed XML')
    output.seek(0)
    return output

st.title("Priloga A v2.7 XML Parser (All Fields)")

uploaded_files = st.file_uploader("Upload one or more Priloga A XML files", type="xml", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for xml_file in uploaded_files:
        try:
            parsed = parse_full_xml(xml_file)
            parsed["Filename"] = xml_file.name
            all_data.append(parsed)
        except Exception as e:
            st.error(f"Error processing {xml_file.name}: {e}")

    if all_data:
        df = pd.DataFrame(all_data)
        st.dataframe(df)

        excel_data = to_excel(df)
        st.download_button(
            label="Download Excel File",
            data=excel_data,
            file_name="priloga_a_parsed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
