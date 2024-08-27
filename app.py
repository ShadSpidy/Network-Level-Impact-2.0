import openpyxl
import requests
from io import BytesIO

# Raw URL of the file on GitHub
file_url = "https://github.com/ShadSpidy/Network-Level-Impact-2.0/raw/main/drc_npi.xlsx"

def download_file(url):
    response = requests.get(url)
    response.raise_for_status()
    return BytesIO(response.content)

try:
    # Load the Excel workbook from GitHub
    response = download_file(file_url)
    npi_file = openpyxl.load_workbook(response)
    npi_report = npi_file["Sheet"]

    # Extract node-wise traffic data
    node_wise_traffic = {}
    for row in range(118, 134):
        node_name = npi_report.cell(row, 3).value.upper().replace(' ', '')
        traffic = npi_report.cell(row, 7).value
        node_wise_traffic[node_name] = traffic

    total_traffic = sum(node_wise_traffic.values())

    # Streamlit app
    import streamlit as st
    st.title("Network Impact Calculator")

    # User input for impacted nodes
    user_input = st.text_input("Please enter the impacted node(s) (separate multiple nodes with commas):").upper()
    impacted_nodes = [node.strip() for node in user_input.split(',')]

    # Calculate traffic for impacted nodes
    impacted_node_traffic = sum(node_wise_traffic.get(node, 0) for node in impacted_nodes)

    if impacted_node_traffic == 0:
        st.error("Invalid input or no valid nodes entered. Please enter valid node names.")
    else:
        # User input for current percentage impact
        percentage_impact = st.number_input("Enter the current percentage impact (%):", min_value=0.0, max_value=100.0, step=0.1)

        # Calculate network level impact
        NW_level_impact = ((impacted_node_traffic / total_traffic) * 100) * (percentage_impact / 100)
        st.write(f"Current Network Level Impact is {round(NW_level_impact, 2)}%")

except Exception as e:
    print(f"An error occurred: {e}")
