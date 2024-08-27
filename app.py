import openpyxl
import requests
from io import BytesIO
import streamlit as st

# Raw URL of the file on GitHub
file_url = "https://github.com/ShadSpidy/Network-Level-Impact-2.0/raw/main/drc_npi.xlsx"

def download_file(url):
    response = requests.get(url)
    response.raise_for_status()
    return BytesIO(response.content)

# Define available OPCOs and node types
opcos = ["Nigeria", "DRC", "Chad"]
node_types = ["SDP", "OCC", "AIR"]

try:
    # Load the Excel workbook from GitHub
    response = download_file(file_url)
    npi_file = openpyxl.load_workbook(response)
    npi_report = npi_file["Sheet"]

    # Streamlit app
    st.title("Network Impact Calculator")

    # OPCO selection (multiple options)
    selected_opcos = st.multiselect("Choose OPCO(s):", opcos)

    # Node type selection
    selected_node_types = st.multiselect("Choose node type(s):", node_types)

    if selected_node_types:
        # Initialize node selection options
        node_options = {}
        for node_type in selected_node_types:
            node_options[node_type] = []

        # Populate node options based on node types and selected OPCOs
        for row in range(118, 134):
            opco = npi_report.cell(row, 1).value.upper().replace(' ', '')  # Assuming OPCO is in column 1
            node_type = npi_report.cell(row, 2).value.upper().replace(' ', '')  # Assuming node type is in column 2
            node_name = npi_report.cell(row, 3).value.upper().replace(' ', '')  # Assuming node name is in column 3
            if opco in selected_opcos and node_type in selected_node_types:
                if node_type not in node_options:
                    node_options[node_type] = []
                if node_name not in node_options[node_type]:
                    node_options[node_type].append(node_name)

        # Display node selection based on chosen types
        selected_nodes = {}
        for node_type in selected_node_types:
            selected_nodes[node_type] = st.multiselect(f"Choose {node_type} nodes:", node_options[node_type])

        if selected_nodes:
            # Calculate traffic for selected nodes
            impacted_node_traffic = 0
            total_traffic = 0
            for node_type, nodes in selected_nodes.items():
                for node in nodes:
                    for row in range(118, 134):
                        if (npi_report.cell(row, 1).value.upper().replace(' ', '') in selected_opcos and
                            npi_report.cell(row, 2).value.upper().replace(' ', '') == node_type and
                            npi_report.cell(row, 3).value.upper().replace(' ', '') == node):
                            impacted_node_traffic += npi_report.cell(row, 7).value

            # Calculate total traffic for the selected OPCOs
            for row in range(118, 134):
                if npi_report.cell(row, 1).value.upper().replace(' ', '') in selected_opcos:
                    total_traffic += npi_report.cell(row, 7).value

            if impacted_node_traffic == 0:
                st.error("Invalid input or no valid nodes entered. Please enter valid node names.")
            else:
                # User input for current percentage impact
                percentage_impact = st.number_input("Enter the current percentage impact (%):", min_value=0.0, max_value=100.0, step=0.1)

                # Calculate network level impact
                NW_level_impact = ((impacted_node_traffic / total_traffic) * 100) * (percentage_impact / 100)
                st.write(f"Current Network Level Impact is {round(NW_level_impact, 2)}%")
        else:
            st.info("Please select nodes to get results.")
    else:
        st.info("Please select node type(s) to get results.")

except Exception as e:
    st.error(f"An error occurred: {e}")
