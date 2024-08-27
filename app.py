import streamlit as st
import openpyxl

# Load the Excel workbook and sheet
npi_file = openpyxl.load_workbook("drc_npi.xlsx")
npi_report = npi_file["Sheet"]

# Extract node-wise traffic data
node_wise_traffic = {}
for row in range(118, 134):
    node_name = npi_report.cell(row, 3).value.upper().replace(' ', '')
    traffic = npi_report.cell(row, 7).value
    node_wise_traffic[node_name] = traffic

total_traffic = sum(node_wise_traffic.values())

# Streamlit app
st.title("Network Level Impact Calculator(2.0)")
st.subheader("By: Shadman")

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
