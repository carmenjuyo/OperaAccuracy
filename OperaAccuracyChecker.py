import streamlit as st
import pandas as pd
from xml.etree import ElementTree
from datetime import datetime
import plotly.graph_objs as go
from plotly.subplots import make_subplots
from io import BytesIO

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Opera Daily Variance and Accuracy Calculator")

# Define the function to parse the XML
def parse_xml(xml_content, filename):
    # Extract the date from the filename
    try:
        file_date = datetime.strptime(filename.split('_')[0], "%Y%m%d")
    except ValueError:
        # Handle the case where the date format in the filename is incorrect
        file_date = None

    # Parse the XML content
    tree = ElementTree.fromstring(xml_content)
    
    # Attempt to find the SYSTEM_TIME element
    system_time_element = tree.find('G_RESORT/SYSTEM_TIME')
    if system_time_element is not None:
        system_time = datetime.strptime(system_time_element.text, "%d-%b-%y %H:%M:%S")
    else:
        # Use the file_date if SYSTEM_TIME is missing
        if file_date:
            system_time = file_date
        else:
            raise ValueError("Both SYSTEM_TIME and a valid date in the filename are missing.")

    data = []
    for g_considered_date in tree.iter('G_CONSIDERED_DATE'):
        date = g_considered_date.find('CONSIDERED_DATE').text
        ind_deduct_rooms = int(g_considered_date.find('IND_DEDUCT_ROOMS').text)
        grp_deduct_rooms = int(g_considered_date.find('GRP_DEDUCT_ROOMS').text)
        ind_deduct_revenue = float(g_considered_date.find('IND_DEDUCT_REVENUE').text)
        grp_deduct_revenue = float(g_considered_date.find('GRP_DEDUCT_REVENUE').text)
        date = datetime.strptime(date, "%d-%b-%y").date()  # Assuming date is in 'DD-MMM-YY' format
        data.append({
            'date': date,
            'system_time': system_time,
            'HF RNs': ind_deduct_rooms + grp_deduct_rooms,
            'HF Rev': ind_deduct_revenue + grp_deduct_revenue
        })
    return pd.DataFrame(data)

# Define color coding for accuracy values
def color_scale(val):
    """Color scale for percentages."""
    if isinstance(val, str) and '%' in val:
        val = float(val.strip('%'))
        if val >= 98:
            return 'background-color: #469798; color: white;'  # green
        elif 95 <= val < 98:
            return 'background-color: #F2A541; color: white;'  # yellow
        else:
            return 'background-color: #BF3100; color: white;'  # red
    return ''

# Function to create Excel file for download with color formatting and accuracy matrix
def create_excel_download(combined_df, base_filename, past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Write the Accuracy Matrix
        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [past_accuracy_rn / 100, past_accuracy_rev / 100],  # Store as decimal
            'Future': [future_accuracy_rn / 100, future_accuracy_rev / 100]  # Store as decimal
        })
        
        accuracy_matrix.to_excel(writer, sheet_name='Accuracy Matrix', index=False, startrow=1)
        worksheet = writer.sheets['Accuracy Matrix']

        # Define formats
        format_date = workbook.add_format({'num_format': 'dd-mmm-yyyy'})
        format_whole = workbook.add_format({'num_format': '0'})  # Whole numbers
        format_float = workbook.add_format({'num_format': '0.00'})  # Floats
        format_number = workbook.add_format({'num_format': '#,##0.00'})  # Number with thousands separator and two decimals
        format_percent = workbook.add_format({'num_format': '0.00%'})  # Percentage format

        # Apply percentage format to the relevant cells in Accuracy Matrix
        worksheet.set_column('B:B', None, format_percent)
        worksheet.set_column('C:C', None, format_percent)

        # Apply simplified conditional formatting for Accuracy Matrix
        format_green = workbook.add_format({'bg_color': '#469798', 'font_color': '#FFFFFF'})
        format_yellow = workbook.add_format({'bg_color': '#F2A541', 'font_color': '#FFFFFF'})
        format_red = workbook.add_format({'bg_color': '#BF3100', 'font_color': '#FFFFFF'})
        
        worksheet.conditional_format('B3:B4', {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
        worksheet.conditional_format('B3:B4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
        worksheet.conditional_format('B3:B4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

        worksheet.conditional_format('C3:C4', {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
        worksheet.conditional_format('C3:C4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
        worksheet.conditional_format('C3:C4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

        # Write the combined past and future results to a single sheet
        if not combined_df.empty:
            combined_df.to_excel(writer, sheet_name='Daily Variance Detail', index=False)
            worksheet_combined = writer.sheets['Daily Variance Detail']

            # Format columns
            worksheet_combined.set_column('A:A', None, format_date)    # Date
            worksheet_combined.set_column('B:B', None, format_whole)   # Whole numbers
            worksheet_combined.set_column('C:C', None, format_float)   # Floats
            worksheet_combined.set_column('D:D', None, format_number)  # Numbers
            worksheet_combined.set_column('E:E', None, format_float)   # Floats
            worksheet_combined.set_column('F:F', None, format_number)  # Numbers
            worksheet_combined.set_column('G:G', None, format_float)   # Floats
            worksheet_combined.set_column('H:H', None, format_percent) # Percentages
            worksheet_combined.set_column('I:I', None, format_percent) # Percentages

            # Apply conditional formatting to the percentage columns (H and I)
            worksheet_combined.conditional_format('H2:H{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_combined.conditional_format('H2:H{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_combined.conditional_format('H2:H{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

            worksheet_combined.conditional_format('I2:I{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_combined.conditional_format('I2:I{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_combined.conditional_format('I2:I{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})
    output.seek(0)
    return output, base_filename

# Streamlit application
def main():
    # Center the title using markdown with HTML
    st.markdown("<h1 style='text-align: center;'> Opera Daily Variance and Accuracy Calculator</h1>", unsafe_allow_html=True)

    # Note/warning box with matching colors
    st.warning("The reference date of the Daily Totals Extract should be equal to the latest History and Forecast file date.")

    # File uploaders in columns
    col1, col2 = st.columns(2)
    with col1:
        xml_files = st.file_uploader("Upload History and Forecast .xml", type=['xml'], accept_multiple_files=True, key="xml_uploader")
    with col2:
        csv_file = st.file_uploader("Upload Daily Totals Extract from Support UI", type=['csv'], key="csv_uploader")

    if xml_files and csv_file:
        load_data_button = st.button("Load Data")
        if load_data_button:
            with st.spinner("Loading data..."):
                combined_xml_df = pd.DataFrame()
                for xml_file in xml_files:
                    xml_df = parse_xml(xml_file.getvalue(), xml_file.name)
                    combined_xml_df = pd.concat([combined_xml_df, xml_df])

                # Keep only the latest entry for each considered date
                combined_xml_df = combined_xml_df.sort_values(by=['date', 'system_time'], ascending=[True, False])
                combined_xml_df = combined_xml_df.drop_duplicates(subset=['date'], keep='first')

                # Process CSV
                csv_df = pd.read_csv(csv_file, delimiter=';', quotechar='"')

                # Drop any columns that are completely empty (all NaN)
                csv_df = csv_df.dropna(axis=1, how='all')

                csv_df.columns = [col.replace('"', '').strip() for col in csv_df.columns]
                csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')
                csv_df['Juyo RN'] = csv_df['rn'].astype(int)
                csv_df['Juyo Rev'] = csv_df['revNet'].astype(float)

                # Ensure both date columns are of type datetime64[ns] before merging
                combined_xml_df['date'] = pd.to_datetime(combined_xml_df['date'], errors='coerce')
                csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')

                st.success("Data loaded successfully! Now proceed with processing.")

                # Enable the processing button after loading data
                process_button = st.button("Process Data")
                if process_button:
                    with st.spinner("Processing data..."):
                        # Merge data
                        merged_df = pd.merge(combined_xml_df, csv_df, left_on='date', right_on='arrivalDate')

                        # Calculate discrepancies for rooms and revenue
                        merged_df['RN Diff'] = merged_df['Juyo RN'] - merged_df['HF RNs']
                        merged_df['Rev Diff'] = merged_df['Juyo Rev'] - merged_df['HF Rev']

                        # Calculate absolute accuracy percentages
                        merged_df['Abs RN Accuracy'] = (1 - abs(merged_df['RN Diff']) / merged_df['HF RNs']) * 100
                        merged_df['Abs Rev Accuracy'] = (1 - abs(merged_df['Rev Diff']) / merged_df['HF Rev']) * 100

                        # Format accuracy percentages as strings with '%' symbol
                        merged_df['Abs RN Accuracy'] = merged_df['Abs RN Accuracy'].map(lambda x: f"{x:.2f}%")
                        merged_df['Abs Rev Accuracy'] = merged_df['Abs Rev Accuracy'].map(lambda x: f"{x:.2f}%")

                        # Calculate overall accuracies
                        current_date = pd.to_datetime('today').normalize()  # Get the current date without the time part
                        past_mask = merged_df['date'] < current_date
                        future_mask = merged_df['date'] >= current_date

                        past_rooms_accuracy = (1 - (abs(merged_df.loc[past_mask, 'RN Diff']).sum() / merged_df.loc[past_mask, 'HF RNs'].sum())) * 100
                        past_revenue_accuracy = (1 - (abs(merged_df.loc[past_mask, 'Rev Diff']).sum() / merged_df.loc[past_mask, 'HF Rev'].sum())) * 100
                        future_rooms_accuracy = (1 - (abs(merged_df.loc[future_mask, 'RN Diff']).sum() / merged_df.loc[future_mask, 'HF RNs'].sum())) * 100
                        future_revenue_accuracy = (1 - (abs(merged_df.loc[future_mask, 'Rev Diff']).sum() / merged_df.loc[future_mask, 'HF Rev'].sum())) * 100

                        # Display accuracy matrix in a table within a container for width control
                        accuracy_data = {
                            "RNs": [f"{past_rooms_accuracy:.2f}%", f"{future_rooms_accuracy:.2f}%"],
                            "Revenue": [f"{past_revenue_accuracy:.2f}%", f"{future_revenue_accuracy:.2f}%"]
                        }
                        accuracy_df = pd.DataFrame(accuracy_data, index=["Past", "Future"])

                        # Center the accuracy matrix table
                        with st.container():
                            st.table(accuracy_df.style.applymap(color_scale).set_table_styles([{"selector": "th", "props": [("backgroundColor", "#f0f2f6")]}]))

                        # Warning about future discrepancies with matching colors
                        st.warning("Future discrepancies might be a result of timing discrepancies between the moment that the data was received and the moment that the history and forecast file was received.")

                        # Interactive bar and line chart for visualizing discrepancies
                        fig = make_subplots(specs=[[{"secondary_y": True}]])

                        # RN Discrepancies - Bar chart
                        fig.add_trace(go.Bar(
                            x=merged_df['date'],
                            y=merged_df['RN Diff'],
                            name='RNs Discrepancy',
                            marker_color='#469798'
                        ), secondary_y=False)

                        # Revenue Discrepancies - Line chart
                        fig.add_trace(go.Scatter(
                            x=merged_df['date'],
                            y=merged_df['Rev Diff'],
                            name='Revenue Discrepancy',
                            mode='lines+markers',
                            line=dict(color='#BF3100', width=2),
                            marker=dict(size=8)
                        ), secondary_y=True)

                        # Update plot layout for dynamic axis scaling and increased height
                        max_room_discrepancy = merged_df['RN Diff'].abs().max()
                        max_revenue_discrepancy = merged_df['Rev Diff'].abs().max()

                        fig.update_layout(
                            height=600,
                            title='RNs and Revenue Discrepancy Over Time',
                            xaxis_title='Date',
                            yaxis_title='RNs Discrepancy',
                            yaxis2_title='Revenue Discrepancy',
                            yaxis=dict(range=[-max_room_discrepancy, max_room_discrepancy]),
                            yaxis2=dict(range=[-max_revenue_discrepancy, max_revenue_discrepancy]),
                            legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
                        )

                        # Align grid lines
                        fig.update_yaxes(matches=None, showgrid=True, gridwidth=1, gridcolor='grey')

                        st.plotly_chart(fig, use_container_width=True)

                        # Display daily variance detail in a table
                        st.markdown("### Daily Variance Detail", unsafe_allow_html=True)
                        detail_container = st.container()
                        with detail_container:
                            formatted_df = merged_df[['date', 'HF RNs', 'HF Rev', 'Juyo RN', 'Juyo Rev', 'RN Diff', 'Rev Diff', 'Abs RN Accuracy', 'Abs Rev Accuracy']]
                            styled_df = formatted_df.style.applymap(color_scale, subset=['Abs RN Accuracy', 'Abs Rev Accuracy']).set_properties(**{'text-align': 'center'})
                            st.table(styled_df)

                        # Combine past and future data for export
                        combined_df = pd.concat([formatted_df[past_mask], formatted_df[future_mask]])

                        # Extract the filename prefix from the uploaded CSV
                        csv_filename = csv_file.name.split('_')[0]
                        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                        base_filename = f"{csv_filename}_AccuracyCheck_{current_time}"

                        # Add Excel export functionality
                        output, filename = create_excel_download(combined_df, base_filename,
                                                                 past_rooms_accuracy, past_revenue_accuracy,
                                                                 future_rooms_accuracy, future_revenue_accuracy)
                        st.download_button(label="Download Excel Report",
                                           data=output,
                                           file_name=f"{filename}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
