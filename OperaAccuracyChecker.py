import streamlit as st
import pandas as pd
from xml.etree import ElementTree
from datetime import datetime
import plotly.graph_objs as go
from plotly.subplots import make_subplots

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Sciant Integration Accuracy Calculator")

# Define the function to parse the XML
def parse_xml(xml_content):
    # Parse the XML content
    tree = ElementTree.fromstring(xml_content)
    # Extract relevant data
    data = []
    system_time = datetime.strptime(tree.find('G_RESORT/SYSTEM_TIME').text, "%d-%b-%y %H:%M:%S")
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
def color_accuracy(val):
    color = 'red'  # default color
    if '%' in val:  # check if value is a percentage
        num = float(val.strip('%'))
        if 'RNs' in val:  # RNs column color coding
            if num >= 98.5:
                color = 'green'
            elif num >= 95:
                color = 'orange'
        else:  # Revenue column color coding
            if num >= 95:
                color = 'green'
            elif num >= 90:
                color = 'orange'
    return f'background-color: {color}; color: white;' if color != 'red' else f'color: {color};'

# Streamlit application
def main():
    st.markdown("<h1 style='text-align: center;'>Sciant Integration Accuracy Calculator</h1>", unsafe_allow_html=True)

    st.warning("The reference date of the Daily Totals Extract should be equal to the History and Forecast file date.")

    col1, col2 = st.columns(2)
    with col1:
        xml_files = st.file_uploader("Upload History and Forecast .xml", type=['xml'], accept_multiple_files=True)
    with col2:
        csv_file = st.file_uploader("Upload Daily Totals Extract from Support UI", type=['csv'])

    if xml_files and csv_file:
        combined_xml_df = pd.DataFrame()
        for xml_file in xml_files:
            xml_df = parse_xml(xml_file.getvalue())
            combined_xml_df = pd.concat([combined_xml_df, xml_df])

        combined_xml_df = combined_xml_df.sort_values(by=['date', 'system_time'], ascending=[True, False])
        combined_xml_df = combined_xml_df.drop_duplicates(subset=['date'], keep='first')

        csv_df = pd.read_csv(csv_file, delimiter=';', quotechar='"')
        csv_df.columns = [col.replace('"', '').strip() for col in csv_df.columns]
        csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')
        csv_df['Juyo RN'] = csv_df['rn'].astype(int)
        csv_df['Juyo Rev'] = csv_df['revNet'].astype(float)

        combined_xml_df['date'] = pd.to_datetime(combined_xml_df['date'], errors='coerce')
        csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')

        merged_df = pd.merge(combined_xml_df, csv_df, left_on='date', right_on='arrivalDate')

        merged_df['RN Diff'] = merged_df['HF RNs'] - merged_df['Juyo RN']
        merged_df['Rev Diff'] = merged_df['HF Rev'] - merged_df['Juyo Rev']

        current_date = pd.to_datetime('today').normalize()
        past_mask = merged_df['date'] < current_date
        future_mask = merged_df['date'] >= current_date

        past_rooms_accuracy = (1 - (abs(merged_df.loc[past_mask, 'RN Diff']).sum() / merged_df.loc[past_mask, 'HF RNs'].sum())) * 100
        past_revenue_accuracy = (1 - (abs(merged_df.loc[past_mask, 'Rev Diff']).sum() / merged_df.loc[past_mask, 'HF Rev'].sum())) * 100
        future_rooms_accuracy = (1 - (abs(merged_df.loc[future_mask, 'RN Diff']).sum() / merged_df.loc[future_mask, 'HF RNs'].sum())) * 100
        future_revenue_accuracy = (1 - (abs(merged_df.loc[future_mask, 'Rev Diff']).sum() / merged_df.loc[future_mask, 'HF Rev'].sum())) * 100

        accuracy_data = {
            "RNs": [f"{past_rooms_accuracy:.2f}%", f"{future_rooms_accuracy:.2f}%"],
            "Revenue": [f"{past_revenue_accuracy:.2f}%", f"{future_revenue_accuracy:.2f}%"]
        }
        accuracy_df = pd.DataFrame(accuracy_data, index=["Past", "Future"])

        st.table(accuracy_df.style.applymap(color_accuracy).set_table_styles([{"selector": "th", "props": [("backgroundColor", "#f0f2f6")]}]))

        st.warning("Future discrepancies might be a result of timing discrepancies between the moment that the data was received and the moment that the history and forecast file was received.")

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        fig.add_trace(go.Bar(
            x=merged_df['date'],
            y=merged_df['RN Diff'],
            name='RNs Discrepancy',
            marker_color='blue'
        ), secondary_y=False)

        fig.add_trace(go.Scatter(
            x=merged_df['date'],
            y=merged_df['Rev Diff'],
            name='Revenue Discrepancy',
            mode='lines+markers',
            line=dict(color='red', width=2),
            marker=dict(size=8)
        ), secondary_y=True)

        fig.update_layout(
            height=600,
            title='RNs and Revenue Discrepancy Over Time',
            xaxis_title='Date',
            yaxis_title='RNs Discrepancy',
            yaxis2_title='Revenue Discrepancy',
            yaxis=dict(range=[-max(merged_df['RN Diff'].abs()), max(merged_df['RN Diff'].abs())]),
            yaxis2=dict(range=[-max(merged_df['Rev Diff'].abs()), max(merged_df['Rev Diff'].abs())]),
            legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
        )

        st.plotly_chart(fig, use_container_width=True)

if __name__ == "__main__":
    main()
