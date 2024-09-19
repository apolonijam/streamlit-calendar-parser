import streamlit as st
from ics import Calendar
import requests
import pandas as pd
from datetime import datetime
import base64
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement

def fetch_ics_data(ics_url):
    try:
        response = requests.get(ics_url)
        response.raise_for_status()  # Check if the request was successful
        return response.text
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching the ICS file: {e}")
        return None

def parse_ics_data(calendar_data, start_date, end_date):
    calendar = Calendar(calendar_data)
    dates_from, dates_to, times, events = [], [], [], []
    
    for event in calendar.events:
        event_start = event.begin.datetime
        event_end = event.end.datetime if event.end else None

        # Make event_start and event_end naive if they are timezone-aware
        if event_start.tzinfo is not None:
            event_start = event_start.replace(tzinfo=None)
        if event_end and event_end.tzinfo is not None:
            event_end = event_end.replace(tzinfo=None)
        
        if event_start < start_date or (event_end and event_end > end_date):
            continue  # Skip events outside the specified range
        
        dates_from.append(event_start.strftime("%d. %m. %Y"))
        times.append(event_start.strftime("%H:%M"))
        events.append(event.name)
        
        if event_end:
            dates_to.append(event_end.strftime("%d. %m. %Y"))
        else:
            dates_to.append("N/A")  # If no end date, show "N/A"

    return dates_from, dates_to, times, events

def display_table(dates_from, dates_to, times, events):
    table = {'Start Date': dates_from, 'End Date': dates_to, 'Time': times, 'Event Name': events}
    df = pd.DataFrame(table)
    st.dataframe(df)
    
    # Add button to download table as CSV
    st.markdown(get_table_download_link(df, 'csv'), unsafe_allow_html=True)
    
    # Add button to download table as Excel
    st.markdown(get_table_download_link(df, 'xlsx'), unsafe_allow_html=True)
    
    # Add button to download table as Word
    st.markdown(get_table_download_link(df, 'docx'), unsafe_allow_html=True)

def get_table_download_link(df, file_type):
    if file_type == 'csv':
        csv = df.to_csv(index=False)
        b64 = base64.b64encode(csv.encode()).decode()  # Convert to base64
        return f'<a href="data:file/csv;base64,{b64}" download="calendar_events.csv">Download CSV</a>'
    elif file_type == 'xlsx':
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Events')
        b64 = base64.b64encode(output.getvalue()).decode()  # Convert to base64
        return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="calendar_events.xlsx">Download Excel</a>'
    elif file_type == 'docx':
        doc = Document()
        table = doc.add_table(rows=1, cols=len(df.columns))
        
        # Add header row
        hdr_cells = table.rows[0].cells
        for i, column_name in enumerate(df.columns):
            hdr_cells[i].text = column_name
        
        # Add data rows
        for index, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)
        
        # Apply table border styles
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblBorders = tblPr.tblBorders
        if tblBorders is None:
            tblBorders = OxmlElement('w:tblBorders')
            tblPr.append(tblBorders)
        border = OxmlElement('w:bottom')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')  # 1px border size in half-points (4 half-points = 2px)
        tblBorders.append(border)
        
        output = BytesIO()
        doc.save(output)
        b64 = base64.b64encode(output.getvalue()).decode()  # Convert to base64
        return f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="calendar_events.docx">Download Word</a>'

def main():
    st.title("ICS Calendar Parser")
    
    ics_url = st.text_input("Enter the ICS calendar link:")
    
    # Default date range
    current_year = datetime.now().year
    start_date = datetime(current_year, 9, 1)
    end_date = datetime(current_year + 1, 9, 1)
    
    # User inputs for date range
    st.sidebar.header("Filter Date Range")
    start_date = st.sidebar.date_input("Start Date", start_date)
    end_date = st.sidebar.date_input("End Date", end_date)
    
    # Convert to datetime objects
    start_date = datetime.combine(start_date, datetime.min.time())
    end_date = datetime.combine(end_date, datetime.min.time())
    
    if ics_url:
        calendar_data = fetch_ics_data(ics_url)
        
        if calendar_data:
            dates_from, dates_to, times, events = parse_ics_data(calendar_data, start_date, end_date)
            display_table(dates_from, dates_to, times, events)

if __name__ == "__main__":
    main()
