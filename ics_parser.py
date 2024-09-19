import streamlit as st
from ics import Calendar
import requests
import pandas as pd
from datetime import datetime, timedelta
import base64
from io import BytesIO
from docx import Document

def fetch_ics_data(ics_url):
    try:
        response = requests.get(ics_url)
        response.raise_for_status()  # Check if the request was successful
        return response.text
    except requests.exceptions.RequestException as e:
        st.error(f"Fehler beim Abrufen der ICS-Datei: {e}")
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

        # Determine if the end date should be shown
        if event_end and (event_end - event_start) > timedelta(days=1):
            dates_to.append(event_end.strftime("%d. %m. %Y"))
        else:
            dates_to.append("")  # Leave it empty if the duration is 1 day or less
        
        dates_from.append(event_start.strftime("%d. %m. %Y"))
        times.append(event_start.strftime("%H:%M"))
        events.append(event.name)

    return dates_from, dates_to, times, events

def display_table(dates_from, dates_to, times, events):
    table = {'Startdatum': dates_from, 'Enddatum': dates_to, 'Uhrzeit': times, 'Ereignis': events}
    df = pd.DataFrame(table)
    st.dataframe(df)
    
    # Add buttons to download the table in different formats
    st.markdown(get_table_download_link(df, 'csv'), unsafe_allow_html=True)
    st.markdown(get_table_download_link(df, 'xlsx'), unsafe_allow_html=True)
    st.markdown(get_table_download_link(df, 'docx'), unsafe_allow_html=True)

def get_table_download_link(df, file_type):
    if file_type == 'csv':
        csv = df.to_csv(index=False)
        b64 = base64.b64encode(csv.encode()).decode()  # Convert to base64
        return f'<a href="data:file/csv;base64,{b64}" download="kalender_ereignisse.csv">CSV herunterladen</a>'
    elif file_type == 'xlsx':
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ereignisse')
        b64 = base64.b64encode(output.getvalue()).decode()  # Convert to base64
        return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="kalender_ereignisse.xlsx">Excel herunterladen</a>'
    elif file_type == 'docx':
        doc = Document()
        doc.add_heading('Kalenderereignisse', 0)
        table = doc.add_table(rows=1, cols=len(df.columns))
        hdr_cells = table.rows[0].cells
        for i, column_name in enumerate(df.columns):
            hdr_cells[i].text = column_name

        for index, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

        output = BytesIO()
        doc.save(output)
        b64 = base64.b64encode(output.getvalue()).decode()  # Convert to base64
        return f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="kalender_ereignisse.docx">Word herunterladen</a>'

def main():
    st.title("ICS-Kalenderparser")
    
    ics_url = st.text_input("Geben Sie den ICS-Kalenderlink ein:")
    
    # Default date range
    current_year = datetime.now().year
    start_date = datetime(current_year, 9, 1)
    end_date = datetime(current_year + 1, 9, 1)
    
    # User inputs for date range
    st.sidebar.header("Filterdatum")
    start_date = st.sidebar.date_input("Startdatum", start_date)
    end_date = st.sidebar.date_input("Enddatum", end_date)
    
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
