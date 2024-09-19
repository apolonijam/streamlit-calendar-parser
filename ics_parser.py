import streamlit as st
from ics import Calendar
import requests
import pandas as pd
from datetime import datetime
import base64

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
    
    # Add button to copy table to clipboard
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)

def get_table_download_link(df):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()  # Convert to base64
    return f'<a href="data:file/csv;base64,{b64}" download="calendar_events.csv">Download CSV</a>'

def main():
    st.title("ICS Calendar Parser")
    
    ics_url = st.text_input("Enter the ICS calendar link:")
    
    # Define the date range for filtering events
    start_date = datetime(2024, 9, 1)
    end_date = datetime(2025, 9, 1)
    
    if ics_url:
        calendar_data = fetch_ics_data(ics_url)
        
        if calendar_data:
            dates_from, dates_to, times, events = parse_ics_data(calendar_data, start_date, end_date)
            display_table(dates_from, dates_to, times, events)

if __name__ == "__main__":
    main()
