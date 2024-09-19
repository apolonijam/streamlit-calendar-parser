import streamlit as st
from ics import Calendar
import requests
import pandas as pd

st.title("ICS Calendar Parser")

ics_url = st.text_input("Enter the ICS calendar link:")

if ics_url:
    try:
        response = requests.get(ics_url)
        response.raise_for_status()
        calendar_data = response.text
        
        calendar = Calendar(calendar_data)
        dates, times, events = [], [], []
        
        for event in calendar.events:
            start = event.begin
            dates.append(start.format("YYYY-MM-DD"))
            times.append(start.format("HH:mm"))
            events.append(event.name)
        
        table = {'Date': dates, 'Time': times, 'Event Name': events}
        df = pd.DataFrame(table)
        
        st.dataframe(df)
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching the ICS file: {e}")
