import streamlit as st
import pandas as pd
import time 
from datetime import datetime

ts=time.time()
date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
timestamp=datetime.fromtimestamp(ts).strftime("%H:%M-%S")

from streamlit_autorefresh import st_autorefresh

count = st_autorefresh(interval=2000, limit=100, key="fizzbuzzcounter")

if count == 0:
    st.write("Count is zero")
elif count % 3 == 0 and count % 5 == 0:
    st.write("FizzBuzz")
elif count % 3 == 0:
    st.write("Fizz")
elif count % 5 == 0:
    st.write("Buzz")
else:
    st.write(f"Count: {count}")

import os
# Check the current working directory (optional, for debugging)
st.write(f"Current Working Directory: {os.getcwd()}")


df=pd.read_csv("Attendance/Attendance_" + date + ".csv")

st.dataframe(df.style.highlight_max(axis=0))