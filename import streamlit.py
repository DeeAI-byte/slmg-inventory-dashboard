import streamlit 
import streamlit as st  
import pandas as pd

<def lazy_export_section(df, filename, prefix):
st.header("Export Data")
<transformer = st.selectbox("Select a transformer", ["None", "Uppercase", "Lowercase"])
if transformer == "Uppercase":
    df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)
    elif transformer == "Lowercase":
    df = df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
    st.download_button

(
    label="Download CSV",
    data=df.to_csv(index=False).encode('utf-8'),
    file_name=f"{prefix}_{filename}.csv",
    mime='text/csv',
)
transformers = {}

def register_transformer(name, func):
    