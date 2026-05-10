import google.generativeai as genai
import streamlit as st

api_key = st.secrets["GOOGLE_API_KEY"]["GOOGLE_API_KEY"]
genai.configure(api_key=api_key)

print("Available models:")
for m in genai.list_models():
    print(m.name)
