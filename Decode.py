import base64
import streamlit as st
import zipfile
import os
import tempfile
import shutil

data = open("C:\Abb.docx", "rb").read()
encoded = base64.b64encode(data)
decoded = base64.b64decode(encoded)
st.download_button('Download Here', decoded, "C:\ALMATY2023\IPYNB2023\decoded_file.docx")
