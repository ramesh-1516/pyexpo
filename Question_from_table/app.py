from transformers import pipeline
import streamlit as st
import pandas as pd
import torch


st.set_page_config(page_title ="Ask me Anything!",page_icon="☁️")

def output(table,ques):
        tqa = pipeline(task="table-question-answering", 
               model="google/tapas-base-finetuned-wtq")
        return(tqa(table=table, query=ques)["answer"])


def main():
    html_temp = """
        <div style="background-color:tomato;padding:10px">
        <h1 style="color:white;text-align:center;">Ask Question from Table</h1>
        </div>
        """
    st.markdown(html_temp, unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload your File",type=['CSV'])
    query = st.text_input('Please Enter your Question!?')

    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        df = df.astype(str)
        st.write(df)
    if st.button('Ask!'):
        st.success(output(df,query)) 

if __name__=='__main__':
    main()