import streamlit as st
from transformers import pipeline

st.set_page_config(page_title ="QA Gen",page_icon="☁️")

@st.cache(allow_output_mutation=True)
def load_qa_model():
    model = pipeline("question-answering")
    return model

qa = load_qa_model()
html_temp = """
        <div style="background-color:skyblue;padding:10px">
        <h1 style="color:white;text-align:center;">Ask Question from Table</h1>
        </div>
        """
st.markdown(html_temp, unsafe_allow_html=True)
sentence = st.text_area('Please paste your article :', height=30)
question = st.text_input("Questions from this article?")
button = st.button("Get me Answers")
with st.spinner("Discovering Answers.."):
    if button and sentence:
        answers = qa(question=question, context=sentence)
        st.success(answers['answer'])