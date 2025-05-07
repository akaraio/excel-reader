from langchain_community.document_loaders import  UnstructuredExcelLoader
from langchain_ollama import ChatOllama
from langchain_core.prompts import (SystemMessagePromptTemplate, 
                                    HumanMessagePromptTemplate,
                                    ChatPromptTemplate)
from langchain_core.output_parsers import StrOutputParser
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Reader")

if 'messages' not in st.session_state:
        st.session_state['messages'] = []

if 'context' not in st.session_state:
        st.session_state['context'] = []

system = SystemMessagePromptTemplate.from_template("""You are helpful AI assistant who answer user question based on the provided context with out any html code.""")
prompt = """Answer user question based on the provided context ONLY! If you do not know the answer, just say "I don't know".
            ### Context:
            {context}

            ### Question:
            {question}

            ### Answer:"""

llm = ChatOllama(model='data', base_url='data')
prompt = HumanMessagePromptTemplate.from_template(prompt)
messages = [system, prompt]
template = ChatPromptTemplate(messages)
qna_chain = template | llm | StrOutputParser()

def chat_with_llm(context, question):
    for event in qna_chain.stream({'context': context, 'question': question}):
        yield event

with st.sidebar:
    excel_doc = st.file_uploader("Upload your excel file", type=['xlsx'], key="excel_doc")
    if excel_doc:
        with st.spinner():
                pd_doc = pd.read_excel(excel_doc)
                pd_doc.to_excel(data)
                loader = UnstructuredExcelLoader(data,  mode="elements")
                excel = loader.load()
                context = excel[0].metadata['text_as_html']
                st.session_state['context'] = context
                st.success("Done")

def conversation():
        for message in st.session_state['messages']:
            with st.chat_message(message['role']):
                st.write(message['content'])

        if prompt := st.chat_input('Type here...'):
            st.session_state['messages'].append({'role': 'user', 'content': prompt})

            with st.chat_message('user'):
                st.write(prompt)
            
            with st.chat_message('assistant'):
                message = st.write_stream(chat_with_llm(st.session_state['context'], prompt))
                st.session_state['messages'].append({'role': 'assistant', 'content': message})

conversation() 