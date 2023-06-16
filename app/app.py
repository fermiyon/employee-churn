import streamlit as st
import pandas as pd
import numpy as np
import pickle
import openai
import os

def html_options(text=None, align="left", size=12, weight="normal", style="normal", color="#F4A460", bg_color=None, bg_size=16, on='main', to_link=None, image_width=None, image_height=None, image_source=None, image_bg_color=None):
    if on == 'main':
        st.markdown(f"""<div style="background-color:{bg_color};padding:{bg_size}px">
        <h2 style='text-align: {align}; font-size: {size}px; font-weight: {weight}; font-style: {style}; color: {color};'>{text} </h2>
        </div>""", unsafe_allow_html=True)
    elif on == 'side':
        st.sidebar.markdown(f"""<div style="background-color:{bg_color};padding:{bg_size}px">
        <h2 style='text-align: {align}; font-size: {size}px; font-weight: {weight}; font-style: {style}; color: {color};'>{text} </h2>
        </div>""", unsafe_allow_html=True)
    elif on == 'link':
        image_style = f"background-color:{image_bg_color};" if image_bg_color else ""
        st.markdown(f"""<div style="text-align: {align};"><img width="{image_width}" height="{image_height}" src="{image_source}" style="{image_style}" /></a></div>""", unsafe_allow_html=True)

# HEAD TO PICTURE
st.image("images/Employee churn prediction.png")
st.write('')
st.write('')


# SIDEBAR 
with st.sidebar:
    st.image("images/Employee Information.png")
    st.write('')
    Departments = st.sidebar.selectbox('Department', ['Select','Sales', 'Technical', 'Support', 'IT', 'Research and Development','Product Manager', 'Marketing', "Accounting", "Human Resources", "Management", "Others"])
    st.write('')
    salary = st.sidebar.selectbox('Salary Status', ['Select','Low', 'Medium', 'High'])
    st.write('')
    satisfaction_level = st.sidebar.slider('Satisfaction Level', min_value=0.0, max_value=1.0, value=0.09, step=0.01)
    st.write('')
    last_evaluation = st.sidebar.slider('Last Evaluation(from Employer)', min_value=0.0, max_value=1.0, value=0.79, step=0.01)
    st.write('')
    number_project = st.sidebar.slider('Assigned Project', min_value=0, max_value=10, value=6, step=1)
    st.write('')
    average_montly_hours = st.sidebar.slider('Monthly Working Time(Hour)', min_value=50, max_value=400, value=293, step=1)
    st.write('')
    time_spend_company = st.sidebar.slider('Time in the Company(Year)', min_value=1, max_value=10, value=5, step=1)
    st.write('')
    Work_accident = st.sidebar.radio('Work Accident', ('True', 'False'))
    st.write('')
    promotion_last_5years = st.sidebar.radio('Get Promoted(Last 5 Year)', ('True', 'False'))


# Create a dataframe using feature inputs
show_df = {'Informations':{'Departments': Departments,
        'Salary': salary,
        'Satisfaction Level': satisfaction_level,
        'Last Evaluation': last_evaluation,
        'Assigned Project': number_project,
        'Monthly Working Time': average_montly_hours,
        'Time in the Company': time_spend_company,
        'Work Accident': Work_accident,
        'Get Promoted': promotion_last_5years}}


# LOAD MODEL

model_df = pd.DataFrame(data=[[satisfaction_level, last_evaluation, number_project, 
                                average_montly_hours, time_spend_company, Work_accident, 
                                promotion_last_5years, Departments, salary]],
                        columns=['satisfaction_level', 'last_evaluation', 'number_project', 
                                'average_montly_hours', 'time_spend_company', 'Work_accident', 
                                'promotion_last_5years', 'departments', 'salary'])
depts_map = {"Sales":"sales" , "Technical":"technical" , "Support":"support" , "IT":"IT" , 
       "Research and Development":"RandD" , "Product Manager":"product_mng" , "Marketing":"marketing" , 
        "Accounting":"accounting" , "Human Resources":"hr" , "Management":"management", "Select":"Select", "Others":"Others"}
salary_map = {"Low":1 , "Medium":2 , "High":0, "Select":"Select", "Other":"Others"}
bool_map = {"False":0, 'True':1}

model_df['promotion_last_5years'] = model_df['promotion_last_5years'].map(bool_map)
model_df['Work_accident'] = model_df['Work_accident'].map(bool_map)
model_df['departments'] = model_df['departments'].map(depts_map)
model_df['salary'] = model_df['salary'].map(salary_map)
model_churn = pickle.load(open('emp_churn_final_model', 'rb'))



# Employee TABLE
html_options(text='Employee', size=40, weight='bold', color='#FF4B4B', align='center')
st.write('')
st.table(show_df)

# ChatGPT
with open("openai_api.txt") as file:
    openai.api_key = file.read()


messages = [{"role":"system", "content":"Write like a world-known expert HR manager. Do not mention that you are HR manager."},]

def AdviceGPT():
    with st.spinner('⚡Gearing up to blow your mind...'):
        leave_text = ''
        if result == 1:
            leave_text = f'This employee is churn according to ml model with {result_proba[1]} score'
        else:
            leave_text = f'This employee is not churn according to ml model with {result_proba[0]} score'
        show_df['Informations']['Monthly Working Time'] = str(show_df['Informations']['Monthly Working Time'])  + ' hours'
        message = f"How can I increase the productivity of this employee? Employee information: {show_df}. {leave_text}. Consider employee information and evaluate each information. Include numbers and values. Also comment on churn with the ML score rounded. Write engaging conclusion."
        if message:
            messages.append({"role":"system", "content":message})
            response = openai.ChatCompletion.create(
                        model = "gpt-3.5-turbo",
                        messages = messages
        )
        gpt_reply = response["choices"][0]["message"]["content"]
        messages.append({"role":"system", "content":gpt_reply})
    st.success(gpt_reply)


def CustomGPT():

    message = st.text_input("")

    if message:
        with open("messages.txt", "a") as file:
            file.write(message+"\n")
        messages.append({"role":"system", "content":message})
        response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=messages
            )
        gpt_reply = response["choices"][0]["message"]["content"]

        with open("replys.txt", "a") as file:
            file.write(gpt_reply+ "\n")

        messages.append({"role":"system", "content":gpt_reply})


# PREDICTION
result = np.nan
result_proba = np.nan
predict = st.button("Predict",type='primary', use_container_width=True)
if predict:
    if Departments == "Select" or salary == 'Select':
        st.warning('Be sure to enter department and salary information.')
    else:
        result = model_churn.predict(model_df)[0]
        result_proba = model_churn.predict_proba(model_df)[0]
        if result == 1:
            st.image("images/true.png")
            st.write("")
            AdviceGPT()
        else:
            st.image("images/false.png")
            st.write("")
            AdviceGPT()


st.write("")
st.write("")
st.write("")
st.write("")
st.write("")
st.image('images/contact.png',width=200)



CustomGPT()
with open("messages.txt", "r") as file:
    mes = file.read()
    st.info("User: "+mes.strip().split('\n')[-1])
with open("replys.txt", "r") as file:
    rep = file.read()
    st.info("Customer Support: "+rep.strip().split('\n')[-1])