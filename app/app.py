import streamlit as st
import pandas as pd
import numpy as np
import pickle
import openai
import scipy.stats as stats
import os

# Creating PDF
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors

# Creating Word
from docx import Document
from docx.shared import Inches

# To add date to the title of PDF file
from datetime import datetime

# Importing df
df = pd.read_csv('HR_dataset.csv')
df.drop_duplicates(inplace=True)
df.reset_index(drop=True,inplace=True)
df.rename(columns={'Departments ':'departments'}, inplace = True)
df_statistical_test = df.drop(columns=['left','Work_accident','promotion_last_5years'])


def generate_pdf(text: str, information_dict: dict, filename: str = 'output.pdf') -> SimpleDocTemplate:
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()
    style = styles["Normal"]
    style.alignment = 4  # Justify alignment
    style.leading = 14  # Line spacing within paragraphs

    title_style = styles["Title"]
    title = Paragraph("<font size='20'>Report</font>", title_style)

    # Create information dictionary table
    data = []
    for key, value in information_dict.items():
        data.append([key + ":", value])

    # Calculate the width of the paragraph
    paragraph_width = doc.width
    paragraph = Paragraph(text, style)
    paragraph_width = paragraph.wrap(paragraph_width, doc.topMargin)[0]

    # Create the table with the calculated width
    table = Table(data, hAlign='LEFT', colWidths=[paragraph_width * 0.49, paragraph_width * 0.49])
    table.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    # Add title and information dictionary table to flowables
    flowables = [title, Spacer(1, 24), table, Spacer(1, 24)]
    paragraphs = [Paragraph("<font size='12'>" + p.strip() + "</font>", style) for p in text.split("\n\n")]
    for p in paragraphs:
        p.alignment = 4  # Justify alignment
        flowables.append(p)
        flowables.append(Spacer(1, 10))  # Line spacing between paragraphs

    doc.build(flowables)
    return doc


def generate_filename() -> str:
    """
    Generates a filename for an employee churn report based on the current date and time.

    Returns:
        str: The filename in the format 'employee_churn_YYYY-MM-DD_HH-MM-SS.pdf'.
    """
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    filename = f'employee_churn_{timestamp}.pdf'
    return filename

def calculate_department_stats(data, sample_df, left=None):
    """
    Calculates the mean values of several employee performance metrics for a specific department in a dataframe.
    
    Args:
        data (pandas.DataFrame): The dataframe containing employee data. 
        sample_df (pandas.DataFrame): A separate dataframe containing information about the department of interest.
        left (bool or None): If left is None, calculate stats for all employees in the department (both left and not left).
                             If left is True, calculate stats only for employees who have left the company.
                             If left is False, calculate stats only for employees who have not yet left the company.
                             
    Returns:
        dict: A dictionary containing the mean values of the following employee performance metrics for the specified department: 
            - satisfaction_level     
            - last_evaluation
            - number_project       
            - average_montly_hours
            - time_spend_company
    """
    
    filtered_df = np.nan
    department = sample_df.departments.iloc[0]
    if left is None:
        filtered_df = data[data.departments == department]   
    else:
        filtered_df = data[(data.departments == department) & (data.left == int(left))]
        
    metrics = ['satisfaction_level', 'last_evaluation', 'number_project', 'average_montly_hours', 'time_spend_company']
    stats_dict = filtered_df[metrics].mean().to_dict()
        
    return stats_dict

def explain_department_stats(stats_dict, department_name, left=None):
    """
    Generates a string explaining the meaning of the values in a dictionary of department statistics.

    Args:
        stats_dict (dict): A dictionary containing the mean values of several employee performance metrics for a department.
        department_name (str): The name of the department the stats_dict corresponds to.
        left (bool or None): If left is None, generate an explanation for all employees in the department (both left and not left).
                             If left is True, generate an explanation only for employees who have left the company.
                             If left is False, generate an explanation only for employees who have not yet left the company.

    Returns:
        str: A string explaining the meaning of the values in the stats_dict dictionary.
    """
    if left is None:
        explanation = f"These are the mean values for the {department_name} department:"
    elif left:
        explanation = f"These are the mean values for employees who is churn of the {department_name} department:"
    else:
        explanation = f"These are the mean values for employees who is not churn of the {department_name} department:"
    for metric, value in stats_dict.items():
        explanation += f" {metric.replace('_', ' ')}: {value:.2f}. "
    # explanation += f"The employee is from this department."
    return explanation

statistical_findings = [
  "There is a statistically significant difference in average values between employees who left and those who stayed for column satisfaction_level.",
  "There is no statistically significant difference in average values between employees who left and those who stayed for column last_evaluation.",
  "There is no statistically significant difference in average values between employees who left and those who stayed for column number_project.",
  "There is a statistically significant difference in average values between employees who left and those who stayed for column average_montly_hours.",
  "There is a statistically significant difference in average values between employees who left and those who stayed for column time_spend_company.",
  "There is a statistically significant difference in average values between employees who left and those who stayed for column Work_accident.",
  "There is a statistically significant difference in average values between employees who left and those who stayed for column promotion_last_5years.",
  "There is a statistically significant difference in average values between employees who left and those who stayed for column left.",
  "There is evidence to suggest a significant difference in the proportion of employees who left the company based on whether they had a work accident or not.",
  "There is a statistically significant difference in the average satisfaction level between employees who had a work accident and those who didn't.",
  "There is a statistically significant association between the salary level of employees and the likelihood of them leaving the company."
]


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

# Encoding
depts_map = {"Sales":"sales" , "Technical":"technical" , "Support":"support" , "IT":"IT" , 
       "Research and Development":"RandD" , "Product Manager":"product_mng" , "Marketing":"marketing" , 
        "Accounting":"accounting" , "Human Resources":"hr" , "Management":"management", "Select":"Select", "Others":"Others"}
salary_map = {"Low":'low' , "Medium":'medium', "High":'high', "Select":"Select", "Other":"Others"}
bool_map = {"False":0, 'True':1}

model_df['promotion_last_5years'] = model_df['promotion_last_5years'].map(bool_map)
model_df['Work_accident'] = model_df['Work_accident'].map(bool_map)
model_df['departments'] = model_df['departments'].map(depts_map)
model_df['salary'] = model_df['salary'].map(salary_map)

# Load model
model_churn = pickle.load(open('emp_churn_final_model', 'rb'))



# Employee TABLE
html_options(text='Employee', size=40, weight='bold', color='#FF4B4B', align='center')
st.write('')
st.table(show_df)

# Setting up OpenAI API key
with open("openai_api.txt") as file:
    openai.api_key = file.read()


messages =  [{"role": "system", "content": "You are a world-known expert analyst. Do not mention that you are an analyst, ever. Use explicit numerical values in parantheses."}]

def AdviceGPT():
    with st.spinner('⚡ Gearing up to blow your mind... '):
        leave_text = ''
        if result == 1:
            leave_text = f'This employee is churn according to ml model with {result_proba[1]} score'
            department_info = explain_department_stats(calculate_department_stats(df,model_df,1),model_df.departments.iloc[0],1)
            department_info += explain_department_stats(calculate_department_stats(df,model_df,0),model_df.departments.iloc[0],0)
            department_info += explain_department_stats(calculate_department_stats(df,model_df),model_df.departments.iloc[0])
        else:
            leave_text = f'This employee is not churn according to ml model with {result_proba[0]} score'
            department_info = explain_department_stats(calculate_department_stats(df,model_df,1),model_df.departments.iloc[0],1)
            department_info += explain_department_stats(calculate_department_stats(df,model_df,0),model_df.departments.iloc[0],0)
            department_info += explain_department_stats(calculate_department_stats(df,model_df),model_df.departments.iloc[0])

        show_df['Informations']['Monthly Working Time'] = str(show_df['Informations']['Monthly Working Time'])  + ' hours'
        message = f"Write an analysis on that regarding churn analysis? Employee information: {show_df}. {leave_text}. These are statistical test results based on hypothesis tests:{' '.join(statistical_findings)} {department_info} Start with an introduction. Consider employee information and evaluate each information in a seperate paragraph. Also comment on churn with the ML score rounded. Include also numbers for means. In one paragraph answer the question of How can I increase the productivity of this employee? Write engaging conclusion."
        
        if message:
            messages.append({"role":"user", "content":message})
            completion = openai.ChatCompletion.create(
                        model = "gpt-3.5-turbo",
                        messages = messages
        )
        gpt_reply = completion["choices"][0].message.content
        messages.append({"role":"system", "content":gpt_reply})
    st.success(gpt_reply)

    filename = generate_filename()
    generate_pdf(gpt_reply,show_df['Informations'],filename)
    with open(filename, "rb") as f:
        st.download_button("Download PDF", f.read(), file_name=filename, mime="application/pdf")


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



CustomGPT()
with open("messages.txt", "r") as file:
    mes = file.read()
    st.info("User: "+mes.strip().split('\n')[-1])
with open("replys.txt", "r") as file:
    rep = file.read()
    st.info("Customer Support: "+rep.strip().split('\n')[-1])