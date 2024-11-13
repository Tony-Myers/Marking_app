import streamlit as st
from openai import OpenAI
import pandas as pd
import docx
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os
from io import BytesIO
import io

# Set your OpenAI API key from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]

# Instantiate the OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

def call_chatgpt(prompt, model="gpt-3.5-turbo", max_tokens=500, temperature=0.7, retries=2):
    """Calls the OpenAI API using the client instance and returns the response as text."""
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens,
                temperature=temperature,
                stop=None
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            if attempt < retries - 1:
                continue
            else:
                st.error(f"API Error: {e}")
                return None

def check_password():
    """Prompts the user for a password and checks it."""
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == PASSWORD:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Remove password from session state
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input("Enter the password", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input("Enter the password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

def main():
    if check_password():
        st.title("🔐 Automated Assignment Grading and Feedback")

        st.sidebar.header("Upload Files")
        rubric_file = st.sidebar.file_uploader("Upload Grading Rubric (CSV)", type=['csv'])
        submissions = st.sidebar.file_uploader("Upload Student Submissions (.docx)", type=['docx'], accept_multiple_files=True)

        if rubric_file and submissions:
            if st.button("Run Marking"):
                # Read the grading rubric
                try:
                    rubric_df = pd.read_csv(rubric_file)
                except Exception as e:
                    st.error(f"Error reading rubric: {e}")
                    return

                # Process each submission
                for submission in submissions:
                    student_name = os.path.splitext(submission.name)[0]
                    st.header(f"Processing {student_name}'s Submission")

                    # Read student submission
                    try:
                        doc = docx.Document(submission)
                        student_text = '\n'.join([para.text for para in doc.paragraphs])
                    except Exception as e:
                        st.error(f"Error reading submission {submission.name}: {e}")
                        continue

                    # Prepare prompt for ChatGPT
                    prompt = f"""
You are an assistant that grades student assignments based on the following rubric:

{rubric_df.to_csv(index=False)}

Student's submission:

{student_text}

Provide:

- Completed grading rubric with scores and brief comments, in CSV format.

- Concise overall comments on the quality of the work.

- Actionable 'feedforward' bullet points for future improvement.

Please output in the following format:

Completed Grading Rubric (CSV):
[CSV data]

Overall Comments:
[Text]

Feedforward:
[Bullet points]
"""

                    # Call ChatGPT API
                    feedback = call_chatgpt(prompt)
                    if feedback:
                        st.success(f"Feedback generated for {student_name}")

                        # Parse the feedback
                        try:
                            # Split the feedback into sections
                            sections = feedback.split('Completed Grading Rubric (CSV):')
                 
