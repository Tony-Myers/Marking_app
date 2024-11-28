# Standard library imports
import streamlit as st
from openai import OpenAI
import pandas as pd
import docx
from PyPDF2 import PdfReader
from io import BytesIO, StringIO
import tiktoken
import csv
import re
import os
import time

# Additional imports for document handling
from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Set your OpenAI API key and password from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]

# Instantiate the OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

# Initialize the tiktoken encoder for GPT-4
try:
    encoding = tiktoken.encoding_for_model("gpt-4o")
except KeyError:
    encoding = tiktoken.get_encoding("cl100k_base")

MAX_TOKENS = 7000  # Maximum tokens for GPT-4
PROMPT_BUFFER = 1000  # Buffer to ensure we don't exceed the limit

# Additional utility functions
def count_tokens(text, encoding):
    """Counts the number of tokens in a given text."""
    return len(encoding.encode(text))

def truncate_text(text, max_tokens, encoding):
    """Truncates text to fit within max_tokens."""
    tokens = encoding.encode(text)
    if len(tokens) > max_tokens:
        truncated_tokens = tokens[:max_tokens]
        return encoding.decode(truncated_tokens)
    return text

def call_chatgpt(prompt, model="gpt-4o", max_tokens=3000, temperature=0.3, retries=2):
    """Calls the OpenAI API using the client instance and returns the response as text."""
    for attempt in range(retries):
        try:
            st.write(f"Attempt {attempt + 1} to contact OpenAI API...")  # Debug output
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens,
                temperature=temperature,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            st.error(f"API Error on attempt {attempt + 1}: {e}")
            time.sleep(2)  # Wait before retrying
            if attempt < retries - 1:
                continue
            else:
                return None

def check_password():
    """Prompts the user for a password and checks it."""
    def password_entered():
        if st.session_state["password"] == PASSWORD:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Enter the password", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Enter the password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
    else:
        return True

def parse_csv_section(csv_text):
    """Parses a CSV section line by line, handling quoted fields."""
    try:
        # Use StringIO to read the CSV text
        csv_io = StringIO(csv_text)
        # Read the CSV with proper quoting
        df = pd.read_csv(csv_io, dtype={'Criterion': str, 'Score': float}, quotechar='"', skipinitialspace=True)
        return df
    except Exception as e:
        st.error(f"Error parsing CSV: {e}")
        return None

def summarize_text(text):
    """Summarizes the given text using OpenAI API."""
    st.write("Calling ChatGPT to summarize text...")  # Debug output
    summary_prompt = f"""
You are an assistant that summarizes academic papers. Please provide a concise summary (max 500 words) of the following text:

{text}
"""
    summary = call_chatgpt(summary_prompt, max_tokens=800, temperature=0.3)
    return summary if summary else text  # Fallback to original text if summarization fails

def extract_weight(criterion_name):
    """
    Extracts the weight from the criterion name.
    For example, from "Linking Theory to Issue (15%)", it extracts 15.0
    """
    match = re.search(r'\((\d+)%\)', criterion_name)
    if match:
        return float(match.group(1))
    else:
        return 0.0  # Default weight if not found

def initialize_session_state():
    """Initializes session state for storing feedback."""
    if 'feedbacks' not in st.session_state:
        st.session_state['feedbacks'] = {}

# Functions to extract text from files
def extract_text_from_docx(docx_file):
    try:
        doc = docx.Document(docx_file)
        return '\n'.join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"Error reading DOCX file: {e}")
        return None

def extract_text_from_pdf(pdf_file):
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text
        return text
    except Exception as e:
        st.error(f"Error reading PDF file: {e}")
        return None

def extract_text_from_txt(txt_file):
    try:
        return txt_file.read().decode("utf-8")
    except Exception as e:
        st.error(f"Error reading TXT file: {e}")
        return None

def main():
    initialize_session_state()
    
    if check_password():
        st.title("✏️ Automated Assignment Grading and Feedback © Tony Myers")

        # Define the column name for criterion here to make it accessible throughout the function
        criterion_column = 'Criterion'

        st.header("Assignment Task and Level of Work")
        assignment_task = st.text_area("Enter the Assignment Task or Instructions, as well as the HE level the work should be marked at", height=150)

        st.header("Upload Files")
        rubric_file = st.file_uploader("Upload Grading Rubric (CSV)", type=['csv'])
        submissions = st.file_uploader("Upload Student Submissions (.docx, .pdf, .txt)", type=['docx', 'pdf', 'txt'], accept_multiple_files=True)

        if rubric_file and submissions:
            if st.button("Run Marking"):
                # Read the grading rubric
                try:
                    original_rubric_df = pd.read_csv(rubric_file, dtype={criterion_column: str})
                    st.write("Rubric file loaded successfully.")  # Debug output
                except Exception as e:
                    st.error(f"Error reading rubric: {e}")
                    return

                if criterion_column not in original_rubric_df.columns:
                    st.error(f"Rubric must contain a '{criterion_column}' column.")
                    return

                # Ensure Criterion column is string type for consistency in both dataframes
                original_rubric_df[criterion_column] = original_rubric_df[criterion_column].astype(str)

                # Extract Weight from Criterion and clean Criterion names
                original_rubric_df['Weight'] = original_rubric_df[criterion_column].apply(extract_weight)
                original_rubric_df[criterion_column] = original_rubric_df[criterion_column].apply(lambda x: re.sub(r'\s*\(\d+%\)', '', x))

                # Generate rubric CSV string with all fields quoted
                rubric_csv_string = original_rubric_df.to_csv(index=False, quoting=csv.QUOTE_ALL)

                # Process each student submission
                for submission in submissions:
                    student_name = os.path.splitext(submission.name)[0]

                    st.header(f"Processing {student_name}'s Submission")

                    # Read student submission
                    student_text = None
                    try:
                        if submission.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                            student_text = extract_text_from_docx(submission)
                        elif submission.type == "application/pdf":
                            student_text = extract_text_from_pdf(submission)
                        elif submission.type == "text/plain":
                            student_text = extract_text_from_txt(submission)
                        else:
                            st.error(f"Unsupported file type: {submission.type}")
                            continue

                        # If student_text extraction failed, skip this submission
                        if not student_text:
                            st.error(f"Failed to extract text from {submission.name}. Skipping.")
                            continue

                    except Exception as e:
                        st.error(f"Error reading submission {submission.name}: {e}")
                        continue

                    st.write(f"Successfully extracted text for {student_name}.")

                    # Summarize the student submission if it's too long
                    student_tokens = count_tokens(student_text, encoding)
                    max_submission_tokens = MAX_TOKENS - PROMPT_BUFFER  # Reserve tokens for prompt and other texts

                    if student_tokens > (MAX_TOKENS * 0.6):  # If submission is more than ~4k tokens
                        st.info(f"Summarizing {student_name}'s submission to reduce token count.")
                        student_text = summarize_text(student_text)
                        if not student_text:
                            st.error(f"Failed to summarize submission for {student_name}. Skipping.")
                            continue

                    st.write(f"Successfully summarized text for {student_name}.")

                    # Generate feedback with all promotional details
                    feedback_prompt = f"""
You are an experienced UK academic tasked with grading a student's assignment based on the provided rubric and assignment instructions. Please ensure that your feedback adheres to UK Higher Education standards for undergraduate work, noting the level provided by the user. Use British English spelling throughout your feedback.

**Instructions:**

- Review the student's submission thoroughly and be strict in applying the criteria.
- For **each criterion** in the list below, assign a numerical score between 0 and 100 (e.g., 75) and provide a brief but nuanced comment.
- Ensure that the score is numeric without any extra symbols or text.
- The scores should reflect the student's performance according to the descriptors in the rubric.
- **Be strict in your grading to align with UK undergraduate standards.**
- **Assess the quality of writing and referencing style, ensuring adherence to the 'Cite them Right' guidelines (2008, Pear Tree Books). Provide a brief comment on these aspects in the overall comments but refer to the referencing style as Birmingham Newman University’s referencing style in feedback.**

**List of Criteria:**
{criteria_string}

**Rubric (in CSV format):**
{rubric_csv_string}

**Assignment Task:**
{assignment_task}

**Student's Submission:**
{student_text}

**Your Output Format:**

- Begin your response with the CSV section, starting with "Criterion,Score,Comment".
- Include all criteria from the list provided.
- Provide overall comments and feedforward.
"""
                    feedback = call_chatgpt(feedback_prompt, max_tokens=1500)
                    if not feedback:
                        st.error(f"Failed to generate feedback for {student_name}. Skipping.")
                        continue

                    # Store feedback in session state
                    st.session_state['feedbacks'][student_name] = {
                        'feedback_text': feedback,
                    }

                    st.success(f"Finished processing {student_name}'s submission.")

        # After processing, provide download buttons for the feedback
        if st.session_state.get('feedbacks'):
            st.header("Generated Feedbacks")
            for student_name, feedback_data in st.session_state['feedbacks'].items():
                try:
                    feedback_doc = docx.Document()

                    feedback_doc.add_heading(f"Feedback for {student_name}", level=1)
                    feedback_doc.add_paragraph(feedback_data['feedback_text'])

                    # Create buffer to hold the document data
                    buffer = BytesIO()
                    feedback_doc.save(buffer)
                    buffer.seek(0)

                    # Provide download button for the feedback document
                    st.download_button(
                        label=f"Download Feedback for {student_name}",
                        data=buffer,
                        file_name=f"{student_name}_feedback.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Error creating feedback document for {student_name}: {e}")

if __name__ == "__main__":
    main()
