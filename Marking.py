import pandas as pd
import streamlit as st
import docx
from openai import OpenAI
from io import StringIO, BytesIO
import os

# Load your API key
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]
client = OpenAI(api_key=OPENAI_API_KEY)

def call_chatgpt(prompt, model="gpt-3.5-turbo", max_tokens=3000, temperature=0.7, retries=2):
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


def main():
    st.title("🔐 Automated Assignment Grading and Feedback")
    rubric_file = st.file_uploader("Upload Grading Rubric (CSV)", type=['csv'])
    submissions = st.file_uploader("Upload Student Submissions (.docx)", type=['docx'], accept_multiple_files=True)

    if rubric_file and submissions:
        rubric_df = pd.read_csv(rubric_file)
        rubric_csv_string = rubric_df.to_csv(index=False)

        for submission in submissions:
            student_name = os.path.splitext(submission.name)[0]
            st.header(f"Processing {student_name}'s Submission")
            doc = docx.Document(submission)
            student_text = '\n'.join([para.text for para in doc.paragraphs])

            prompt = f"""
            You are an experienced educator tasked with grading based on the following rubric.
            Provide feedback in CSV format with columns 'Criterion', 'Score', 'Comment', then add Overall Comments.

            Rubric (CSV):
            {rubric_csv_string}

            Student's Submission:
            {student_text}
            """

            feedback = call_chatgpt(prompt, max_tokens=3000)

            if feedback:
                # Extract CSV section
                try:
                    csv_feedback = feedback.split('Overall Comments:')[0].strip()
                    comments = feedback.split('Overall Comments:')[1].strip()

                    # Load CSV from AI output
                    completed_rubric_df = pd.read_csv(StringIO(csv_feedback))

                    # Merge and save output
                    feedback_doc = docx.Document()
                    feedback_doc.add_heading(f"Feedback for {student_name}", level=1)

                    # Add rubric as a table
                    table = feedback_doc.add_table(rows=1, cols=len(completed_rubric_df.columns))
                    hdr_cells = table.rows[0].cells
                    for i, column in enumerate(completed_rubric_df.columns):
                        hdr_cells[i].text = column
                    for _, row in completed_rubric_df.iterrows():
                        row_cells = table.add_row().cells
                        for i, col_name in enumerate(completed_rubric_df.columns):
                            row_cells[i].text = str(row[col_name])

                    # Add overall comments
                    feedback_doc.add_heading('Overall Comments', level=2)
                    feedback_doc.add_paragraph(comments)

                    # Save and download
                    buffer = BytesIO()
                    feedback_doc.save(buffer)
                    buffer.seek(0)

                    st.download_button(
                        label=f"Download Feedback for {student_name}",
                        data=buffer,
                        file_name=f"{student_name}_feedback.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                except Exception as e:
                    st.error(f"Error parsing AI response: {e}")
                    st.code(feedback)

if __name__ == "__main__":
    main()

