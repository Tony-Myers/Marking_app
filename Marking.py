import streamlit as st
from openai import OpenAI
import pandas as pd
import docx
from io import BytesIO, StringIO 
import os
import json

# Set your OpenAI API key and password from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]

# Instantiate the OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

def call_chatgpt(prompt, model="gpt-4", max_tokens=3000, temperature=0.7, retries=2):
    """Calls the OpenAI API using the client instance and returns the response as text."""
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens,
                temperature=temperature
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
    if check_password():
        st.title("🔐 Automated Assignment Grading and Feedback")

        st.header("Assignment Task")
        assignment_task = st.text_area("Enter the Assignment Task or Instructions (Optional)")

        st.header("Upload Files")
        rubric_file = st.file_uploader("Upload Grading Rubric (CSV)", type=['csv'])
        submissions = st.file_uploader("Upload Student Submissions (.docx)", type=['docx'], accept_multiple_files=True)

        if rubric_file and submissions:
            if st.button("Run Marking"):
                # Load and validate rubric
                try:
                    original_rubric_df = pd.read_csv(rubric_file, usecols=["Criterion", "comment"])
                    if original_rubric_df["Criterion"].isnull().any():
                        st.error("The 'Criterion' column in the rubric contains missing values.")
                        return
                except Exception as e:
                    st.error(f"Error reading rubric: {e}")
                    return

                # Convert rubric to CSV string for prompt
                rubric_csv_string = original_rubric_df.to_csv(index=False)

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
                    You are an experienced educator tasked with grading student assignments based on the following rubric and assignment instructions. Provide feedback directly addressing the student.

                    Rubric (CSV):
                    {rubric_csv_string}

                    Assignment Task:
                    {assignment_task}

                    Student's Submission:
                    {student_text}

                    Your responsibilities:
                    - Complete the grading rubric with specific scores and feedback for each criterion, updating the 'comment' column with relevant feedback.
                    - Write concise overall comments on the student's work.
                    - Provide actionable suggestions for future improvement.

                    Please output in CSV format with 'Criterion', 'Score', and 'comment' columns, followed by 'Overall Comments' and 'Feedforward' sections.
                    """

                    # Call ChatGPT API
                    feedback = call_chatgpt(prompt, max_tokens=3000)
                    if feedback:
                        st.success(f"Feedback generated for {student_name}")

                        # Parse the feedback
                        try:
                            # Split the feedback into CSV and comments sections
                            csv_feedback = feedback.split('Overall Comments:')[0].strip()
                            comments_section = feedback.split('Overall Comments:')[1].strip()

                            # Load the CSV section into DataFrame
                            completed_rubric_df = pd.read_csv(StringIO(csv_feedback))
                            overall_comments, feedforward = comments_section.split('Feedforward:')

                            # Merge with original rubric if needed
                            merged_rubric_df = original_rubric_df.merge(
                                completed_rubric_df[['Criterion', 'Score', 'comment']],
                                on='Criterion',
                                how='left'
                            )

                        except Exception as e:
                            st.error(f"Error parsing AI response: {e}")
                            st.write("AI Response:")
                            st.code(feedback)
                            continue

                        # Generate Word document for feedback
                        feedback_doc = docx.Document()
                        feedback_doc.add_heading(f"Feedback for {student_name}", level=1)

                        # Add the rubric as a table
                        if not merged_rubric_df.empty:
                            table = feedback_doc.add_table(rows=1, cols=len(merged_rubric_df.columns))
                            table.style = 'Table Grid'
                            hdr_cells = table.rows[0].cells
                            for i, column in enumerate(merged_rubric_df.columns):
                                hdr_cells[i].text = str(column)

                            for _, row in merged_rubric_df.iterrows():
                                row_cells = table.add_row().cells
                                for i, col_name in enumerate(merged_rubric_df.columns):
                                    row_cells[i].text = str(row[col_name])

                        # Add overall comments and feedforward
                        feedback_doc.add_heading('Overall Comments', level=2)
                        feedback_doc.add_paragraph(overall_comments.strip())
                        feedback_doc.add_heading('Feedforward', level=2)
                        feedback_doc.add_paragraph(feedforward.strip())

                        # Save and provide download link
                        buffer = BytesIO()
                        feedback_doc.save(buffer)
                        buffer.seek(0)
                        st.download_button(
                            label=f"Download Feedback for {student_name}",
                            data=buffer,
                            file_name=f"{student_name}_feedback.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error(f"Failed to generate feedback for {student_name}")

if __name__ == "__main__":
    main()
