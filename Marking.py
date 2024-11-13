import streamlit as st
from openai import OpenAI
import pandas as pd
import docx
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os
from io import BytesIO, StringIO

# Set your OpenAI API key and password from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]

# Instantiate the OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

def call_chatgpt(prompt, model="gpt-4o", max_tokens=3000, temperature=0.7, retries=2):
    """Calls the OpenAI API using the client instance and returns the response as text."""
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens,
                temperature=temperature,
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
                # Read the grading rubric
                try:
                    original_rubric_df = pd.read_csv(rubric_file)
                except Exception as e:
                    st.error(f"Error reading rubric: {e}")
                    return

                criterion_column = 'Criterion'  # Default column name
                if criterion_column not in original_rubric_df.columns:
                    st.error(f"Rubric must contain a '{criterion_column}' column.")
                    return

                rubric_csv_string = original_rubric_df.to_csv(index=False)

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
You are an experienced educator tasked with grading student assignments based on the following rubric and assignment instructions.

Rubric (in CSV format):
{rubric_csv_string}

Assignment Task:
{assignment_task}

Student's Submission:
{student_text}

Your responsibilities:
- Provide a completed grading rubric with scores and brief comments for each criterion, in CSV format, matching the rubric provided.
- Ensure the CSV includes the columns '{criterion_column}', 'Score', and 'Comment' for each criterion.
- Write concise overall comments on the quality of the work.
- List actionable 'feedforward' bullet points for future improvement.

Please output in the following format:

Criterion,Score,Comment
Criterion 1,Score 1,Comment 1
Criterion 2,Score 2,Comment 2
... (continue for all criteria)

Overall Comments:
[Text]

Feedforward:
[Bullet points]
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

                            # Clean and ensure lines have exactly three fields
                            csv_lines = []
                            for line in csv_feedback.splitlines():
                                if line.count(',') >= 2:
                                    fields = line.split(',', 2)  # Limit to 3 fields
                                    if len(fields) == 3:
                                        csv_lines.append(','.join(fields))

                            csv_feedback_cleaned = '\n'.join(csv_lines)

                            # Load the CSV section into DataFrame
                            completed_rubric_df = pd.read_csv(StringIO(csv_feedback_cleaned))
                            overall_comments, feedforward = comments_section.split('Feedforward:')

                            # Merge with original rubric if needed
                            merged_rubric_df = original_rubric_df.merge(
                                completed_rubric_df[[criterion_column, 'Score', 'Comment']],
                                on=criterion_column,
                                how='left'
                            )

                        except Exception as e:
                            st.error(f"Error parsing AI response: {e}")
                            st.write("AI Response:")
                            st.code(feedback)
                            continue

                        # Create Word document for feedback
                        feedback_doc = docx.Document()
                        feedback_doc.add_heading(f"Feedback for {student_name}", level=1)

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

                        feedback_doc.add_heading('Overall Comments', level=2)
                        feedback_doc.add_paragraph(overall_comments.strip())
                        feedback_doc.add_heading('Feedforward', level=2)
                        feedback_doc.add_paragraph(feedforward.strip())

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
    
