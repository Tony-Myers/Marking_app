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

def call_chatgpt(prompt, model="gpt-4o", max_tokens=1500, temperature=0.7, retries=2):
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
                    original_rubric_df = pd.read_csv(rubric_file)
                except Exception as e:
                    st.error(f"Error reading rubric: {e}")
                    return

                # Ensure there is a unique identifier for criteria
                if 'Criterion' not in original_rubric_df.columns:
                    st.error("Rubric must contain a 'Criterion' column.")
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
You are an assistant that grades student assignments based on the following rubric:

{rubric_csv_string}

Student's submission:

{student_text}

Provide:

- Completed grading rubric with scores and brief comments, in CSV format, matching the rubric provided.

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
                    feedback = call_chatgpt(prompt, max_tokens=1500)
                    if feedback:
                        st.success(f"Feedback generated for {student_name}")

                        # Parse the feedback
                        try:
                            # Split the feedback into sections
                            sections = feedback.split('Completed Grading Rubric (CSV):')
                            if len(sections) < 2:
                                st.error("Failed to parse the completed grading rubric from the AI response.")
                                continue
                            rest = sections[1]
                            rubric_section, rest = rest.split('Overall Comments:', 1)
                            overall_comments_section, feedforward_section = rest.split('Feedforward:', 1)

                            # Read the completed rubric CSV
                            rubric_csv = io.StringIO(rubric_section.strip())
                            completed_rubric_df = pd.read_csv(rubric_csv)

                            # Get overall comments and feedforward
                            overall_comments = overall_comments_section.strip()
                            feedforward = feedforward_section.strip()

                            # Merge the original rubric with the completed rubric
                            merged_rubric_df = original_rubric_df.merge(
                                completed_rubric_df[['Criterion', 'Score', 'Comment']],
                                on='Criterion',
                                how='left'
                            )

                        except Exception as e:
                            st.error(f"Error parsing AI response: {e}")
                            continue

                        # Create a Word document for the feedback
                        feedback_doc = docx.Document()
                        feedback_doc.add_heading(f"Feedback for {student_name}", level=1)

                        # Add the full rubric as a table
                        if not merged_rubric_df.empty:
                            table = feedback_doc.add_table(rows=1, cols=len(merged_rubric_df.columns))
                            table.style = 'Table Grid'

                            # Add header row
                            hdr_cells = table.rows[0].cells
                            for i, column in enumerate(merged_rubric_df.columns):
                                hdr_cells[i].text = str(column)

                            # Add data rows
                            for index, row in merged_rubric_df.iterrows():
                                row_cells = table.add_row().cells
                                for i, cell in enumerate(row_cells):
                                    cell.text = str(row[i])

                                # Highlight rows where 'Score' is not null
                                if not pd.isnull(row['Score']):
                                    for cell in row_cells:
                                        shading_elm = parse_xml(r'<w:shd {} w:fill="D9EAD3"/>'.format(nsdecls('w')))
                                        cell._tc.get_or_add_tcPr().append(shading_elm)

                        # Add overall comments
                        feedback_doc.add_heading('Overall Comments', level=2)
                        feedback_doc.add_paragraph(overall_comments)

                        # Add feedforward
                        feedback_doc.add_heading('Feedforward', level=2)
                        feedback_doc.add_paragraph(feedforward)

                        # Save the feedback document to a buffer
                        buffer = BytesIO()
                        feedback_doc.save(buffer)
                        buffer.seek(0)

                        # Provide download link
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
