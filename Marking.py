import streamlit as st
from openai import OpenAI
import pandas as pd
import docx
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os
from io import BytesIO, StringIO

# Set your OpenAI API key and password from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]

# Instantiate the OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

def call_chatgpt(prompt, model="gpt-4", max_tokens=3000, temperature=0.5, retries=2):
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
            if attempt < retries -1:
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

def parse_csv_section(csv_text):
    """Parses a CSV section line by line."""
    return csv_text.strip()

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
                # Define the column name for criterion before usage
                criterion_column = 'Criterion'

                # Read the grading rubric
                try:
                    original_rubric_df = pd.read_csv(rubric_file, dtype={criterion_column: str})
                except Exception as e:
                    st.error(f"Error reading rubric: {e}")
                    return

                if criterion_column not in original_rubric_df.columns:
                    st.error(f"Rubric must contain a '{criterion_column}' column.")
                    return

                # Ensure Criterion column is string type for consistency in both dataframes
                original_rubric_df[criterion_column] = original_rubric_df[criterion_column].astype(str)
                rubric_csv_string = original_rubric_df.to_csv(index=False)

                # Get the list of percentage range columns
                percentage_columns = [col for col in original_rubric_df.columns if '%' in col]

                # Get the list of criteria
                criteria_list = original_rubric_df[criterion_column].tolist()
                criteria_string = '\n'.join(criteria_list)

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
You are an experienced educator tasked with grading a student's assignment based on the provided rubric and assignment instructions.

**Instructions:**

- Review the student's submission thoroughly.
- For **each criterion** in the list below, assign a numerical score between 0 and 100 (e.g., 75) and provide a brief comment.
- Ensure that the score is numeric without any extra symbols or text.
- The scores should reflect the student's performance according to the descriptors in the rubric.

**List of Criteria:**
{criteria_string}

**Rubric (in CSV format):**
{rubric_csv_string}

**Assignment Task:**
{assignment_task}

**Student's Submission:**
{student_text}

**Your Output Format:**

Please output your feedback in the exact format below, ensuring you include **all criteria**:


**Important Notes:**

- **Ensure that 'Overall Comments:' and 'Feedforward:' are included exactly as shown, with the colon and on separate lines.**
- **Do not include any additional text outside of the specified format.**
- **Do not omit any sections.**
- **Do not use markdown or bullet points in the CSV section.**

"""

                    # Call ChatGPT API
                    feedback = call_chatgpt(prompt, max_tokens=3000)
                    if feedback:
                        st.success(f"Feedback generated for {student_name}")

                        # Display AI's response for debugging
                        st.text_area("AI Response", feedback, height=300)

                        # Parse the feedback
                        try:
                            # Split the feedback into CSV and comments sections
                            if 'Overall Comments:' in feedback:
                                csv_feedback = feedback.split('Overall Comments:', 1)[0].strip()
                                comments_section = feedback.split('Overall Comments:', 1)[1].strip()
                            else:
                                st.error("The AI's response is missing 'Overall Comments:'. Please adjust the prompt or try again.")
                                st.write("AI Response:")
                                st.code(feedback)
                                continue

                            csv_feedback_cleaned = parse_csv_section(csv_feedback)

                            # Load the cleaned CSV section into DataFrame
                            completed_rubric_df = pd.read_csv(StringIO(csv_feedback_cleaned), dtype={criterion_column: str, 'Score': float})

                            # Ensure all criteria are present
                            missing_criteria = set(original_rubric_df[criterion_column]) - set(completed_rubric_df[criterion_column])
                            if missing_criteria:
                                st.warning(f"The AI feedback is missing the following criteria: {missing_criteria}")

                            # Extract overall comments and feedforward
                            overall_comments = ''
                            feedforward = ''
                            if 'Feedforward:' in comments_section:
                                overall_comments = comments_section.split('Feedforward:', 1)[0].strip()
                                feedforward = comments_section.split('Feedforward:', 1)[1].strip()
                            else:
                                st.warning("The AI's response is missing 'Feedforward:'. Please adjust the prompt or try again.")
                                overall_comments = comments_section.strip()
                                feedforward = ''

                            # Merge the dataframes with specified suffixes
                            merged_rubric_df = original_rubric_df.merge(
                                completed_rubric_df[[criterion_column, 'Score', 'Comment']],
                                on=criterion_column,
                                how='left',
                                suffixes=('', '_ai')  # Suffix for AI feedback columns
                            ).dropna(how="all", axis=1)

                            # Display DataFrames for debugging
                            st.write("Completed Rubric DataFrame:")
                            st.dataframe(completed_rubric_df)

                            st.write("Merged Rubric DataFrame:")
                            st.dataframe(merged_rubric_df)

                        except Exception as e:
                            st.error(f"Error parsing AI response: {e}")
                            st.write("AI Response:")
                            st.code(feedback)
                            continue

                        # Create Word document for feedback
                        feedback_doc = docx.Document()

                        # Set page to landscape
                        section = feedback_doc.sections[0]
                        section.orientation = WD_ORIENT.LANDSCAPE
                        new_width, new_height = section.page_height, section.page_width
                        section.page_width = new_width
                        section.page_height = new_height

                        feedback_doc.add_heading(f"Feedback for {student_name}", level=1)

                        if not merged_rubric_df.empty:
                            # Prepare columns for the Word table
                            table_columns = [criterion_column] + percentage_columns + ['Score', 'Comment']
                            table = feedback_doc.add_table(rows=1, cols=len(table_columns))
                            table.style = 'Table Grid'
                            hdr_cells = table.rows[0].cells
                            for i, column in enumerate(table_columns):
                                hdr_cells[i].text = str(column)

                            # Add data rows and apply shading to the appropriate descriptor cell
                            for _, row in merged_rubric_df.iterrows():
                                row_cells = table.add_row().cells
                                score = row['Score']
                                for i, col_name in enumerate(table_columns):
                                    cell = row_cells[i]
                                    cell_text = str(row[col_name])
                                    cell.text = cell_text

                                    # Apply shading to the descriptor cell matching the score range
                                    if col_name in percentage_columns and pd.notnull(score):
                                        # Extract numeric values from the percentage range
                                        range_text = col_name.replace('%', '').strip()
                                        lower_upper = range_text.split('-')
                                        if len(lower_upper) == 2:
                                            try:
                                                lower = float(lower_upper[0].strip())
                                                upper = float(lower_upper[1].strip())

                                                # Use the score as a float
                                                score_value = float(score)

                                                if lower <= score_value <= upper:
                                                    # Apply green shading to this cell
                                                    shading_elm = parse_xml(r'<w:shd {} w:fill="D9EAD3"/>'.format(nsdecls('w')))
                                                    cell._tc.get_or_add_tcPr().append(shading_elm)
                                            except ValueError as e:
                                                st.warning(f"Error converting score or range to float: {e}")
                                                continue

                        # Add overall comments and feedforward
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
