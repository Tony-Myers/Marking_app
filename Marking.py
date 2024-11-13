import streamlit as st
from openai import OpenAI
import pandas as pd
import docx
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os
import json
from io import BytesIO

# Set your OpenAI API key and password from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]

# Instantiate the OpenAI client
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
    if check_password():
        st.title("🔐 Automated Assignment Grading and Feedback")

        st.header("Assignment Task")
        assignment_task = st.text_area("Enter the Assignment Task or Instructions (Optional)")

        st.header("Upload Files")
        rubric_file = st.file_uploader("Upload Grading Rubric (CSV)", type=['csv'])
        submissions = st.file_uploader("Upload Student Submissions (.docx)", type=['docx'], accept_multiple_files=True)

        if rubric_file and submissions:
            if st.button("Run Marking"):
                # Read and validate the grading rubric
                try:
                    original_rubric_df = pd.read_csv(rubric_file)
                    
                    # Define required columns
                    required_columns = ['Criterion', 'Max_Score', 'Description']
                    
                    # Check for required columns
                    for col in required_columns:
                        if col not in original_rubric_df.columns:
                            st.error(f"Rubric file must contain a '{col}' column.")
                            return
                    
                    # Ensure 'Max_Score' column is numeric
                    if not pd.api.types.is_numeric_dtype(original_rubric_df['Max_Score']):
                        st.error("The 'Max_Score' column must contain numeric values.")
                        return
                    
                    # Ensure no missing values in essential columns
                    if original_rubric_df[required_columns].isnull().any().any():
                        st.error("The rubric file contains missing values in required columns.")
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
                    You are an experienced educator tasked with grading student assignments based on the following rubric and assignment instructions. Provide feedback directly addressing the student (e.g., "You have demonstrated...") rather than speaking about them in third person (e.g., "The student has demonstrated...").

                    Rubric (in CSV format):
                    {rubric_csv_string}

                    Assignment Task:
                    {assignment_task}

                    Student's Submission:
                    {student_text}

                    Your responsibilities:

                    - Provide a completed grading rubric with scores and brief comments for each criterion, in JSON format, matching the rubric provided.
                    - Ensure that the JSON includes the keys '{criterion_column}', 'Score', and 'Comment' for each criterion.
                    - Write concise overall comments on the quality of the work, using language directly addressing the student.
                    - List actionable 'feedforward' bullet points for future improvement, also using direct language.

                    Please output in the following format:

                    Completed Grading Rubric (JSON):
                    [{{"Criterion": "Criterion 1", "Score": "Score 1", "Comment": "Comment 1"}}, 
                     {{"Criterion": "Criterion 2", "Score": "Score 2", "Comment": "Comment 2"}},
                     ... (continue for all criteria)]

                    Overall Comments:
                    [Text]

                    Feedforward:
                    [Bullet points]
                    """

                    # Call ChatGPT API
                    feedback = call_chatgpt(prompt, max_tokens=3000)
                    if feedback:
                        st.success(f"Feedback generated for {student_name}")
                        # Process and generate feedback document...


                    if feedback:
                        st.success(f"Feedback generated for {student_name}")

                        try:
                            # Check for JSON format in AI response
                            sections = feedback.split('Completed Grading Rubric (JSON):')
                            if len(sections) < 2:
                                st.error("Failed to parse the completed grading rubric from the AI response.")
                                st.write("AI Response:")
                                st.code(feedback)
                                continue
                            
                            rest = sections[1]
                            rubric_section, rest = rest.split('Overall Comments:', 1)
                            overall_comments_section, feedforward_section = rest.split('Feedforward:', 1)

                            # Check JSON structure
                            rubric_json = rubric_section.strip()
                            if not rubric_json.startswith("[") or not rubric_json.endswith("]"):
                                st.error("The rubric JSON format appears invalid.")
                                st.write("AI Response:")
                                st.code(feedback)
                                continue

                            completed_rubric_data = json.loads(rubric_json)
                            completed_rubric_df = pd.DataFrame(completed_rubric_data)

                            # Ensure 'Score' and 'Comment' columns are present
                            if 'Score' not in completed_rubric_df.columns or 'Comment' not in completed_rubric_df.columns:
                                st.error("The AI response is missing 'Score' or 'Comment' keys.")
                                st.write("AI Response:")
                                st.code(feedback)
                                continue

                            # Extract overall comments and feedforward
                            overall_comments = overall_comments_section.strip()
                            feedforward = feedforward_section.strip()

                            # Merge the original rubric with the completed rubric
                            merged_rubric_df = original_rubric_df.merge(
                                completed_rubric_df[[criterion_column, 'Score', 'Comment']],
                                on=criterion_column,
                                how='left'
                            )

                        except json.JSONDecodeError:
                            st.error("Failed to decode JSON in the AI response.")
                            st.write("AI Response:")
                            st.code(feedback)
                            continue
                        except Exception as e:
                            st.error(f"Error parsing AI response: {e}")
                            st.write("AI Response:")
                            st.code(feedback)
                            continue

                        # Create a Word document for the feedback
                        feedback_doc = docx.Document()

                        # Set the page orientation to landscape
                        section = feedback_doc.sections[0]
                        section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
                        section.page_width, section.page_height = section.page_height, section.page_width

                        # Add heading for the student's feedback
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
                            for _, row in merged_rubric_df.iterrows():
                                row_cells = table.add_row().cells
                                for i, col_name in enumerate(merged_rubric_df.columns):
                                    cell = row_cells[i]
                                    cell.text = str(row[col_name])
                                    if col_name in ['Score', 'Comment'] and pd.notnull(row[col_name]):
                                        shading_elm = parse_xml(r'<w:shd {} w:fill="D9EAD3"/>'.format(nsdecls('w')))
                                        cell._tc.get_or_add_tcPr().append(shading_elm)

                        # Add overall comments and feedforward
                        feedback_doc.add_heading('Overall Comments', level=2)
                        feedback_doc.add_paragraph(overall_comments)
                        feedback_doc.add_heading('Feedforward', level=2)
                        feedback_doc.add_paragraph(feedforward)

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
