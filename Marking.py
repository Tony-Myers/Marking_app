import streamlit as st
import openai
import pandas as pd
import docx
import os
from io import BytesIO

# Set your OpenAI API key from secrets
PASSWORD = st.secrets["password"]
OPENAI_API_KEY = st.secrets["openai_api_key"]
openai.api_key = OPENAI_API_KEY  # Set the OpenAI API key

except Exception as e:
        return f"An error occurred in generate_response: {str(e)}"
def call_chatgpt(prompt, model="gpt-4o", max_tokens=500, temperature=0.7, retries=2):
    """Calls the OpenAI API and returns the response as text."""
    for attempt in range(retries):
        try:
            rresponse = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens,
                temperature=temperature,
                stop=None
            )
            return response['choices'][0]['message']['content'].strip()
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
- Completed grading rubric with scores and brief comments.
- Concise overall comments on the quality of the work.
- Actionable 'feedforward' bullet points for future improvement.
"""

                # Call ChatGPT API
                feedback = call_chatgpt(prompt)
                if feedback:
                    st.success(f"Feedback generated for {student_name}")

                    # Create a Word document for the feedback
                    feedback_doc = docx.Document()
                    feedback_doc.add_heading(f"Feedback for {student_name}", level=1)
                    feedback_doc.add_paragraph(feedback)

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
