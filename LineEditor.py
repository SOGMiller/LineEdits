import streamlit as st
from docx import Document
from docx.shared import RGBColor
import openai
import io
from pathlib import Path
import time

class GrammarChecker:
    def __init__(self, api_key: str):
        """Initialize the grammar checker with OpenAI API key."""
        self.client = openai.OpenAI(api_key=api_key)
    
    def check_sentence(self, sentence: str) -> dict:
        """Check a sentence for grammar and spelling issues using GPT-4.
        
        Args:
            sentence (str): Sentence to check
            
        Returns:
            dict: Results including original text, corrections, and explanation
        """
        try:
            prompt = f"""Review this sentence for grammar and spelling:

"{sentence}"

If there are any errors, provide the correction. If there are no errors, indicate that it's correct.

Format your response exactly as:
Original: [Original sentence]
Corrected: [Corrected sentence or "No corrections needed"]
Explanation: [Brief explanation of changes or "No changes necessary"]"""

            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert proofreader and grammar checker."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            
            # Parse response into dictionary
            result = {}
            for line in response.choices[0].message.content.strip().split('\n'):
                key, value = line.split(': ', 1)
                result[key.lower()] = value.strip()
            
            return result
            
        except Exception as e:
            return {
                "original": sentence,
                "corrected": "Error during check",
                "explanation": f"Error: {str(e)}"
            }
    
    def split_into_sentences(self, text: str) -> list:
        """Split text into sentences using GPT-4 for accurate sentence boundary detection."""
        try:
            prompt = f"""Split this text into sentences, maintaining original formatting:

{text}

Return only the sentences, one per line."""

            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "Split text into sentences accurately."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            
            return [s.strip() for s in response.choices[0].message.content.strip().split('\n') if s.strip()]
            
        except Exception as e:
            # Fallback to simple period splitting if API call fails
            return [s.strip() for s in text.split('.') if s.strip()]
    
    def create_correction_doc(self, filename: str, corrections: list) -> Document:
        """Create document with original text and corrections.
        
        Args:
            filename (str): Original filename
            corrections (list): List of correction dictionaries
            
        Returns:
            Document: New document with corrections
        """
        doc = Document()
        
        # Add header
        doc.add_heading(f'Grammar and Spelling Check: {Path(filename).stem}', 0)
        
        # Add introduction
        doc.add_paragraph('Below are the results of the grammar and spelling check. Original sentences are shown with their corrections and explanations.')
        
        # Add a line for visual separation
        doc.add_paragraph('_' * 50)
        
        # Add corrections
        for i, correction in enumerate(corrections, 1):
            # Add section number
            p = doc.add_paragraph(f'Section {i}:')
            p.add_run('\nOriginal: ').bold = True
            p.add_run(correction['original'])
            
            p = doc.add_paragraph()
            p.add_run('Corrected: ').bold = True
            if correction['corrected'] != "No corrections needed":
                run = p.add_run(correction['corrected'])
                font = run.font
                font.color.rgb = RGBColor(0, 128, 0)  # Green color for corrections
            else:
                p.add_run(correction['corrected'])
            
            p = doc.add_paragraph()
            p.add_run('Explanation: ').bold = True
            p.add_run(correction['explanation'])
            
            # Add separator between sections
            doc.add_paragraph('_' * 50)
        
        return doc

def main():
    st.title("Document Grammar and Spelling Checker")
    
    # Initialize session state
    if 'processed_doc' not in st.session_state:
        st.session_state.processed_doc = None
    
    # API Key input
    api_key = st.text_input(
        "Enter your OpenAI API key:",
        type="password",
        help="Your API key will be used to access GPT-4 for grammar checking"
    )
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Upload Document",
        type=['docx'],
        help="Select a Word document to check"
    )
    
    if uploaded_file and api_key and st.button("Check Document"):
        try:
            # Initialize checker
            checker = GrammarChecker(api_key)
            
            # Show progress elements
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Calculate total steps for progress bar
            # Step 1: Reading document (5%)
            # Step 2: Splitting sentences (10%)
            # Step 3: Checking sentences (75%)
            # Step 4: Generating final document (10%)
            
            # Step 1: Read document (5%)
            status_text.text("Reading document...")
            progress_bar.progress(0.02)
            doc = Document(uploaded_file)
            progress_bar.progress(0.05)
            
            # Extract text from paragraphs
            full_text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
            
            # Step 2: Split into sentences (10%)
            status_text.text("Splitting text into sentences...")
            progress_bar.progress(0.07)
            sentences = checker.split_into_sentences(full_text)
            progress_bar.progress(0.15)
            
            # Step 3: Check each sentence (75% of total progress)
            corrections = []
            sentence_progress_weight = 0.75  # 75% of total progress
            progress_per_sentence = sentence_progress_weight / len(sentences)
            current_progress = 0.15  # Starting from previous progress
            
            for i, sentence in enumerate(sentences):
                status_text.text(f"Checking sentence {i+1} of {len(sentences)}...")
                
                result = checker.check_sentence(sentence)
                corrections.append(result)
                
                # Update progress
                current_progress += progress_per_sentence
                progress_bar.progress(current_progress)
                time.sleep(0.5)  # Reduced sleep time for better UX
            
            # Step 4: Generate final document (10%)
            status_text.text("Generating correction document...")
            progress_bar.progress(0.90)
            
            output_doc = checker.create_correction_doc(uploaded_file.name, corrections)
            
            # Save to memory
            doc_stream = io.BytesIO()
            output_doc.save(doc_stream)
            doc_stream.seek(0)
            
            # Store in session state
            st.session_state.processed_doc = {
                'data': doc_stream.getvalue(),
                'filename': f"{Path(uploaded_file.name).stem}_Corrected.docx"
            }
            
            # Complete the progress
            progress_bar.progress(1.0)
            status_text.text("Processing complete!")
            
            # Show download button
            st.download_button(
                label="Download Corrected Document",
                data=st.session_state.processed_doc['data'],
                file_name=st.session_state.processed_doc['filename'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
