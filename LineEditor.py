import streamlit as st
from docx import Document
from docx.shared import RGBColor
import openai
import io
from pathlib import Path
import time
import pandas as pd
import altair as alt
from collections import Counter
import re

# Add at the beginning of the file, after imports
st.set_page_config(
    page_title="Document Grammar Checker",
    page_icon="📝",
    layout="wide",  # This makes the layout wider
    initial_sidebar_state="expanded"
)

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

    def check_sentences_batch(self, sentences: list) -> list:
        """Check a batch of sentences for grammar and spelling issues using GPT-4.
        
        Args:
            sentences (list): List of sentences to check
            
        Returns:
            list: List of dictionaries containing results for each sentence
        """
        try:
            sentences_text = "\n".join([f"{i+1}. {sent}" for i, sent in enumerate(sentences)])
            prompt = f"""Review these sentences for grammar and spelling:

{sentences_text}

For each numbered sentence, provide corrections if needed.
Format your response exactly as follows for each sentence:

1.
Original: [Original sentence]
Corrected: [Corrected sentence or "No corrections needed"]
Explanation: [Brief explanation of changes or "No changes necessary"]

2.
[Continue for each sentence...]"""

            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert proofreader and grammar checker."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            
            # Parse response into list of dictionaries
            results = []
            current_result = {}
            
            for line in response.choices[0].message.content.strip().split('\n'):
                line = line.strip()
                if not line:
                    continue
                    
                # If line starts with a number and period, it's a new sentence
                if line[0].isdigit() and line[1] == '.':
                    if current_result:
                        results.append(current_result)
                    current_result = {}
                    continue
                    
                if ':' in line:
                    key, value = line.split(':', 1)
                    current_result[key.lower()] = value.strip()
            
            # Append the last result
            if current_result:
                results.append(current_result)
            
            return results
            
        except Exception as e:
            return [{"original": sent, "corrected": "Error during check", 
                    "explanation": f"Error: {str(e)}"} for sent in sentences]

class DocumentStats:
    def __init__(self, text: str, sentences: list, corrections: list):
        self.text = text
        self.sentences = sentences
        self.corrections = corrections
        
    def get_basic_stats(self) -> dict:
        """Calculate basic document statistics."""
        words = re.findall(r'\w+', self.text.lower())
        return {
            'Total Words': len(words),
            'Total Sentences': len(self.sentences),
            'Average Words per Sentence': round(len(words) / len(self.sentences), 1),
            'Sentences Needing Correction': sum(1 for c in self.corrections if c['corrected'] != "No corrections needed"),
            'Error-free Sentences': sum(1 for c in self.corrections if c['corrected'] == "No corrections needed")
        }
    
    def get_word_length_distribution(self) -> pd.DataFrame:
        """Calculate distribution of word lengths."""
        words = re.findall(r'\w+', self.text.lower())
        lengths = [len(word) for word in words]
        counts = Counter(lengths)
        return pd.DataFrame({
            'Word Length': list(counts.keys()),
            'Count': list(counts.values())
        }).sort_values('Word Length')
    
    def get_correction_types(self) -> pd.DataFrame:
        """Analyze types of corrections made."""
        correction_types = Counter()
        for corr in self.corrections:
            if corr['corrected'] != "No corrections needed":
                if "spelling" in corr['explanation'].lower():
                    correction_types['Spelling'] += 1
                if "grammar" in corr['explanation'].lower():
                    correction_types['Grammar'] += 1
                if "punctuation" in corr['explanation'].lower():
                    correction_types['Punctuation'] += 1
                
        return pd.DataFrame({
            'Type': list(correction_types.keys()),
            'Count': list(correction_types.values())
        })

def display_document_stats(stats: DocumentStats):
    """Display document statistics and visualizations."""
    
    # Use a container for better spacing
    with st.container():
        # Basic stats in columns with more spacing
        basic_stats = stats.get_basic_stats()
        cols = st.columns(len(basic_stats))
        for col, (stat, value) in zip(cols, basic_stats.items()):
            with col:
                st.metric(stat, value)
                st.write("")  # Add some vertical spacing
    
    # Create three columns for better space utilization
    col1, col2, col3 = st.columns([4, 4, 3])
    
    with col1:
        st.subheader("Word Length Distribution")
        word_lengths = stats.get_word_length_distribution()
        chart = alt.Chart(word_lengths).mark_bar().encode(
            x=alt.X('Word Length:Q', bin=False),
            y='Count:Q',
            tooltip=['Word Length', 'Count']
        ).properties(height=400)  # Increased height
        st.altair_chart(chart, use_container_width=True)
    
    with col2:
        st.subheader("Correction Analysis")
        correction_data = stats.get_correction_types()
        if not correction_data.empty:
            pie = alt.Chart(correction_data).mark_arc().encode(
                theta='Count:Q',
                color=alt.Color('Type:N', scale=alt.Scale(scheme='category10')),
                tooltip=['Type', 'Count']
            ).properties(height=400)  # Increased height
            st.altair_chart(pie, use_container_width=True)
        else:
            st.info("No corrections needed in the document!")
    
    with col3:
        # Add accuracy score in its own column
        st.subheader("Document Quality Score")
        accuracy = (basic_stats['Error-free Sentences'] / basic_stats['Total Sentences']) * 100
        st.progress(accuracy / 100)
        st.markdown(f"<h1 style='text-align: center; color: {'green' if accuracy > 80 else 'orange' if accuracy > 60 else 'red'}; font-size: 48px;'>{accuracy:.1f}%</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Document Accuracy</p>", unsafe_allow_html=True)

def get_initial_document_stats(doc: Document) -> dict:
    """Calculate initial document statistics before grammar checking."""
    # Extract text from paragraphs
    full_text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
    words = re.findall(r'\w+', full_text.lower())
    
    # Simple sentence splitting for initial stats
    simple_sentences = [s.strip() for s in re.split(r'[.!?]+', full_text) if s.strip()]
    
    return {
        'text': full_text,
        'stats': {
            'Total Words': len(words),
            'Total Sentences': len(simple_sentences),
            'Average Words per Sentence': round(len(words) / max(len(simple_sentences), 1), 1),
            'Unique Words': len(set(words)),
            'Document Length': f"{len(full_text)} characters"
        },
        'word_lengths': pd.DataFrame({
            'Word Length': list(Counter(len(word) for word in words).keys()),
            'Count': list(Counter(len(word) for word in words).values())
        }).sort_values('Word Length')
    }

def display_initial_stats(stats: dict):
    """Display initial document statistics before grammar checking."""
    st.header("📊 Initial Document Analysis")
    
    # Use a container for better spacing
    with st.container():
        # Display basic stats in columns with more spacing
        cols = st.columns(len(stats['stats']))
        for col, (stat, value) in zip(cols, stats['stats'].items()):
            with col:
                st.metric(stat, value)
                st.write("")  # Add some vertical spacing
    
    # Create three columns for better space utilization
    col1, col2, col3 = st.columns([4, 4, 3])
    
    with col1:
        st.subheader("Word Length Distribution")
        chart = alt.Chart(stats['word_lengths']).mark_bar().encode(
            x=alt.X('Word Length:Q', bin=False),
            y='Count:Q',
            tooltip=['Word Length', 'Count']
        ).properties(height=400)  # Increased height
        st.altair_chart(chart, use_container_width=True)
    
    with col2:
        st.subheader("Most Common Words")
        words = re.findall(r'\w+', stats['text'].lower())
        word_freq = pd.DataFrame({
            'Word': list(Counter(words).keys())[:10],
            'Frequency': list(Counter(words).values())[:10]
        })
        
        chart = alt.Chart(word_freq).mark_bar().encode(
            y=alt.Y('Word:N', sort='-x'),
            x='Frequency:Q',
            tooltip=['Word', 'Frequency'],
            color=alt.Color('Frequency:Q', scale=alt.Scale(scheme='blues'))
        ).properties(height=400)  # Increased height
        st.altair_chart(chart, use_container_width=True)
    
    with col3:
        st.subheader("Document Overview")
        # Add a summary card with styled metrics
        st.markdown("""
        <style>
        .metric-card {
            padding: 20px;
            border-radius: 10px;
            background-color: #f0f2f6;
            margin: 10px 0;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="metric-card">
            <h3>Key Metrics</h3>
            <p>📝 Document Length: {stats['stats']['Document Length']}</p>
            <p>📚 Unique Words: {stats['stats']['Unique Words']}</p>
            <p>📊 Vocabulary Density: {(stats['stats']['Unique Words'] / stats['stats']['Total Words'] * 100):.1f}%</p>
        </div>
        """, unsafe_allow_html=True)

def main():
    # Title with custom styling
    st.markdown("""
        <h1 style='text-align: center; margin-bottom: 2rem;'>
            📝 Document Grammar and Spelling Checker
        </h1>
    """, unsafe_allow_html=True)
    
    # Add a sidebar with information
    with st.sidebar:
        st.header("About")
        st.info("""
        This tool analyzes your document for:
        - Grammar errors
        - Spelling mistakes
        - Punctuation issues
        
        It provides detailed corrections and 
        comprehensive document statistics.
        """)
    
    # Initialize session state
    if 'processed_doc' not in st.session_state:
        st.session_state.processed_doc = None
    if 'initial_stats' not in st.session_state:
        st.session_state.initial_stats = None
    
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
    
    # Show initial statistics as soon as file is uploaded
    if uploaded_file:
        try:
            # Only calculate initial stats if not already done for this file
            if not st.session_state.initial_stats:
                doc = Document(uploaded_file)
                st.session_state.initial_stats = get_initial_document_stats(doc)
            
            # Display initial statistics
            display_initial_stats(st.session_state.initial_stats)
            
            # Add separator before grammar check section
            st.markdown("---")
            st.header("Grammar Check")
            st.info("Click 'Check Document' to perform grammar and spelling analysis")
        
        except Exception as e:
            st.error(f"Error reading document: {str(e)}")
    
    # Grammar check button and processing
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
            
            # Step 3: Check sentences in batches (75% of total progress)
            corrections = []
            sentence_progress_weight = 0.75
            total_batches = (len(sentences) + 9) // 10  # Round up division
            progress_per_batch = sentence_progress_weight / total_batches
            current_progress = 0.15  # Starting from previous progress
            
            # Process sentences in batches of 10
            for i in range(0, len(sentences), 10):
                batch = sentences[i:i+10]
                current_batch_num = (i // 10) + 1
                status_text.text(f"Checking batch {current_batch_num} of {total_batches} (sentences {i+1}-{min(i+10, len(sentences))})...")
                
                # Check the batch of sentences
                batch_results = checker.check_sentences_batch(batch)
                corrections.extend(batch_results)
                
                # Update progress
                current_progress += progress_per_batch
                progress_bar.progress(current_progress)
                time.sleep(0.5)  # Reduced sleep time for better UX
            
            # Display document statistics
            st.header("📊 Detailed Grammar Analysis")
            doc_stats = DocumentStats(st.session_state.initial_stats['text'], sentences, corrections)
            display_document_stats(doc_stats)
            
            # Add separator
            st.markdown("---")
            
            # Generate correction document with updated header
            st.header("📑 Correction Details")
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
