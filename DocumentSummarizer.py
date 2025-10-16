import streamlit as st
import PyPDF2
import docx
import pandas as pd
from openai import OpenAI
import io
import traceback

# Page configuration
st.set_page_config(page_title="GCC Microsite - Document Analyzer & Chat", layout="wide", page_icon="📄")

# Initialize session state
if 'messages' not in st.session_state:
    st.session_state.messages = []
if 'document_content' not in st.session_state:
    st.session_state.document_content = ""
if 'summary' not in st.session_state:
    st.session_state.summary = None
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False

def extract_text_from_pdf(file):
    """Extract text from PDF file"""
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None

def extract_text_from_docx(file):
    """Extract text from DOCX file"""
    try:
        doc = docx.Document(file)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        st.error(f"Error reading DOCX: {str(e)}")
        return None

def extract_text_from_txt(file):
    """Extract text from TXT file"""
    try:
        return file.read().decode('utf-8')
    except Exception as e:
        st.error(f"Error reading TXT: {str(e)}")
        return None

def extract_text_from_excel(file):
    """Extract text from Excel file"""
    try:
        df = pd.read_excel(file, sheet_name=None)
        text = ""
        for sheet_name, sheet_data in df.items():
            text += f"\n\n=== Sheet: {sheet_name} ===\n"
            text += sheet_data.to_string(index=False)
        return text
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return None

def process_document(file):
    """Process uploaded document based on file type"""
    if file is None:
        return None
    
    file_extension = file.name.split('.')[-1].lower()
    
    try:
        if file_extension == 'pdf':
            return extract_text_from_pdf(file)
        elif file_extension in ['docx', 'doc']:
            return extract_text_from_docx(file)
        elif file_extension == 'txt':
            return extract_text_from_txt(file)
        elif file_extension in ['xls', 'xlsx']:
            return extract_text_from_excel(file)
        else:
            st.error("Unsupported file format")
            return None
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        traceback.print_exc()
        return None

def get_document_summary(content, api_key):
    """Get summary from OpenAI API"""
    try:
        client = OpenAI(api_key=api_key)
        
        # Truncate content if too long
        max_chars = 12000
        truncated_content = content[:max_chars]
        if len(content) > max_chars:
            truncated_content += "\n\n[Document truncated for analysis...]"
        
        prompt = f"""Analyze the following document and provide a detailed summary with:

1. **Key Highlights**: List the main points, findings, and important information
2. **Important Metrics**: Extract all numbers, percentages, statistics, financial figures, and quantitative data
3. **Dates**: List all dates mentioned with their context

Format your response clearly with these three sections.

Document:
{truncated_content}
"""
        
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a document analysis expert who extracts key information, metrics, and dates from documents."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=2000
        )
        
        return response.choices[0].message.content
    except Exception as e:
        error_msg = f"Error calling OpenAI API: {str(e)}"
        st.error(error_msg)
        traceback.print_exc()
        return None

def chat_with_document(question, document_content, chat_history, api_key):
    """Chat about the document using OpenAI API"""
    try:
        client = OpenAI(api_key=api_key)
        
        # Truncate document if too long
        max_chars = 10000
        truncated_doc = document_content[:max_chars]
        
        messages = [
            {"role": "system", "content": f"You are a helpful assistant that answers questions about a document. Base your answers only on the document content provided.\n\nDocument:\n{truncated_doc}"}
        ]
        
        # Add recent chat history for context
        for msg in chat_history[-8:]:
            messages.append({"role": msg["role"], "content": msg["content"]})
        
        messages.append({"role": "user", "content": question})
        
        response = client.chat.completions.create(
            model="gpt-4",
            messages=messages,
            temperature=0.7,
            max_tokens=1000
        )
        
        return response.choices[0].message.content
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        traceback.print_exc()
        return error_msg

# Main UI
st.title("📄 GCC Microsite - Document Analyzer & Chat Assistant")
st.markdown("Upload your documents and get intelligent summaries with interactive Q&A")

# Sidebar for API key and file upload
with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input(
        "OpenAI API Key", 
        type="password", 
        help="Enter your OpenAI API key (starts with sk-)",
        placeholder="sk-..."
    )
    
    st.markdown("---")
    st.header("📤 Upload Document")
    uploaded_file = st.file_uploader(
        "Choose a file",
        type=['pdf', 'doc', 'docx', 'txt', 'xls', 'xlsx'],
        help="Supported formats: PDF, DOC, DOCX, TXT, XLS, XLSX"
    )
    
    if uploaded_file:
        st.success(f"✅ File loaded: {uploaded_file.name}")
        st.info(f"File size: {uploaded_file.size / 1024:.2f} KB")
    
    analyze_button = st.button("🔍 Analyze Document", type="primary", use_container_width=True)
    
    if analyze_button:
        if not uploaded_file:
            st.error("⚠️ Please upload a file first")
        elif not api_key:
            st.error("⚠️ Please enter your OpenAI API key")
        elif not api_key.startswith('sk-'):
            st.error("⚠️ Invalid API key format. It should start with 'sk-'")
        else:
            with st.spinner("🔄 Processing document..."):
                content = process_document(uploaded_file)
                if content and len(content.strip()) > 0:
                    st.session_state.document_content = content
                    
                    with st.spinner("🤖 Generating AI summary..."):
                        summary = get_document_summary(content, api_key)
                        if summary:
                            st.session_state.summary = summary
                            st.session_state.messages = []
                            st.session_state.file_processed = True
                            st.success("✅ Document analyzed successfully!")
                            st.rerun()
                        else:
                            st.error("Failed to generate summary")
                else:
                    st.error("❌ Could not extract text from document")
    
    st.markdown("---")
    st.markdown("### 💡 Tips")
    st.markdown("""
    - Ensure your API key is valid
    - Larger files may take longer
    - Chat remembers context
    """)

# Main content area
col1, col2 = st.columns([1, 1])

with col1:
    st.header("📊 Document Summary")
    if st.session_state.summary:
        st.markdown(st.session_state.summary)
        
        with st.expander("📄 View Full Document Text"):
            st.text_area(
                "Document Content", 
                st.session_state.document_content, 
                height=300, 
                disabled=True,
                label_visibility="collapsed"
            )
    else:
        st.info("👈 Upload a document and click 'Analyze Document' to see the summary here")

with col2:
    st.header("💬 Chat with Document")
    
    if st.session_state.file_processed:
        # Chat messages container
        chat_container = st.container(height=400)
        with chat_container:
            if not st.session_state.messages:
                st.info("💭 Ask me anything about the document!")
            
            for message in st.session_state.messages:
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])
        
        # Chat input
        if prompt := st.chat_input("Ask a question about the document...", key="chat_input"):
            if not api_key:
                st.error("⚠️ Please enter your OpenAI API key in the sidebar")
            else:
                # Add user message
                st.session_state.messages.append({"role": "user", "content": prompt})
                
                # Get AI response
                with st.spinner("🤔 Thinking..."):
                    response = chat_with_document(
                        prompt,
                        st.session_state.document_content,
                        st.session_state.messages,
                        api_key
                    )
                
                # Add assistant message
                st.session_state.messages.append({"role": "assistant", "content": response})
                st.rerun()
        
        # Clear chat button
        if st.session_state.messages:
            if st.button("🗑️ Clear Chat History", use_container_width=True):
                st.session_state.messages = []
                st.rerun()
    else:
        st.info("👈 Please analyze a document first to start chatting")
        st.markdown("""
        ### 🎯 What you can ask:
        - "What are the main conclusions?"
        - "Summarize the key findings"
        - "What dates are mentioned?"
        - "What are the important numbers?"
        - "Explain section X in detail"
        """)

# Footer info
if not st.session_state.file_processed:
    st.markdown("---")
    st.markdown("""
    ### 📖 How to use:
    1. **Enter your OpenAI API key** in the sidebar (get one at platform.openai.com)
    2. **Upload a document** (PDF, DOCX, TXT, XLS, or XLSX)
    3. **Click "Analyze Document"** to get an automated summary
    4. **Use the chat** to ask specific questions about the document
    
    ### ✨ Features:
    - 🎯 **Key Highlights**: Main points and findings
    - 📊 **Important Metrics**: Numbers, statistics, and data
    - 📅 **Dates**: All mentioned dates with context
    - 💬 **Interactive Chat**: Ask follow-up questions
    """)

st.markdown("---")
st.caption("Built with Streamlit and OpenAI GPT-4")