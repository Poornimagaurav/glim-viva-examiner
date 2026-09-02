"""
Simple Project Viva Examiner — text-only, Groq-powered.

A stripped-down oral-viva bot: student uploads their project report (PDF or
DOCX), the AI asks a fixed number of questions about it one at a time via
plain text chat, then gives a final score and feedback. No voice, no
anti-cheat/proctoring, no manual score override — just the core Q&A loop.

SETUP
  1. pip install -r requirements_simple.txt  (see bottom of this file for
     what to put in it)
  2. Get a free Groq API key at https://console.groq.com/keys
  3. Set it as an environment variable GROQ_API_KEY, or (if deploying on
     Streamlit Community Cloud) add it under Settings -> Secrets as:
         GROQ_API_KEY = "gsk_..."
  4. Run:  streamlit run simple_viva.py
"""

import os
import re

import streamlit as st
from groq import Groq
import fitz  # PyMuPDF, for reading PDFs
from docx import Document

# ---- Configuration --------------------------------------------------------
TOTAL_QUESTIONS = 6          # how many Q&A exchanges before the viva ends
MODEL = "openai/gpt-oss-20b" # Groq model to use
MAX_PROJECT_CHARS = 6000     # how much of the uploaded report to feed the AI

st.set_page_config(page_title="Simple Project Viva", page_icon="🎓")
st.title("🎓 Simple Project Viva Examiner")
st.caption("Powered by Groq — text-only, question-and-answer viva based on your uploaded project")

client = Groq(api_key=st.secrets.get("GROQ_API_KEY") or os.environ.get("GROQ_API_KEY"))


# ---- Reading the uploaded project file ------------------------------------
def extract_text(uploaded_file):
    """Pulls plain text out of an uploaded PDF or DOCX file."""
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        return "\n".join(page.get_text() for page in pdf)
    elif name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)
    return ""


def truncate(text, max_chars=MAX_PROJECT_CHARS):
    return text[:max_chars]


# ---- Talking to Groq --------------------------------------------------------
def chat_with_llm(messages, max_retries=3):
    """Calls the Groq API, with a small retry for transient errors."""
    last_error = None
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=messages,
                max_tokens=512,
            )
            return response.choices[0].message.content
        except Exception as e:
            last_error = e
    st.error(f"⚠️ Could not reach the examiner AI after {max_retries} attempts: {last_error}")
    return None


def parse_score(text):
    """Pulls a 0-10 score out of the closing message, if present."""
    if not text:
        return None
    match = re.search(r'score\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*/\s*10', text, re.IGNORECASE)
    if match:
        value = float(match.group(1))
        if 0 <= value <= 10:
            return value
    return None


# ---- Session state ----------------------------------------------------------
for key, default in [
    ("messages", []),
    ("started", False),
    ("ended", False),
    ("exchange_count", 0),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ---- Sidebar: setup ----------------------------------------------------------
with st.sidebar:
    st.header("Setup")
    student_name = st.text_input("Student Name")
    uploaded_file = st.file_uploader("Upload project report", type=["pdf", "docx"])
    start_btn = st.button(
        "Start Viva", type="primary",
        disabled=(uploaded_file is None or not student_name or st.session_state.started),
    )

# ---- Kick off the viva --------------------------------------------------------
if start_btn and not st.session_state.started:
    project_text = extract_text(uploaded_file)

    system_prompt = f"""You are an oral viva examiner questioning {student_name} about their project report.

Ask ONE question at a time, based ONLY on the project text below. Keep each
question concise — this is a spoken-style oral viva. Wait for the student's
answer before asking the next question. Never ask more than one question in
a single reply.

After exactly {TOTAL_QUESTIONS} questions have been asked and answered, stop
asking questions and instead reply with EXACTLY this closing format, and
nothing else:

VIVA COMPLETE
Score: X/10
Feedback: [2-3 sentences on the student's overall performance]

Replace X with an actual number from 0 to 10 based on how well the student
answered — never leave X as a literal letter.

--- PROJECT TEXT ---
{truncate(project_text)}
--- END PROJECT TEXT ---

Begin now by greeting {student_name} and asking your first question."""

    st.session_state.messages = [{"role": "system", "content": system_prompt}]
    with st.spinner("Examiner is preparing..."):
        reply = chat_with_llm(st.session_state.messages)
    if reply:
        st.session_state.messages.append({"role": "assistant", "content": reply})
        st.session_state.started = True

# ---- Display the conversation so far ------------------------------------------
for msg in st.session_state.messages:
    if msg["role"] == "system":
        continue
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ---- Handle the student's typed answers ---------------------------------------
if st.session_state.started and not st.session_state.ended:
    user_input = st.chat_input("Type your answer here...")
    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.session_state.exchange_count += 1
        with st.chat_message("user"):
            st.markdown(user_input)

        with st.chat_message("assistant"):
            with st.spinner("Examiner is responding..."):
                reply = chat_with_llm(st.session_state.messages)
            if reply:
                st.markdown(reply)

        if reply:
            st.session_state.messages.append({"role": "assistant", "content": reply})
            if "VIVA COMPLETE" in reply or st.session_state.exchange_count >= TOTAL_QUESTIONS:
                st.session_state.ended = True

# ---- Final result -------------------------------------------------------------
if st.session_state.ended:
    st.divider()
    st.subheader("🏁 Viva Complete")
    final_text = st.session_state.messages[-1]["content"]
    score = parse_score(final_text)
    if score is not None:
        st.metric("Final Score", f"{score}/10")
    else:
        st.info("No numeric score was found in the examiner's closing message above.")
