import os
import json
import streamlit as st

if "GOOGLE_CREDENTIALS" in st.secrets:
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    with open("credentials.json", "w") as f:
        json.dump(creds_dict, f)
elif os.path.exists("credentials.json"):
    with open("credentials.json", "r") as f:
        creds_dict = json.load(f)
else:
    st.error("No Google credentials found! Please add credentials.json to the project root or set GOOGLE_CREDENTIALS in Streamlit secrets.")
    st.stop()

import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import time
import random
from plyer import notification

# If modifying these SCOPES, delete the file token.json
SCOPES = ['https://www.googleapis.com/auth/documents']


def extract_doc_id(link):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", link)
    return match.group(1) if match else None

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            if os.environ.get("STREAMLIT_SERVER_HEADLESS", "") == "1":
                # On Streamlit Cloud, use run_console (no browser available)
                creds = flow.run_console()
            else:
                # Local, use run_local_server
                creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def read_doc_content(service, doc_id):
    doc = service.documents().get(documentId=doc_id).execute()
    content = doc.get('body', {}).get('content', [])
    text = ''
    for value in content:
        if 'paragraph' in value:
            elements = value['paragraph'].get('elements', [])
            for elem in elements:
                if 'textRun' in elem:
                    text += elem['textRun'].get('content', '')
    return text

def get_end_index(service, doc_id):
    doc = service.documents().get(documentId=doc_id).execute()
    content = doc.get('body', {}).get('content', [])
    if content:
        return content[-1]['endIndex'] - 1
    return 1

def append_text_to_doc(service, doc_id, text):
    end_index = get_end_index(service, doc_id)
    requests = [
        {
            'insertText': {
                'location': {
                    'index': end_index
                },
                'text': text
            }
        }
    ]
    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

# QWERTY keyboard adjacency map for typo simulation
QWERTY_ADJACENCY = {
    'a': 'qwsz', 'b': 'vghn', 'c': 'xdfv', 'd': 'ersfcx', 'e': 'wsdr', 'f': 'rtgdvc',
    'g': 'tyfhvb', 'h': 'yugjnb', 'i': 'ujko', 'j': 'uikhmn', 'k': 'ijolm', 'l': 'kop',
    'm': 'njk', 'n': 'bhjm', 'o': 'iklp', 'p': 'ol', 'q': 'wa', 'r': 'edft',
    's': 'wedxza', 't': 'rfgy', 'u': 'yhji', 'v': 'cfgb', 'w': 'qase', 'x': 'zsdc',
    'y': 'tghu', 'z': 'asx',
    'A': 'QWSZ', 'B': 'VGHN', 'C': 'XDFV', 'D': 'ERSFCX', 'E': 'WSDR', 'F': 'RTGDVC',
    'G': 'TYFHVB', 'H': 'YUGJNB', 'I': 'UJKO', 'J': 'UIKHMN', 'K': 'IJOLM', 'L': 'KOP',
    'M': 'NJK', 'N': 'BHJM', 'O': 'IKLP', 'P': 'OL', 'Q': 'WA', 'R': 'EDFT',
    'S': 'WEDXZA', 'T': 'RFGY', 'U': 'YHJI', 'V': 'CFGB', 'W': 'QASE', 'X': 'ZSDC',
    'Y': 'TGHU', 'Z': 'ASX'
}

def make_typo(word):
    if len(word) < 2:
        return word
    i = random.randint(0, len(word) - 1)
    original_char = word[i]
    if original_char in QWERTY_ADJACENCY and QWERTY_ADJACENCY[original_char]:
        typo_char = random.choice(QWERTY_ADJACENCY[original_char])
    else:
        typo_char = original_char
    return word[:i] + typo_char + word[i+1:]

def delete_last_word(service, doc_id, word):
    end_index = get_end_index(service, doc_id)
    start_index = end_index - len(word) - 1  # +1 for the space
    requests = [
        {
            'deleteContentRange': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                }
            }
        }
    ]
    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

# Streamlit session state for pause/stop and typing progress
if 'typing' not in st.session_state:
    st.session_state['typing'] = False
if 'paused' not in st.session_state:
    st.session_state['paused'] = False
if 'stopped' not in st.session_state:
    st.session_state['stopped'] = False
if 'words' not in st.session_state:
    st.session_state['words'] = []
if 'total_words' not in st.session_state:
    st.session_state['total_words'] = 0
if 'current_index' not in st.session_state:
    st.session_state['current_index'] = 0

st.title("Google Docs Copier Bot")

# Remove this line, as we are not using Streamlit secrets locally
# st.write("Secrets keys available:", list(st.secrets.keys()))

st.write("Paste the links to your Google Docs below:")

source_link = st.text_input("Source Google Doc Link (Read from)")
target_link = st.text_input("Target Google Doc Link (Type to)")

# Show estimated completion time slider before typing starts
source_id = extract_doc_id(source_link)
batch_size = 3  # keep in sync with main logic
words_preview = []
total_words_preview = 0
if source_id:
    try:
        creds = authenticate()
        service = build('docs', 'v1', credentials=creds)
        source_text = read_doc_content(service, source_id)
        words_preview = source_text.split()
        total_words_preview = len(words_preview)
    except Exception as e:
        st.error(f"Error loading source doc: {e}")

if total_words_preview > 0:
    typo_rate = 0.05  # 1 in 20 words
    avg_typo_penalty = 1.2  # seconds
    break_rate = 1/30  # every 30 words
    avg_break_duration = 2.5  # seconds
    api_penalty = 0.2  # seconds per batch
    def estimate_time(wpm):
        base_typing_time = total_words_preview / wpm * 60
        typo_time = total_words_preview * typo_rate * avg_typo_penalty
        break_time = (total_words_preview * break_rate) * avg_break_duration
        api_time = (total_words_preview / batch_size) * api_penalty
        return base_typing_time + typo_time + break_time + api_time
    def estimate_wpm(seconds):
        # Invert the estimate_time function to solve for WPM
        # Use binary search between 50 and 175
        low, high = 50, 175
        for _ in range(20):
            mid = (low + high) / 2
            t = estimate_time(mid)
            if abs(t - seconds) < 1:
                return int(mid)
            if t > seconds:
                low = mid
            else:
                high = mid
        return int(mid)
    min_time = 234  # 3 min 54 sec in seconds
    max_time = int(estimate_time(50))
    if 'est_time_setting' not in st.session_state:
        st.session_state['est_time_setting'] = int(estimate_time(80))
    # Show a larger estimated time slider adjustable by seconds
    def format_time(seconds):
        mins, secs = divmod(seconds, 60)
        return f"{mins} min {secs} sec"
    est_time_val = st.session_state['est_time_setting']
    est_time_slider = st.slider('Estimated Completion Time', min_value=min_time, max_value=max_time, value=est_time_val, step=1, key='est_time_slider', format=None, label_visibility='visible')
    # Update session state and WPM based on slider
    st.session_state['est_time_setting'] = est_time_slider
    wpm_setting = estimate_wpm(st.session_state['est_time_setting'])
    mins, secs = divmod(st.session_state['est_time_setting'], 60)
    st.info(f"Estimated completion time: {mins} min {secs} sec")
    st.info(f"Calculated WPM: {wpm_setting}")
else:
    st.info("Estimated completion time will appear here once a valid source doc link is entered.")

# Use the calculated WPM for the rest of the app
wpm_setting = st.session_state.get('wpm_setting', 80)

start = st.button(f"Start Typing")
pause = st.button("Pause Typing")
continue_typing = st.button("Continue Typing")
stop = st.button("Stop Typing")

if start:
    st.session_state['typing'] = True
    st.session_state['paused'] = False
    st.session_state['stopped'] = False
    st.session_state['current_index'] = 0
    # Prepare words list
    source_id = extract_doc_id(source_link)
    target_id = extract_doc_id(target_link)
    if not source_id or not target_id:
        st.error("Invalid Google Docs link(s). Please check your input.")
        st.session_state['typing'] = False
    else:
        try:
            creds = authenticate()
            service = build('docs', 'v1', credentials=creds)
            source_text = read_doc_content(service, source_id)
            st.session_state['words'] = source_text.split()
            st.session_state['total_words'] = len(st.session_state['words'])
            st.session_state['wpm_setting'] = wpm_setting
        except Exception as e:
            st.error(f"Error: {e}")
            st.session_state['typing'] = False

if pause:
    st.session_state['paused'] = True
if continue_typing:
    st.session_state['paused'] = False
if stop:
    st.session_state['stopped'] = True
    st.session_state['typing'] = False
    st.session_state['current_index'] = 0
    st.session_state['words'] = []
    st.session_state['total_words'] = 0

if st.session_state['typing'] and not st.session_state['stopped'] and st.session_state['words']:
    target_id = extract_doc_id(target_link)
    try:
        creds = authenticate()
        service = build('docs', 'v1', credentials=creds)
        st.info(f"Typing into target document at human speed...")
        progress = st.progress(0)
        i = st.session_state['current_index']
        total_words = st.session_state['total_words']
        words = st.session_state['words']
        batch_size = 3
        # Split source text into paragraphs and words, preserving paragraph breaks
        paragraphs = source_text.split('\n') if 'source_text' in locals() else []
        para_word_lists = [p.split() for p in paragraphs]
        para_count = len(para_word_lists)
        para_index = 0
        word_index = 0
        while para_index < para_count:
            words = para_word_lists[para_index]
            total_words = len(words)
            i = word_index if para_index == st.session_state.get('para_index', 0) else 0
            while i < total_words:
                if st.session_state['stopped']:
                    st.warning("Typing stopped.")
                    break
                batch_words = []
                batch_count = 0
                while batch_count < batch_size and i < total_words:
                    while st.session_state['paused']:
                        st.info("Paused. Click 'Continue Typing' to resume.")
                        st.session_state['current_index'] = i
                        st.session_state['para_index'] = para_index
                        time.sleep(0.5)
                    if not st.session_state['paused']:
                        word = words[i]
                        # Use user-selected WPM with a small random variation for realism
                        base_wpm = st.session_state.get('wpm_setting', 80)
                        lower = max(50, base_wpm - 5)
                        upper = min(220, base_wpm + 5)
                        if lower > upper:
                            wpm = lower
                        else:
                            wpm = random.randint(lower, upper)
                        delay = 60 / (wpm * 1.0)
                        if random.randint(1, 20) == 1 and len(word) > 3:
                            typo = make_typo(word)
                            batch_words.append(typo)
                            st.info(f"(Mistyped '{word}' as '{typo}', correcting...)")
                            time.sleep(random.uniform(0.3, 0.7))
                            batch_words[-1] = word
                        else:
                            batch_words.append(word)
                        i += 1
                        st.session_state['current_index'] = i
                        progress.progress(i / total_words)
                        if i % random.randint(20, 40) == 0 and i != 0:
                            st.info("Taking a short break...")
                            time.sleep(random.uniform(1.5, 3.5))
                        else:
                            time.sleep(delay)
                        batch_count += 1
                if batch_words:
                    append_text_to_doc(service, target_id, ' '.join(batch_words) + ' ')
            # After each paragraph, insert a paragraph break (newline)
            if para_index < para_count - 1:
                append_text_to_doc(service, target_id, '\n')
            para_index += 1
            word_index = 0
        if not st.session_state['stopped']:
            st.success("Done typing into target document!")
            try:
                if hasattr(notification, "notify") and callable(notification.notify):
                    notification.notify(
                        title='Google Docs Copier Bot',
                        message='Typing has been completed!',
                        timeout=5
                    )
            except Exception as notify_err:
                st.info(f"Notification could not be shown: {notify_err}")
        st.session_state['typing'] = False
        st.session_state['current_index'] = 0
        st.session_state['words'] = []
        st.session_state['total_words'] = 0
        st.session_state['para_index'] = 0
    except Exception as e:
        st.error(f"Error: {e}")
        st.session_state['typing'] = False
        st.session_state['current_index'] = 0
        st.session_state['words'] = []
        st.session_state['total_words'] = 0 