from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta
import os, time, openai, json
import textwrap
import re
from bs4 import BeautifulSoup
from flask import Flask, jsonify, send_from_directory, request, session, render_template_string, redirect, url_for
from flask_cors import CORS
import threading
import hashlib
import secrets
from functools import wraps
from datetime import datetime, timedelta
import requests
import logging


load_dotenv()

### CONSTANTS

# Start from now - only process new emails going forward
START_TIME = datetime.now(timezone.utc) - timedelta(days=1)  # Check last 1 day for initial sync
DELTA_TOKENS_FILE = "delta_tokens.json"

def load_delta_tokens():
    """Load delta tokens from JSON file"""
    try:
        with open(DELTA_TOKENS_FILE, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print("📚 No existing delta tokens found, will start fresh sync")
        return {}
    except Exception as e:
        print(f"⚠️ Could not load delta tokens: {e}, starting fresh")
        return {}

def save_delta_tokens(tokens):
    """Save delta tokens to JSON file atomically"""
    try:
        tmp_file = DELTA_TOKENS_FILE + ".tmp"
        with open(tmp_file, 'w') as f:
            json.dump(tokens, f, indent=2)
        os.replace(tmp_file, DELTA_TOKENS_FILE)  # Atomic write
        print(f"💾 Saved delta tokens: {list(tokens.keys())}")
    except Exception as e:
        print(f"⚠️ Could not save delta tokens: {e}")

def get_bearer_token():
    """Get bearer token from O365 account for direct Graph API calls"""
    try:
        # For this O365 library version, try getting token from the session
        session = account.connection.get_session()
        
        # Debug what's available in the session
        print(f"🔍 Session type: {type(session)}")
        print(f"🔍 Session attributes: {[attr for attr in dir(session) if not attr.startswith('_')]}")
        
        # Debug token-related attributes
        print(f"🔍 access_token value: {getattr(session, 'access_token', 'None')}")
        print(f"🔍 token value: {getattr(session, 'token', 'None')}")
        print(f"🔍 authorized: {getattr(session, 'authorized', 'None')}")
        
        # Method 1: Direct access_token attribute (OAuth2Session)
        if hasattr(session, 'access_token') and session.access_token:
            print(f"🔑 Got bearer token from access_token: {session.access_token[:20]}...")
            return session.access_token
            
        # Method 2: From token attribute
        if hasattr(session, 'token') and session.token:
            print(f"🔍 Token type: {type(session.token)}")
            if isinstance(session.token, dict):
                access_token = session.token.get('access_token')
                if access_token:
                    print(f"🔑 Got bearer token from token dict: {access_token[:20]}...")
                    return access_token
            elif isinstance(session.token, str):
                print(f"🔑 Got bearer token from token string: {session.token[:20]}...")
                return session.token
        
        # Method 3: Try to get the Authorization header from the session
        if hasattr(session, 'headers'):
            auth_header = session.headers.get('Authorization', '')
            if auth_header.startswith('Bearer '):
                token = auth_header[7:]  # Remove 'Bearer ' prefix
                print(f"🔑 Got bearer token from headers: {token[:20]}...")
                return token
        
        # Try making a dummy request to see if we can intercept the auth
        try:
            # Use the existing mailbox connection to make an authenticated request
            # This should trigger token refresh if needed
            resp = session.get('https://graph.microsoft.com/v1.0/me', timeout=10)
            print(f"🔍 Test API call status: {resp.status_code}")
            
            # Check if the request was made with auth headers
            if hasattr(resp, 'request') and hasattr(resp.request, 'headers'):
                auth_header = resp.request.headers.get('Authorization', '')
                if auth_header.startswith('Bearer '):
                    token = auth_header[7:]
                    print(f"🔑 Got bearer token from request: {token[:20]}...")
                    return token
                    
        except Exception as api_e:
            print(f"⚠️ Test API call failed: {api_e}")
                
        print(f"❌ Could not extract bearer token from any method")
        return None
        
    except Exception as e:
        print(f"❌ Error getting bearer token: {e}")
        import traceback
        print(f"❌ Traceback: {traceback.format_exc()}")
        return None

def get_folder_ids():
    """Get the actual folder IDs for inbox and junk from Microsoft Graph"""
    try:
        # Use the existing O365 objects to get folder IDs
        inbox_id = inbox_folder.folder_id if hasattr(inbox_folder, 'folder_id') else 'inbox'
        junk_id = junk_folder.folder_id if hasattr(junk_folder, 'folder_id') else 'junkemail'
        
        print(f"📁 Folder IDs - Inbox: {inbox_id}, Junk: {junk_id}")
        return {"inbox": inbox_id, "junk": junk_id}
    except Exception as e:
        print(f"❌ Error getting folder IDs: {e}")
        # Fallback to standard folder names
        return {"inbox": "inbox", "junk": "junkemail"}

def graph_delta_sync(folder_name, folder_id, delta_url=None, since_utc_iso=None):
    """
    Perform delta sync using direct Microsoft Graph API calls
    Returns: (changes_list, new_delta_url)
    """
    bearer_token = get_bearer_token()
    if not bearer_token:
        print(f"❌ Cannot sync {folder_name}: No bearer token")
        return [], None
    
    # Build the URL
    if delta_url:
        url = delta_url
        print(f"🔄 Resuming delta sync for {folder_name} with existing token")
    else:
        # Initial seeding
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages/delta"
        if since_utc_iso:
            url += f"?$filter=receivedDateTime ge {since_utc_iso}"
        print(f"🌱 Starting fresh delta sync for {folder_name} since {since_utc_iso}")
    
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json"
    }
    
    changes = []
    final_delta_url = None
    
    while url:
        try:
            print(f"📡 Calling Graph API: {url[:100]}...")
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            # Process changes
            for item in data.get("value", []):
                if "@removed" in item:
                    changes.append({"removed": True, "id": item.get("id")})
                    print(f"🗑️ Email removed: {item.get('id')}")
                else:
                    changes.append({"removed": False, "message": item})
                    subject = item.get("subject", "No subject")[:50]
                    sender = item.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                    print(f"📧 Email found: {sender} | {subject}")
            
            # Follow pagination or get final delta link
            next_link = data.get("@odata.nextLink")
            delta_link = data.get("@odata.deltaLink")
            
            if next_link:
                url = next_link
            else:
                final_delta_url = delta_link
                url = None
                
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 410:
                print(f"⚠️ Delta token expired for {folder_name}, will need fresh sync")
                return [], None
            else:
                print(f"❌ HTTP error in delta sync for {folder_name}: {e.response.status_code} - {e}")
                return [], None
        except Exception as e:
            print(f"❌ Error in delta sync for {folder_name}: {e}")
            return [], None
    
    print(f"✅ Delta sync completed for {folder_name}: {len(changes)} changes, new token: {'Yes' if final_delta_url else 'No'}")
    return changes, final_delta_url

def process_graph_message(message_data, folder_name):
    """Process a message from Microsoft Graph API response"""
    try:
        # Extract key fields from Graph API response
        msg_id = message_data.get('id', '')
        subject = message_data.get('subject', 'No subject')
        sender_info = message_data.get('from', {}).get('emailAddress', {})
        sender_address = sender_info.get('address', 'Unknown')
        received_datetime = message_data.get('receivedDateTime', '')
        is_read = message_data.get('isRead', False)
        categories = message_data.get('categories', [])
        body_content = message_data.get('body', {}).get('content', '')
        
        print(f"\n🔍 Processing message from Graph API:")
        print(f"    ID: {msg_id}")
        print(f"    Subject: {subject}")
        print(f"    From: {sender_address}")
        print(f"    Received: {received_datetime}")
        print(f"    Read: {is_read}")
        print(f"    Categories: {categories}")
        
        # Skip if already processed (marked with PAIRActioned)
        if any((c or '').startswith('PAIRActioned') for c in categories):
            print(f"⏭️ Skipping already processed message: {subject}")
            return
        
        # Skip if already read (unless it's an internal reply)
        if is_read:
            print(f"⏭️ Skipping read message: {subject}")
            return
            
        # TODO: Convert Graph API message data to format compatible with existing handle_new_email function
        # This will require creating a message-like object or adapting handle_new_email to work with Graph data
        
        print(f"🎯 Would process: {subject} from {sender_address}")
        
    except Exception as e:
        print(f"❌ Error processing Graph message: {e}")

def process_folders_with_graph_api():
    """Process both inbox and junk folders using direct Microsoft Graph API"""
    try:
        folder_ids = get_folder_ids()
        start_time_iso = START_TIME.isoformat()
        
        for folder_name in ["inbox", "junk"]:
            print(f"\n🔄 Processing {folder_name.upper()} folder...")
            
            folder_id = folder_ids[folder_name]
            existing_delta_url = delta_tokens.get(folder_name)
            
            # Perform delta sync
            if existing_delta_url:
                changes, new_delta_url = graph_delta_sync(folder_name, folder_id, delta_url=existing_delta_url)
            else:
                changes, new_delta_url = graph_delta_sync(folder_name, folder_id, since_utc_iso=start_time_iso)
            
            # Process each change
            for change in changes:
                if change["removed"]:
                    print(f"🗑️ Message removed: {change['id']}")
                else:
                    process_graph_message(change["message"], folder_name)
            
            # Save new delta token
            if new_delta_url:
                delta_tokens[folder_name] = new_delta_url
                print(f"💾 Updated delta token for {folder_name}")
            else:
                print(f"⚠️ No new delta token received for {folder_name}")
        
        # Save all delta tokens
        save_delta_tokens(delta_tokens)
        
    except Exception as e:
        print(f"❌ Error in Graph API folder processing: {e}")
        import traceback
        print(f"❌ Traceback: {traceback.format_exc()}")

# Load delta tokens at startup
delta_tokens = load_delta_tokens()
print(f"📚 Loaded delta tokens for: {list(delta_tokens.keys())}")

# Keep processed_messages for now (will be removed in future refactor)
processed_messages = set()

def save_processed_messages():
    """Placeholder - will be removed when fully migrated to delta tokens"""
    pass

CLIENT_ID = "c0abfd02-2166-4a52-b052-16d1aa084afb"  # MLFA app registration
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
TENANT_ID = os.getenv("O365_TENANT_ID")
REPLY_ID_TAG = "Pair_Reply_Reference_ID"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EMAIL_TO_WATCH = os.getenv("EMAIL_TO_WATCH")


EMAILS_TO_FORWARD = ['Mujahid.rasul@mlfa.org', 'Syeda.sadiqa@mlfa.org', 'Arshia.ali.khan@mlfa.org', 'Maria.laura@mlfa.org', 'info@mlfa.org', 'aisha.ukiu@mlfa.org', 'shawn@strategichradvisory.com', 'Marium.Uddin@mlfa.org']
NONREAD_CATEGORIES = {"marketing"}  # Keep these unread
SKIP_CATEGORIES = {'spam', 'cold_outreach', 'newsletter', 'irrelevant_other'}

HUMAN_CHECK = True  # Enable human check for approval hub

# Storage for multiple pending emails
pending_emails = {}  # Dictionary to store multiple emails by ID
current_email_id = None  # Track which email is currently being shown

# Storage for forwarded email recipients (for CC functionality)
forwarded_recipients = {}  # Maps message_id to list of recipients

# Flask app for approval hub
app = Flask(__name__, static_folder='.')
# Generate a secure secret key if not provided - MUST be consistent across restarts
FIXED_SECRET_KEY = 'mlfa-2025-secret-key-for-sessions-do-not-share'
app.secret_key = os.getenv('SECRET_KEY', FIXED_SECRET_KEY)

# Security settings - adjust based on environment
IS_PRODUCTION = os.getenv('ENVIRONMENT', 'development').lower() == 'production'
app.config['SESSION_COOKIE_SECURE'] = False  # Set to True only with HTTPS
app.config['SESSION_COOKIE_HTTPONLY'] = True  # Prevent JS access to cookies
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'  # Changed to Lax to allow redirects
app.config['SESSION_COOKIE_NAME'] = 'mlfa_session'
app.config['SESSION_COOKIE_PATH'] = '/'
app.config['SESSION_COOKIE_DOMAIN'] = None  # Allow any domain/IP
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=2)  # Auto-logout after 2 hours
CORS(app, supports_credentials=True)

# Rate limiting for login attempts
login_attempts = {}  # Track failed login attempts
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_TIME = 300  # 5 minutes in seconds

# Password hashing - NEVER store plain passwords
ADMIN_PASSWORD_HASH = os.getenv('ADMIN_PASSWORD_HASH')  # Store the hash in .env
if not ADMIN_PASSWORD_HASH:
    # For initial setup only - generates hash from plain password
    temp_password = os.getenv('ADMIN_PASSWORD', secrets.token_urlsafe(16))
    ADMIN_PASSWORD_HASH = hashlib.sha256(temp_password.encode()).hexdigest()
    print(f"⚠️  No password hash found. Generated temporary password: {temp_password}")
    print(f"⚠️  Add this to your .env file: ADMIN_PASSWORD_HASH={ADMIN_PASSWORD_HASH}")

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        logged_in = session.get('logged_in')
        login_time = session.get('login_time', 0)
        
        print(f"🔐 login_required check - Path: {request.path}, logged_in: {logged_in}, IP: {request.remote_addr}")
        
        # Check if session has expired (2 hours)
        if logged_in and (time.time() - login_time > 7200):
            session.clear()
            logged_in = False
            print(f"⏰ Session expired for {request.remote_addr}")
        
        if not logged_in:
            print(f"❌ Not logged in, redirecting to login page")
            if request.path.startswith('/api/'):
                return jsonify({"error": "Authentication required"}), 401
            return redirect(url_for('login'))
        
        return f(*args, **kwargs)
    return decorated_function




###CONNECTING


from openai import OpenAI
client = OpenAI(api_key=OPENAI_API_KEY)

# Use client credentials for production (no user interaction needed)
if IS_PRODUCTION and CLIENT_SECRET:
    credentials = (CLIENT_ID, CLIENT_SECRET)
    account = Account(credentials, auth_flow_type="credentials", tenant_id=TENANT_ID)
    if not account.is_authenticated:
        account.authenticate()
else:
    # Development mode with OAuth flow
    credentials = (CLIENT_ID, None)
    token_backend = FileSystemTokenBackend(token_path=".", token_filename="o365_token.txt")
    account = Account(credentials, auth_flow_type="authorization", token_backend=token_backend)
    if not account.is_authenticated:
        account.authenticate(scopes=['basic', 'message_all'])

mailbox = account.mailbox(resource=EMAIL_TO_WATCH)
inbox_folder = mailbox.inbox_folder()
junk_folder = mailbox.junk_folder()




# Delta token functions moved above - using JSON persistence now


#Passes the subject and body of the email to chat gpt, which figures out how to handle the email. 
#Chat-GPT nicely returns the information in json format. 

#Args:
#subject (str): The subject of the email. 
#body (str): The body of the email. 

#Returns:
#A json script that includes the category of the email, who it needs to be forwarded to, and why. 

def classify_email(subject, body):
    prompt = f"""You are an email routing assistant for MLFA (Muslim Legal Fund of America), a nonprofit organization focused on legal advocacy for Muslims in the United States.

    Your job is to classify incoming emails based on their **content, sender intent, and relevance** to MLFA’s mission. Do not rely on keywords alone. Use the routing rules below to assign one or more categories and determine appropriate recipients if applicable.
    Additionally, **identify the sender’s name** when possible and include it as `name_sender` in the JSON. Prefer the “From” display name; if unavailable or generic, use a clear sign-off/signature in the body. If you cannot determine the name confidently, set `name_sender` to `"Sender"`.

    HUMAN-STYLE REPLY ESCALATION (IMPORTANT):
    Flag emails that should NOT get a generic auto-reply because they are personal/referral-like or contain substantial case detail.
    Set `needs_personal_reply=true` if ANY of these are present:
    - **Referral signals:** mentions of being referred by a person/org (e.g., imam, attorney, community leader, “X told me to contact you,” CC’ing a referrer).
    - **Personal narrative with specifics:** detailed timeline, names, dates, locations, docket/case numbers, court filings, detention/deportation details, attorney names, or attached evidence.
    - **Clearly individualized appeal:** tone reads as one-to-one help-seeking rather than a form blast.
    - **Brevity & Generic Content safeguard:** If the email is *short, vague, and generic* (e.g., “I need legal help” or “Please assist”), and does **not** include referral language or specific personal details, then set `needs_personal_reply=false` even if it asks for help.

    If none of the above apply, set `needs_personal_reply=false`.

    ROUTING RULES & RECIPIENTS:

    - **Legal inquiries** → If someone is explicitly **asking for legal help or representation**, categorize as `"legal"`. These users should be referred to MLFA’s "Apply for Help" form (no forwarding needed).

    - **Donor-related inquiries** → Categorize as `"donor"` only if the **sender is a donor** or is asking about a **specific donation**, such as issues with payment, receipts, or donation follow-ups. Forward to:
    Mujahid.rasul@mlfa.org, Syeda.sadiqa@mlfa.org

    - **Sponsorship requests** → If someone is **requesting sponsorship or financial support from MLFA**, categorize as `"sponsorship"`. Forward to:
    Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org

    - **Fellowship inquiries** → If someone is **applying for, asking about, or offering a fellowship** (legal, advocacy, or nonprofit-focused), categorize as `"fellowship"`. Forward to:
    aisha.ukiu@mlfa.org

    - **Organizational questions** → If the sender is asking about **MLFA’s internal operations**, such as leadership, partnerships, or collaboration, categorize as `"organizational"`. Forward to:
    Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org

    - **Volunteer inquiries** → If someone is **offering to volunteer** their time or skills to MLFA **or** is **asking about volunteering** (for themselves or on behalf of someone else), categorize as `"volunteer"`. These will receive an automated reply with the volunteer application form.

    - **Job applications** → If someone is **applying for a paid job**, sending a resume, or asking about open employment positions, categorize as `"job_application"`. Forward to:
    shawn@strategichradvisory.com

    - **Internship applications** → If someone is **applying for an internship** (paid or unpaid), sending a resume for an internship program, or inquiring about internship opportunities, categorize as `"internship"`. Forward to:
    aisha.ukiu@mlfa.org

    - **Media inquiries** → If the sender is a **reporter or journalist asking for comments, interviews, or statements**, categorize as `"media"`. Forward to:
    Marium.Uddin@mlfa.org

    - **Email marketing/sales** → If the sender is **offering a product, service, or software**, categorize as `"marketing"` only if:
    1) The offering is **relevant to MLFA’s nonprofit or legal work**, **and**
    2) The sender shows **clear contextual awareness** (e.g., refers to MLFA’s legal mission, Muslim families, or nonprofit context), **and**
    3) The product is **niche-specific**, such as legal case management, zakat compliance tools, intake systems for nonprofits, or Islamic legal software.
    Move to the "Sales emails" folder.
    **Do not treat generic, untargeted, or mass-promotional emails as marketing.**

    - **Cold outreach** → Any **unsolicited sales email** that lacks clear tailoring to MLFA’s work. Categorize as `"cold_outreach"` if:
    - The sender shows **no meaningful awareness** of MLFA’s mission
    - The offer is **broad, mass-marketed, or hype-driven**
    - The email uses commercial hooks like “Act now,” “800% increase,” “Only $99/month,” or “Click here”
    Even if the topic sounds legal or nonprofit-adjacent, if it **feels generic**, classify it as cold outreach.
    Mark as read; **do not** treat as marketing.

    - **Spam** → Obvious scams, phishing, AI-generated nonsense, or malicious intent. Move to Junk.

    - **Newsletter** → Bulk content like PR updates, blog digests, or mass announcements not addressed to MLFA directly. Place in "Newsletters" if available.

    - **Irrelevant (other)** → Anything that doesn't match the above and is unrelated to MLFA’s mission — e.g., misdirected emails, general inquiries, or off-topic messages. Mark as read and ignore.

    IMPORTANT GUIDELINES:
    1. Focus on **relevance and specificity**, not just keywords. The more the sender understands MLFA, the more likely it is to be legitimate.
    2. If an email is a **niche legal tech offer clearly crafted for MLFA or Muslim nonprofits**, treat it as `"marketing"` — even if unsolicited.
    3. If the offer is **generic or clearly sent in bulk**, it’s `"cold_outreach"` — even if it references legal themes or Muslim communities.
    4. Never mark cold outreach or mass sales emails as `"marketing"`, even if they reference MLFA’s field.
    5. If someone is **offering legal services**, classify as `"organizational"` only if relevant and serious (not promotional).
    6. Emails can and should have **multiple categories** when appropriate (e.g., a donor asking to volunteer → `"donor"` and `"volunteer"`).
    7. Use `all_recipients` only for forwarded categories: `"donor"`, `"sponsorship"`, `"fellowship"`, `"organizational"`, `"job_application"`, `"internship"`, `"media"`.
    8. For `"legal"`, `"volunteer"`, `"marketing"`, and all `"irrelevant"` types, leave `all_recipients` empty.

    PRIORITY & TIES:
    - If `"legal"` applies, **still include all other relevant categories** — `"legal"` is additive, never exclusive.
    - `"marketing"` vs `"cold_outreach"`: choose only one based on tailoring (see rules above).

    Return a JSON object with:
    - `categories`: array from ["legal","donor","sponsorship","fellowship","organizational","volunteer","job_application","internship","media","marketing","spam","cold_outreach","newsletter","irrelevant_other"]
    - `all_recipients`: list of MLFA email addresses (may be empty)
    - `needs_personal_reply`: boolean per the Escalation section
    - `reason`: dictionary mapping each category to a brief justification
    - `escalation_reason`: brief string explaining why `needs_personal_reply` is true (empty string if false)
    - `name_sender`: the sender’s name if confidently identified; otherwise exactly `"Sender"`

    Subject: {subject}

    Body:
    {body}
    """


    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        raw = response.choices[0].message.content.strip()
        if raw.startswith("```json"): raw = raw[len("```json"):].strip() #to prevent errors from occuring. 
        if raw.endswith("```"): raw = raw[:-3].strip()
        return json.loads(raw)
    except Exception as e:
        print(f"Classification error: {e}")
        return {}


def process_folder(folder, name, delta_token):
    """
    Delta items are treated as signals only.
    For each changed conversation, fetch ALL unread child messages and process
    them individually (oldest -> newest), never reprocessing the original/root.
    Internal replies are detected and handled before classification.
    """
    # Build delta query (optionally select a few cheap fields to reduce "shallow" items)
    qs = folder.new_query()
    if delta_token:
        qs = qs.delta_token(delta_token)
    else:
        # No delta token means first run - only get emails from START_TIME onwards
        qs = qs.on_attribute('receivedDateTime').greater_equal(START_TIME)
    qs = qs.select([
        'id', 'conversationId', 'isRead', 'receivedDateTime', 'from', 'sender', 'subject', 'categories'
    ])

    try:
        print(f"📡 Querying {name} folder for messages...")
        msgs = folder.get_messages(query=qs)
        print(f"📬 Found {len(list(msgs)) if hasattr(msgs, '__len__') else 'unknown'} messages in {name}")
        
        # Reset msgs iterator since we consumed it above
        msgs = folder.get_messages(query=qs)

        for msg in msgs:
            # For each conversation that changed, act ONLY on unread children
            conv_id = getattr(msg, 'conversation_id', None)
            if not conv_id:
                # Fallback: skip if no conversation id (rare)
                dedup_key = getattr(msg, 'internet_message_id', None) or msg.object_id
                if dedup_key in processed_messages:
                    # Silently skip already processed messages
                    continue
                try:
                    msg.refresh()
                except Exception:
                    pass
                
                # Skip if already processed (marked with PAIRActioned)
                if any((c or '').startswith('PAIRActioned') for c in (msg.categories or [])):
                    # Silently skip already processed messages
                    processed_messages.add(dedup_key)
                    continue
                
                # If this solitary item is unread, process it as a last resort
                if not msg.is_read:
                    # Internal-reply detection (rare path)
                    sender_addr = (msg.sender.address or "").lower() if msg.sender else ""
                    sender_is_staff = sender_addr in [e.lower() for e in EMAILS_TO_FORWARD]
                    is_automated_reply = bool(re.search(fr"{REPLY_ID_TAG}\s*([^\s<]+)", msg.body or "", flags=re.I|re.S))
                    if sender_is_staff and is_automated_reply and not any((c or '').startswith('PAIRActioned') for c in (msg.categories or [])):
                        handle_internal_reply(msg)
                        processed_messages.add(dedup_key)
                        continue

                    body_to_analyze = get_clean_message_text(msg)

                    print(f"\nNEW:  [{name}] {msg.received.strftime('%Y-%m-%d %H:%M')} | "
                          f"{msg.sender.address if msg.sender else 'UNKNOWN'} | {msg.subject}")
                    result = classify_email(msg.subject, body_to_analyze)
                    if HUMAN_CHECK: 
                        print(json.dumps(result, indent=2))
                        # Skip if this email is already in pending queue (prevent duplicates)
                        email_id = msg.object_id
                        if email_id not in pending_emails:
                            pending_emails[email_id] = {
                                "subject": msg.subject,
                                "body": body_to_analyze,
                                "classification": result,
                                "sender": msg.sender.address,
                                "received": msg.received.strftime('%Y-%m-%d %H:%M'),
                                "message_obj": msg
                            }
                            print(f"📧 Email stored for approval: {msg.subject}")
                        else:
                            print(f"⏭️  Email already in pending queue, skipping: {msg.subject}")
                        processed_messages.add(dedup_key)
                    else: 
                        print(json.dumps(result, indent=2))
                        handle_new_email(msg, result)
                        processed_messages.add(dedup_key)
                continue  # done with this delta item

            # Normal path: fetch unread messages in this conversation
            try:
                unread_msgs = unread_in_conversation(folder, mailbox, conv_id)
            except Exception as e:
                print(f"   Could not fetch unread children for {conv_id}: {e}")
                continue

            if not unread_msgs:
                # No unread children → nothing to do for this conversation
                continue

            # Process each unread child once, oldest -> newest
            for child in unread_msgs:
                dedup_key = getattr(child, 'internet_message_id', None) or child.object_id
                if dedup_key in processed_messages:
                    # Silently skip already processed messages
                    continue

                # Make sure we have up-to-date fields on the child
                try:
                    child.refresh()
                except Exception:
                    pass

                # Skip if already processed (marked with PAIRActioned)
                if any((c or '').startswith('PAIRActioned') for c in (child.categories or [])):
                    # Silently skip already processed messages
                    processed_messages.add(dedup_key)
                    continue

                # 1) Internal reply path (staff replies captured by your hidden REPLY_ID_TAG)
                sender_addr = (child.sender.address or "").lower() if child.sender else ""
                sender_is_staff = sender_addr in [e.lower() for e in EMAILS_TO_FORWARD]
                is_automated_reply = bool(re.search(fr"{REPLY_ID_TAG}\s*([^\s<]+)", child.body or "", flags=re.I|re.S))
                if sender_is_staff and is_automated_reply and not any(
                    (c or '').startswith('PAIRActioned') for c in (child.categories or [])
                ):
                    handle_internal_reply(child)
                    processed_messages.add(dedup_key)
                    continue

                # 3) Classify using reply-only text, then handle
                body_to_analyze = get_clean_message_text(child)
                print(f"\nNEW:  [{name}] {child.received.strftime('%Y-%m-%d %H:%M')} | "
                      f"{child.sender.address if child.sender else 'UNKNOWN'} | {child.subject}")
                result = classify_email(child.subject, body_to_analyze)
                print(json.dumps(result, indent=2))
                
                if HUMAN_CHECK:
                    # Skip if this email is already in pending queue (prevent duplicates)
                    email_id = child.object_id
                    if email_id not in pending_emails:
                        pending_emails[email_id] = {
                            "subject": child.subject,
                            "body": body_to_analyze,
                            "classification": result,
                            "sender": child.sender.address,
                            "received": child.received.strftime('%Y-%m-%d %H:%M'),
                            "message_obj": child
                        }
                        print(f"📧 Email stored for approval: {child.subject}")
                    else:
                        print(f"⏭️  Email already in pending queue, skipping: {child.subject}")
                else:
                    handle_new_email(child, result)

                # 4) Dedup remember
                processed_messages.add(dedup_key)

        # Debug: Check what attributes msgs actually has
        print(f"🔍 msgs object type: {type(msgs)}")
        print(f"🔍 msgs attributes: {[attr for attr in dir(msgs) if 'delta' in attr.lower() or 'token' in attr.lower()]}")
        
        # Try different ways to get delta token
        delta_attrs = ['delta_token', 'deltaToken', '_delta_token', 'next_link', 'delta_link', '@odata.deltaLink']
        new_delta_token = None
        for attr in delta_attrs:
            token_val = getattr(msgs, attr, None)
            if token_val:
                print(f"🎯 Found delta token in attribute '{attr}': {str(token_val)[:50]}...")
                new_delta_token = token_val
                break
        
        if not new_delta_token:
            new_delta_token = delta_token  # Return old token if no new one found
            
        print(f"🎯 Delta token for {name}: {'Found' if new_delta_token else 'None'}")
        return new_delta_token

    except requests.exceptions.HTTPError as e:
        # Catch the specific HTTP error for delta token invalidation
        print(f"🚨 HTTP Error in {name}: Status {e.response.status_code} - {e}")
        if e.response.status_code == 410:
            print(f"⚠️ Delta token for {name} is invalid (410 Gone). Initiating full re-sync...")
            # A 410 error means the token is permanently invalid.
            # We must fall back to a full query without any token.
            try:
                # Perform a full query from START_TIME
                fallback_qs = (folder.new_query()
                               .on_attribute('receivedDateTime').greater_equal(START_TIME)
                               .select([
                                   'id', 'conversationId', 'isRead', 'receivedDateTime', 
                                   'from', 'sender', 'subject', 'categories'
                               ]))
                
                print(f"🔄 Performing full re-sync for {name} from {START_TIME}...")
                msgs = folder.get_messages(query=fallback_qs)

                # Process the messages from the full query (same logic as above)
                for msg in msgs:
                    conv_id = getattr(msg, 'conversation_id', None)
                    if not conv_id:
                        dedup_key = getattr(msg, 'internet_message_id', None) or msg.object_id
                        if dedup_key in processed_messages:
                            continue
                        try:
                            msg.refresh()
                        except Exception:
                            pass
                        
                        if any((c or '').startswith('PAIRActioned') for c in (msg.categories or [])):
                            processed_messages.add(dedup_key)
                            continue
                        
                        if not msg.is_read:
                            sender_addr = (msg.sender.address or "").lower() if msg.sender else ""
                            sender_is_staff = sender_addr in [e.lower() for e in EMAILS_TO_FORWARD]
                            is_automated_reply = bool(re.search(fr"{REPLY_ID_TAG}\s*([^\s<]+)", msg.body or "", flags=re.I|re.S))
                            if sender_is_staff and is_automated_reply and not any((c or '').startswith('PAIRActioned') for c in (msg.categories or [])):
                                handle_internal_reply(msg)
                                processed_messages.add(dedup_key)
                                continue

                            body_to_analyze = get_clean_message_text(msg)

                            print(f"\nNEW:  [{name}] {msg.received.strftime('%Y-%m-%d %H:%M')} | "
                                  f"{msg.sender.address if msg.sender else 'UNKNOWN'} | {msg.subject}")
                            result = classify_email(msg.subject, body_to_analyze)
                            if HUMAN_CHECK: 
                                print(json.dumps(result, indent=2))
                                email_id = msg.object_id
                                if email_id not in pending_emails:
                                    pending_emails[email_id] = {
                                        "subject": msg.subject,
                                        "body": body_to_analyze,
                                        "classification": result,
                                        "sender": msg.sender.address,
                                        "received": msg.received.strftime('%Y-%m-%d %H:%M'),
                                        "message_obj": msg
                                    }
                                    print(f"📧 Email stored for approval: {msg.subject}")
                                processed_messages.add(dedup_key)
                            else: 
                                print(json.dumps(result, indent=2))
                                handle_new_email(msg, result)
                                processed_messages.add(dedup_key)
                        continue

                    # Normal path: fetch unread messages in this conversation
                    try:
                        unread_msgs = unread_in_conversation(folder, mailbox, conv_id)
                    except Exception as e:
                        print(f"   Could not fetch unread children for {conv_id}: {e}")
                        continue

                    if not unread_msgs:
                        continue

                    # Process each unread child once, oldest -> newest
                    for child in unread_msgs:
                        dedup_key = getattr(child, 'internet_message_id', None) or child.object_id
                        if dedup_key in processed_messages:
                            continue

                        try:
                            child.refresh()
                        except Exception:
                            pass

                        if any((c or '').startswith('PAIRActioned') for c in (child.categories or [])):
                            processed_messages.add(dedup_key)
                            continue

                        sender_addr = (child.sender.address or "").lower() if child.sender else ""
                        sender_is_staff = sender_addr in [e.lower() for e in EMAILS_TO_FORWARD]
                        is_automated_reply = bool(re.search(fr"{REPLY_ID_TAG}\s*([^\s<]+)", child.body or "", flags=re.I|re.S))
                        if sender_is_staff and is_automated_reply and not any(
                            (c or '').startswith('PAIRActioned') for c in (child.categories or [])
                        ):
                            handle_internal_reply(child)
                            processed_messages.add(dedup_key)
                            continue

                        body_to_analyze = get_clean_message_text(child)
                        print(f"\nNEW:  [{name}] {child.received.strftime('%Y-%m-%d %H:%M')} | "
                              f"{child.sender.address if child.sender else 'UNKNOWN'} | {child.subject}")
                        result = classify_email(child.subject, body_to_analyze)
                        print(json.dumps(result, indent=2))
                        
                        if HUMAN_CHECK:
                            email_id = child.object_id
                            if email_id not in pending_emails:
                                pending_emails[email_id] = {
                                    "subject": child.subject,
                                    "body": body_to_analyze,
                                    "classification": result,
                                    "sender": child.sender.address,
                                    "received": child.received.strftime('%Y-%m-%d %H:%M'),
                                    "message_obj": child
                                }
                                print(f"📧 Email stored for approval: {child.subject}")
                        else:
                            handle_new_email(child, result)

                        processed_messages.add(dedup_key)

                # Now, return the new, fresh delta token from this successful fallback query
                new_token = getattr(msgs, 'delta_token', None)
                print(f"✅ Successfully re-synced {name} and got fresh delta token.")
                return new_token
                
            except Exception as fallback_e:
                print(f"❌ Fallback re-sync failed for {name}: {fallback_e}")
                # Critical failure: unable to even perform a full sync.
                # Returning None will trigger a re-try on the next loop
                return None
        else:
            # For any other HTTP errors (e.g., 401, 500), log and return old token
            print(f"❌ HTTP error accessing {name}: Status {e.response.status_code} - {e}")
            return delta_token

    except Exception as e:
        # This catches all other non-HTTP errors (e.g., network issues, coding bugs)
        print(f"❌ General error accessing {name}: {e}")
        print(f"❌ Error type: {type(e).__name__}")
        import traceback
        print(f"❌ Full traceback: {traceback.format_exc()}")
        return delta_token


def handle_new_email(msg, result):
    """
    Takes a message and its AI classification result, then acts on it.
    It does NOT call the AI again.
    """
    categories = result.get("categories", [])
    recipients_set = set(result.get("all_recipients", []))
    name_sender = result.get("name_sender")
    
    # We pass the message and its categories to be tagged
    tag_email(msg, categories, replyTag=False)
    # We use the results to perform specific actions
    handle_emails(categories, result, recipients_set, msg, name_sender)

    if recipients_set:
        fwd = msg.forward()
        recipients_list = list(recipients_set)
        fwd.to.add(recipients_list)
        
        # Store recipients for later CC functionality
        forwarded_recipients[msg.object_id] = recipients_list
        
        # Add the hidden tracking ID into the top of the forwarded body
        instruction_html = f"""<div style="display:none;">{REPLY_ID_TAG}{msg.object_id}</div>"""
        
        # Prepend to the auto-generated forward body
        fwd.body = "Please press 'Reply All,' and reply to info@mlfa.org. You're email will automatically be sent to the correct person. " + instruction_html
        fwd.body_type = 'HTML'
        
        # Send the forward
        fwd.send()


    if not set(categories).issubset(NONREAD_CATEGORIES):
        mark_as_read(msg)


def get_time_based_greeting(name_sender):
    """Return appropriate greeting based on time of day and sender name"""
    if name_sender and name_sender != "Sender":
        return f"Dear {name_sender},"
    
    # Use time-based greeting when name is not available
    current_hour = datetime.now().hour
    if 5 <= current_hour < 12:
        return "Good morning,"
    elif 12 <= current_hour < 17:
        return "Good afternoon,"
    elif 17 <= current_hour < 21:
        return "Good evening,"
    else:
        return "Good morning,"  # Default to good morning for late night/early morning

def handle_emails(categories, result, recipients_set, msg, name_sender): 
    for category in categories:
        if category == "legal":
            reply_message = msg.reply(to_all=False)
            # Check if this email needs a personal reply based on classification
            needs_personal = result.get("needs_personal_reply", False)
            greeting = get_time_based_greeting(name_sender)

            if needs_personal:
                reply_message.body = f"""
                    <p>{greeting}</p>

                    <p>Thank you for contacting the Muslim Legal Fund of America (MLFA). 
                    We are grateful that you reached out and placed your trust in us to potentially support your legal matter.</p>

                    <p>If you have not already done so, please submit a formal application for legal assistance through our website:<br>
                    <a href="https://mlfa.org/application-for-legal-assistance/">https://mlfa.org/application-for-legal-assistance/</a></p>

                    <p>Once submitted, our team will carefully review your application and follow up with next steps. 
                    If you have any questions about the application process or need help completing it, please don't hesitate to reach out.</p>

                    <p>We appreciate your patience as we work through applications, and we look forward to learning more about how we might be able to help.</p>

                    <p>Warm regards,<br>
                    The MLFA Team<br>
                    Muslim Legal Fund of America</p>
                """

                reply_message.body_type = "HTML"
            else:
                reply_message.body = f"""
                    <p>{greeting}</p>

                    <p>Thank you for contacting the Muslim Legal Fund of America (MLFA).</p>

                    <p>If you have not already done so, please submit a formal application for legal assistance 
                    through our website:<br>
                    <a href="https://mlfa.org/application-for-legal-assistance/">https://mlfa.org/application-for-legal-assistance/</a></p>

                    <p>This ensures our legal team has the information needed to review your case promptly.</p>

                    <p>Sincerely,<br>
                    The MLFA Team</p>
                """

                reply_message.body_type = "HTML"
            reply_message.send()
            
            # Move to Apply for help folder
            inbox = mailbox.inbox_folder()
            try:
                apply_folder = inbox.get_folder(folder_name="Apply for help")
                print("Moving to Apply for help folder.")
                msg.move(apply_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Apply for help folder: {e}")

        elif category == "donor":
            recipients_set.update([f"{EMAILS_TO_FORWARD[0]}", f"{EMAILS_TO_FORWARD[1]}"])
            # Move to Doner_Related folder
            inbox = mailbox.inbox_folder()
            try:
                donor_folder = inbox.get_folder(folder_name="Doner_Related")
                print("Moving to Doner_Related folder.")
                msg.move(donor_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Doner_Related folder: {e}")

        elif category == "sponsorship":
            recipients_set.update([f"{EMAILS_TO_FORWARD[2]}", f"{EMAILS_TO_FORWARD[3]}"])

        elif category == "organizational":
            recipients_set.update([f"{EMAILS_TO_FORWARD[2]}", f"{EMAILS_TO_FORWARD[3]}"])
            # Move to Organizational inquiries folder
            inbox = mailbox.inbox_folder()
            try:
                org_folder = inbox.get_folder(folder_name="Organizational inquiries")
                print("Moving to Organizational inquiries folder.")
                msg.move(org_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Organizational inquiries folder: {e}")

        elif category == "volunteer":
            # Send automated reply with volunteer application form instead of forwarding
            reply_message = msg.reply(to_all=False)
            greeting = get_time_based_greeting(name_sender)
            
            reply_message.body = f"""
                <p>{greeting}</p>

                <p>Thank you for your interest in volunteering with the Muslim Legal Fund of America (MLFA)!</p>

                <p>We are grateful for your willingness to support our mission of providing legal assistance to Muslims in need. To get started with the volunteer process, please complete our volunteer application form:</p>

                <p><a href="https://forms.office.com/Pages/ResponsePage.aspx?id=oiB_iSDzkUu20kpWPbd_DnxSOj2KmWxOomg5Rm0KtBNUMElYQkdOQUU2WUxLTlNHMkY4S0tFOU1XViQlQCN0PWcu">MLFA Volunteer Application Form</a></p>

                <p>Once you submit the form, our team will review your application and follow up with next steps about volunteer opportunities that match your skills and interests.</p>

                <p>Thank you again for your support!</p>

                <p>Best regards,<br>
                The MLFA Team<br>
                Muslim Legal Fund of America</p>
            """
            
            reply_message.body_type = "HTML"
            reply_message.send()
            
            # Move to Volunteer folder
            inbox = mailbox.inbox_folder()
            try:
                volunteer_folder = inbox.get_folder(folder_name="Volunteer")
                print("Moving to Volunteer folder.")
                msg.move(volunteer_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Volunteer folder: {e}")

        elif category == "internship":
            recipients_set.update([f"{EMAILS_TO_FORWARD[5]}"])
            # Move to Internship folder
            inbox = mailbox.inbox_folder()
            try:
                internship_folder = inbox.get_folder(folder_name="Internship")
                print("Moving to Internship folder.")
                msg.move(internship_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Internship folder: {e}")

        elif category == "job_application":
            recipients_set.update([f"{EMAILS_TO_FORWARD[6]}"])
            # Move to Job_Application folder
            inbox = mailbox.inbox_folder()
            try:
                job_folder = inbox.get_folder(folder_name="Job_Application")
                print("Moving to Job_Application folder.")
                msg.move(job_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Job_Application folder: {e}")

        elif category == "fellowship":
            recipients_set.update([f"{EMAILS_TO_FORWARD[5]}"])
            # Move to Fellowship folder
            inbox = mailbox.inbox_folder()
            try:
                fellowship_folder = inbox.get_folder(folder_name="Fellowship")
                print("Moving to Fellowship folder.")
                msg.move(fellowship_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Fellowship folder: {e}")
            
        elif category == "media":
            recipients_set.update([f"{EMAILS_TO_FORWARD[7]}"])  # Marium.Uddin@mlfa.org
            # Move to Media folder
            inbox = mailbox.inbox_folder()
            try:
                media_folder = inbox.get_folder(folder_name="Media")
                print("Moving to Media folder.")
                msg.move(media_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Media folder: {e}")

        elif category == "marketing":
            inbox = mailbox.inbox_folder()
            try:
                sales_folder = inbox.get_folder(folder_name="Sales emails")
                print("Moving to Sales emails folder.")
                msg.move(sales_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Sales emails folder: {e}")
        
        elif category == "cold_outreach":
            inbox = mailbox.inbox_folder()
            try:
                irrelevant_folder = inbox.get_folder(folder_name="Irrelevant")
                cold_outreach_folder = irrelevant_folder.get_folder(folder_name="Cold_Outreach")
                print("Moving to Irrelevant/Cold_Outreach folder.")
                msg.move(cold_outreach_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Irrelevant/Cold_Outreach folder: {e}")
        
        elif category == "spam":
            inbox = mailbox.inbox_folder()
            try:
                irrelevant_folder = inbox.get_folder(folder_name="Irrelevant")
                spam_folder = irrelevant_folder.get_folder(folder_name="Spam")
                print("Moving to Irrelevant/Spam folder.")
                msg.move(spam_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Irrelevant/Spam folder: {e}")
        
        elif category == "newsletter":
            inbox = mailbox.inbox_folder()
            try:
                for_reference_folder = inbox.get_folder(folder_name="For reference")
                newsletter_folder = for_reference_folder.get_folder(folder_name="subscriptions and newsletters")
                print("Moving to For reference/subscriptions and newsletters folder.")
                msg.move(newsletter_folder)
            except Exception as e:
                print(f"⚠️ Could not move to For reference/subscriptions and newsletters folder: {e}")
        
        elif category == "irrelevant_other":
            inbox = mailbox.inbox_folder()
            try:
                irrelevant_folder = inbox.get_folder(folder_name="Irrelevant")
                other_folder = irrelevant_folder.get_folder(folder_name="Other")
                print("Moving to Irrelevant/Other folder.")
                msg.move(other_folder)
            except Exception as e:
                print(f"⚠️ Could not move to Irrelevant/Other folder: {e}")

def tag_email(msg, categories, replyTag):
    # 1) Load existing categories safely
    existing = set((msg.categories or []))

    # 2) Build new tags for this operation
    new_tags = set()
    for c in categories or []:
        c = (c or "").strip()
        if not c:
            continue
        if replyTag:
            new_tags.add(f"PAIRActioned/replied/{c}")
        else:
            if c in ('spam', 'cold_outreach', 'newsletter'):
                new_tags.add(f"PAIRActioned/irrelevant/{c}")
            else:
                new_tags.add(f"PAIRActioned/{c}")

    # Always keep the umbrella marker
    new_tags.add("PAIRActioned")

    # 3) Merge (union) — do NOT drop existing tags
    merged = existing.union(new_tags)

    # 4) Save only if there’s a change
    if merged != existing:
        msg.categories = sorted(merged)
        msg.save_message()


def mark_as_read(msg): 
    print("   Marking email as read...")
    try:
        msg.mark_as_read()
        print("   Marked as read")
    except Exception as e:
        print(f"    Could not mark as read: {e}")

def handle_internal_reply(msg): 
    print(f"\nREPLY DETECTED: From {msg.sender.address} | {msg.subject}")
    body_parts = msg.body.split(REPLY_ID_TAG)
    if len(body_parts) < 2:
        print(" ERROR: Could not find the reply id, therefore, we cannot reply. ")
        return

    html_chunk = body_parts[0]
    soup = BeautifulSoup(html_chunk, 'html.parser')
    reply_content = str(soup)

    if not reply_content: 
        print("   WARNING: Reply appears to be empty. Not sending. ")
        #We need to maybe re-email the person who wrote the reply to the forwarded email to try again. 
        return

    match = re.search(f"{REPLY_ID_TAG}(.+?)</", msg.body)
    if not match: 
        print("   ERROR: Could not find the original message ID.")
        return
    original_message_id = match.group(1).strip()
    
    try:
        original_msg = mailbox.get_message(original_message_id)
        
        # Get other recipients to CC from our stored dictionary
        other_forwardees = []
        sender_email = msg.sender.address.lower()
        if original_message_id in forwarded_recipients:
            all_recipients = forwarded_recipients[original_message_id]
            other_forwardees = [email for email in all_recipients if email.lower() != sender_email]
            print(f"   Found recipients: {all_recipients}, will CC: {other_forwardees}")
        
        # Create the reply
        final_reply = original_msg.reply(to_all=False)
        final_reply.body = reply_content
        final_reply.body_type = "HTML"
        
        # CC the other forwardees if any
        if other_forwardees:
            for cc_email in other_forwardees:
                final_reply.cc.add(cc_email)
            print(f"   Sent reply to original sender: {original_msg.sender.address}, CC'd: {other_forwardees}")
        else:
            print(f"   Sent reply to original sender: {original_msg.sender.address}")
        
        final_reply.send()
    except Exception as e:
        print(f"   ERROR: Could not send final reply. Error: {e}")
        return

    msg.mark_as_read()
    print("   Cleanup complete. Reply process finished.")


def newest_unread_in_conversation(folder, mailbox, conversation_id):
    """
    Return the newest unread message in the given conversation (or None).
    Uses server-side filter/order and limits to 1 item.
    """
    if not conversation_id:
        return None

    # Build the query FIRST
    q = (mailbox.new_query()
         .on_attribute('conversationId').equals(conversation_id)
         .chain('and').on_attribute('isRead').equals(False)
         .order_by('receivedDateTime', ascending=False)  # newest first
         .select([
             # Use Graph field names in $select (camelCase)
             'id', 'conversationId', 'internetMessageId',
             'isRead', 'receivedDateTime',
             'from', 'sender', 'subject',
             'categories', 'uniqueBody', 'body'
         ]))

    # Then execute with a limit of 1 (SDK-compatible way to "top(1)")
    items = list(folder.get_messages(query=q, limit=1, order_by='receivedDateTime desc'))
    if not items:
        return None

    msg = items[0]

    # Hydrate to ensure properties like categories/unique_body are fresh
    try:
        # If you prefer a full re-fetch by id instead of refresh():
        # msg = folder.get_message(object_id=msg.object_id) or msg
        msg.refresh()
    except Exception:
        pass

    return msg


def unread_in_conversation(folder, mailbox, conversation_id, page_limit=30):
    if not conversation_id:
        return []

    q = (mailbox.new_query()
         .on_attribute('conversationId').equals(conversation_id)
         .select([
             'id','conversationId','internetMessageId',
             'isRead','receivedDateTime',
             'from','sender','subject',
             'categories','uniqueBody','body'
         ]))

    items = list(folder.get_messages(query=q, limit=page_limit))
    unread = [m for m in items if not getattr(m, 'is_read', False)]
    unread.sort(key=lambda m: m.received or m.created)  # oldest→newest unread
    for m in unread:
        try: m.refresh()
        except: pass
    return unread

def get_clean_message_text(msg):
    """
    Return only the reply content for this message.
    Prefer Graph's unique_body (just the new text),
    otherwise strip quoted history from the full body.
    """
    QUOTE_SEPARATORS = [
        r'^\s*On .* wrote:\s*$',              
        r'^\s*From:\s.*$',                    
        r'^\s*-----Original Message-----\s*$',
        r'^\s*De:\s.*$',                      
        r'^\s*Sent:\s.*$',
        r'^\s*To:\s.*$',
    ]

    def strip_quoted_reply(html_or_text: str) -> str:
        if not html_or_text:
            return ""
        text = html_or_text
        try:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(html_or_text, 'html.parser')
            for sel in [
                'blockquote',
                'div.gmail_quote',
                'div[type=cite]',
                'div.moz-cite-prefix',
                'div.OutlookMessageHeader',
            ]:
                for node in soup.select(sel):
                    node.decompose()
            text = soup.get_text("\n")
        except Exception:
            pass

        import re
        lines = [ln.rstrip() for ln in text.splitlines()]
        out = []
        for ln in lines:
            if ln.strip().startswith('>'):
                break
            if any(re.match(pat, ln, flags=re.IGNORECASE) for pat in QUOTE_SEPARATORS):
                break
            out.append(ln)

        return "\n".join(out).strip()[:8000]

    body = getattr(msg, 'unique_body', None) or getattr(msg, 'body', None) or ""
    return strip_quoted_reply(body)



# Login page template
LOGIN_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>MLFA Email Hub - Login</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;600&family=Inter:wght@300;400;500;600&display=swap');
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: #0d1117;
            color: #f0f6fc;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0;
            font-size: 14px;
            line-height: 1.5;
        }
        
        .login-container {
            width: 100%;
            max-width: 400px;
            padding: 20px;
        }
        
        .login-header {
            text-align: center;
            margin-bottom: 32px;
        }
        
        .login-header h1 {
            font-family: 'JetBrains Mono', monospace;
            font-size: 18px;
            font-weight: 500;
            color: #58a6ff;
            margin: 0;
            letter-spacing: -0.025em;
        }
        
        .login-box {
            background: #161b22;
            padding: 32px;
            border-radius: 6px;
            border: 1px solid #30363d;
            width: 100%;
        }
        
        .login-box input[type="password"] {
            width: 100%;
            padding: 12px;
            margin-bottom: 16px;
            background: #0d1117;
            border: 1px solid #30363d;
            border-radius: 4px;
            color: #f0f6fc;
            font-size: 14px;
            font-family: 'JetBrains Mono', monospace;
            transition: border-color 0.15s ease;
        }
        
        .login-box input[type="password"]:focus {
            outline: none;
            border-color: #58a6ff;
        }
        
        .login-box input[type="password"]::placeholder {
            color: #7d8590;
            font-family: 'JetBrains Mono', monospace;
        }
        
        .login-box button {
            width: 100%;
            padding: 8px 16px;
            background: #238636;
            border: 1px solid #238636;
            border-radius: 4px;
            color: #f0f6fc;
            font-size: 13px;
            font-weight: 500;
            cursor: pointer;
            font-family: 'Inter', sans-serif;
            transition: all 0.15s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 32px;
        }
        
        .login-box button:hover {
            background: #2ea043;
            border-color: #2ea043;
        }
        
        .error {
            color: #f85149;
            margin-bottom: 16px;
            font-size: 13px;
            font-family: 'Inter', sans-serif;
            text-align: center;
            padding: 8px 12px;
            background: #0d1117;
            border: 1px solid #da3633;
            border-radius: 4px;
        }
        
        @media (max-width: 768px) {
            .login-container {
                padding: 16px;
            }
            .login-box {
                padding: 24px;
            }
        }
    </style>
</head>
<body>
    <div class="login-container">
        <div class="login-header">
            <h1>Approval Hub for info@mlfa.org</h1>
        </div>
        <div class="login-box">
            {% if error %}
            <div class="error">{{ error }}</div>
            {% endif %}
            <form method="post">
                <input type="password" name="password" placeholder="enter password" required autofocus>
                <button type="submit">Login</button>
            </form>
        </div>
    </div>
</body>
</html>
'''

# Flask routes
@app.route('/login', methods=['GET', 'POST'])
def login():
    client_ip = request.remote_addr
    
    # Check if IP is locked out
    if client_ip in login_attempts:
        attempts, last_attempt = login_attempts[client_ip]
        if attempts >= MAX_LOGIN_ATTEMPTS:
            if time.time() - last_attempt < LOCKOUT_TIME:
                remaining = int(LOCKOUT_TIME - (time.time() - last_attempt))
                return render_template_string(LOGIN_TEMPLATE, 
                    error=f'Too many failed attempts. Try again in {remaining} seconds')
            else:
                # Reset after lockout period
                del login_attempts[client_ip]
    
    if request.method == 'POST':
        password = request.form.get('password')
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        
        if password_hash == ADMIN_PASSWORD_HASH:
            # Clear failed attempts on successful login
            if client_ip in login_attempts:
                del login_attempts[client_ip]
            
            session['logged_in'] = True
            session['login_time'] = time.time()
            session.permanent = True
            print(f"✅ User logged in successfully from {client_ip}")
            print(f"📍 Session after login: logged_in={session.get('logged_in')}, login_time={session.get('login_time')}")
            return redirect(url_for('index'))
        else:
            # Track failed attempt
            if client_ip not in login_attempts:
                login_attempts[client_ip] = (1, time.time())
            else:
                attempts, _ = login_attempts[client_ip]
                login_attempts[client_ip] = (attempts + 1, time.time())
            
            attempts = login_attempts[client_ip][0]
            remaining = MAX_LOGIN_ATTEMPTS - attempts
            
            if remaining > 0:
                error = f'Invalid password. {remaining} attempts remaining'
            else:
                error = f'Account locked. Try again in {LOCKOUT_TIME} seconds'
            
            print(f"❌ Failed login attempt {attempts}/{MAX_LOGIN_ATTEMPTS} from {client_ip}")
            return render_template_string(LOGIN_TEMPLATE, error=error)
    
    return render_template_string(LOGIN_TEMPLATE)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    print(f"📍 Index route accessed - Session: logged_in={session.get('logged_in')}, IP={request.remote_addr}")
    return send_from_directory('.', 'approval-hub.html')

@app.route('/api/emails')
@login_required
def get_emails():
    print(f"📧 API call to /api/emails - pending emails count: {len(pending_emails)}")
    emails = []
    for email_id, email_data in pending_emails.items():
        # Format data for the interface
        categories = email_data["classification"].get('categories', [])
        category_display = ', '.join([cat.replace('_', ' ').title() for cat in categories])
        
        recipients = email_data["classification"].get('all_recipients', [])
        recipients_display = ', '.join(recipients) if recipients else 'None'
        
        reasons = email_data["classification"].get('reason', {})
        reason_display = '; '.join(reasons.values()) if reasons else 'No reason provided'
        
        email = {
            "id": email_id,
            "meta": f"FROM: [INBOX] {email_data['received']} | {email_data['sender']} | {email_data['subject']}",
            "senderName": email_data["classification"].get('name_sender', 'Unknown'),
            "category": category_display,
            "recipients": recipients_display,
            "needsReply": "Yes" if email_data["classification"].get('needs_personal_reply', False) else "No",
            "reason": reason_display,
            "escalation": email_data["classification"].get('escalation_reason') or 'None',
            "originalContent": email_data["body"],
            "status": "pending"
        }
        emails.append(email)
    return jsonify(emails)

@app.route('/api/emails/<email_id>/approve', methods=['POST'])
@login_required
def approve_email(email_id):
    if email_id in pending_emails:
        email_data = pending_emails[email_id]
        msg = email_data["message_obj"]
        classification = email_data["classification"]
        
        print(f"✅ Email approved: {email_data['subject']} - Processing normally")
        
        # Process the email normally using the stored message and classification
        handle_new_email(msg, classification)
        
        # Add to processed messages to prevent reappearance
        dedup_key = getattr(msg, 'internet_message_id', None) or msg.object_id
        processed_messages.add(dedup_key)
        processed_messages.add(email_id)  # Also add the email_id itself
        
        # Remove from pending emails
        del pending_emails[email_id]
        
    return jsonify({"status": "success", "message": "Email approved and will be processed normally"})

@app.route('/api/emails/<email_id>/reject', methods=['POST'])
@login_required
def reject_email(email_id):
    data = request.get_json()
    reason = data.get('reason', 'No reason provided')
    
    if email_id in pending_emails:
        email_data = pending_emails[email_id]
        msg = email_data["message_obj"]
        print(f"❌ Email rejected: {email_data['subject']} - Reason: {reason}")
        
        # Move rejected email to special folder
        try:
            inbox = mailbox.inbox_folder()
            rejected_folder = None
            
            # Try to find existing "declined" folder
            try:
                rejected_folder = inbox.get_folder(folder_name="declined")
                print(f"📁 Found existing 'declined' folder")
            except:
                try:
                    rejected_folder = inbox.get_folder(folder_name="Declined")
                    print(f"📁 Found existing 'Declined' folder")
                except:
                    print(f"⚠️ Could not find 'declined' or 'Declined' folder")
            
            # Move the email to rejected folder
            if rejected_folder:
                try:
                    msg.move(rejected_folder)
                    print(f"📁 Moved email to 'declined' folder")
                except Exception as e:
                    print(f"⚠️ Could not move email to rejected folder: {e}")
            
            # Just mark as processed without adding new tags
            try:
                existing_cats = set(msg.categories or [])
                existing_cats.add("PAIRActioned")  # Only add the standard processed marker
                msg.categories = sorted(existing_cats)
                msg.save_message()
                print(f"📋 Marked email as processed (rejected)")
            except Exception as e:
                print(f"⚠️ Could not update email categories: {e}")
            
            
            # Mark as read and add to processed messages to prevent reprocessing
            try:
                mark_as_read(msg)
                dedup_key = getattr(msg, 'internet_message_id', None) or msg.object_id
                processed_messages.add(dedup_key)
                processed_messages.add(email_id)  # Also add the email_id itself
                print(f"✅ Marked email as processed")
            except Exception as e:
                print(f"⚠️ Could not mark as processed: {e}")
            
        except Exception as e:
            print(f"❌ Error handling rejected email: {e}")
        
        # Remove from pending emails
        del pending_emails[email_id]
    
    return jsonify({"status": "success", "message": f"Email rejected: {reason}"})


def start_web_server():
    """Start the Flask web server in a separate thread"""
    def run_server():
        print("🌐 Starting approval hub at http://localhost:5000")
        app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
    
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()


# Start the web server
start_web_server()

# Initialize delta tokens from the JSON store
inbox_delta = delta_tokens.get("inbox")
junk_delta = delta_tokens.get("junk")
print(f"📚 Starting with delta tokens - Inbox: {'✓' if inbox_delta else 'None'}, Junk: {'✓' if junk_delta else 'None'}")

# Check authentication status
print(f"🔐 Account authenticated: {account.is_authenticated}")
print(f"📧 Email to watch: {EMAIL_TO_WATCH}")
print(f"🗂️ Inbox folder: {inbox_folder}")
print(f"🗑️ Junk folder: {junk_folder}")

print(f"Monitoring inbox + junk for: {EMAIL_TO_WATCH} … Ctrl-C to stop.")
print(f"📧 Approval hub available at: http://localhost:5000")

def reconnect_account():
    """Re-authenticate and reconnect to mailbox"""
    global account, mailbox, inbox_folder, junk_folder
    
    print("🔄 Re-authenticating with Microsoft Graph API...")
    try:
        # Re-authenticate with fresh token
        if IS_PRODUCTION and CLIENT_SECRET:
            credentials = (CLIENT_ID, CLIENT_SECRET)
            account = Account(credentials, auth_flow_type="credentials", tenant_id=TENANT_ID)
            account.authenticate()  # Force new authentication
        else:
            credentials = (CLIENT_ID, None)
            token_backend = FileSystemTokenBackend(token_path=".", token_filename="o365_token.txt")
            account = Account(credentials, auth_flow_type="authorization", token_backend=token_backend)
            account.authenticate(scopes=['basic', 'message_all'])
        
        # Reconnect to mailbox
        mailbox = account.mailbox(resource=EMAIL_TO_WATCH)
        inbox_folder = mailbox.inbox_folder()
        junk_folder = mailbox.junk_folder()
        
        print("✅ Successfully re-authenticated!")
        return True
    except Exception as e:
        print(f"❌ Re-authentication failed: {e}")
        return False

consecutive_errors = 0
last_successful_check = time.time()

while True:
    try:
        print(f"🔄 Checking for new emails... (Pending: {len(pending_emails)}, Processed: {len(processed_messages)})")
        
        # Check if it's been too long since last successful check (1 hour)
        if time.time() - last_successful_check > 3600:
            print("⚠️ No successful checks in 1 hour, forcing re-authentication...")
            reconnect_account()
        
        # Revert to original approach but with better time window  
        print(f"🔄 Using improved O365 library approach...")
        
        # Process with a more recent time window to catch latest emails
        current_time = datetime.now(timezone.utc)
        recent_window = current_time - timedelta(hours=2)  # Last 2 hours
        
        print(f"📅 Processing emails from: {recent_window.isoformat()}")
        
        # Use original process_folder but with recent time window
        inbox_query = inbox_folder.new_query().on_attribute('receivedDateTime').greater_equal(recent_window)
        junk_query = junk_folder.new_query().on_attribute('receivedDateTime').greater_equal(recent_window)
        
        print(f"📡 Querying INBOX for recent emails...")
        inbox_msgs = list(inbox_folder.get_messages(query=inbox_query))
        print(f"📬 Found {len(inbox_msgs)} recent emails in INBOX")
        
        print(f"📡 Querying JUNK for recent emails...")
        junk_msgs = list(junk_folder.get_messages(query=junk_query))  
        print(f"📬 Found {len(junk_msgs)} recent emails in JUNK")
        
        # Process each message
        all_msgs = [(msg, "INBOX") for msg in inbox_msgs] + [(msg, "JUNK") for msg in junk_msgs]
        
        for msg, folder_name in all_msgs:
            try:
                print(f"\n🔍 Found email: {msg.subject} from {msg.sender.address if msg.sender else 'Unknown'}")
                
                # Skip if already read
                if msg.is_read:
                    print(f"⏭️ Skipping read email")
                    continue
                
                # Skip if already processed
                dedup_key = getattr(msg, 'internet_message_id', None) or msg.object_id
                if dedup_key in processed_messages:
                    print(f"⏭️ Skipping already processed email")
                    continue
                
                # Skip if already tagged as processed
                if any((c or '').startswith('PAIRActioned') for c in (msg.categories or [])):
                    print(f"⏭️ Skipping already tagged email")
                    continue
                
                # Process the email
                body_to_analyze = msg.body or ""
                print(f"🎯 Processing: {msg.subject}")
                
                result = classify_email(msg.subject, body_to_analyze)
                print(json.dumps(result, indent=2))
                
                if HUMAN_CHECK:
                    email_id = msg.object_id
                    if email_id not in pending_emails:
                        pending_emails[email_id] = {
                            "subject": msg.subject,
                            "body": body_to_analyze,
                            "classification": result,
                            "sender": msg.sender.address if msg.sender else "Unknown",
                            "received": msg.received.strftime('%Y-%m-%d %H:%M') if msg.received else "Unknown",
                            "message_obj": msg
                        }
                        print(f"📧 Email stored for approval: {msg.subject}")
                    processed_messages.add(dedup_key)
                else:
                    handle_new_email(msg, result)
                    processed_messages.add(dedup_key)
                    
            except Exception as e:
                print(f"❌ Error processing email: {e}")
                
        print(f"✅ Processed {len(all_msgs)} recent emails")
        
        # Reset error counter on success
        consecutive_errors = 0
        last_successful_check = time.time()
        
    except Exception as e:
        consecutive_errors += 1
        print(f"❌ Error in main loop (attempt {consecutive_errors}): {e}")
        
        # If we get 3 errors in a row, try to re-authenticate
        if consecutive_errors >= 3:
            print("⚠️ Multiple consecutive errors detected, attempting to reconnect...")
            if reconnect_account():
                consecutive_errors = 0
            else:
                print("😴 Waiting 60 seconds before retry...")
                time.sleep(60)
    
    time.sleep(10)
