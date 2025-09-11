from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta
import os, time, openai, json
import textwrap
import re
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import requests
import logging


load_dotenv()
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

### CONSTANTS

# Start from now - only process new emails going forward
START_TIME = datetime.now(timezone.utc) - timedelta(days=4)

CLIENT_ID = "c0abfd02-2166-4a52-b052-16d1aa084afb"  # MLFA app registration
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
TENANT_ID = os.getenv("O365_TENANT_ID")
REPLY_ID_TAG = "Pair_Reply_Reference_ID"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EMAIL_TO_WATCH = os.getenv("EMAIL_TO_WATCH")

PENDING_TAG = "PAIRActioned/pending_review"


EMAILS_TO_FORWARD = ['Mujahid.rasul@mlfa.org', 'Syeda.sadiqa@mlfa.org', 'Arshia.ali.khan@mlfa.org', 'Maria.laura@mlfa.org', 'info@mlfa.org', 'aisha.ukiu@mlfa.org', 'shawn@strategichradvisory.com', 'Marium.Uddin@mlfa.org']
NONREAD_CATEGORIES = {"marketing"}  # Keep these unread
SKIP_CATEGORIES = {'spam', 'cold_outreach', 'newsletter', 'irrelevant_other'}

HUMAN_CHECK = True  # Enable human check for approval hub

# Storage for multiple pending emails
pending_emails = {}  # Dictionary to store multiple emails by ID
current_email_id = None  # Track which email is currently being shown

# Storage for forwarded email recipients (for CC functionality)
forwarded_recipients = {}  # Maps message_id to list of recipients

###CONNECTING


from openai import OpenAI
client = OpenAI(api_key=OPENAI_API_KEY)

# Security settings - adjust based on environment
IS_PRODUCTION = os.getenv('ENVIRONMENT', 'development').lower() == 'production'

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
    account = Account(credentials, auth_flow_type="authorization", token_backend=token_backend, tenant_id=TENANT_ID)
    if not account.is_authenticated:
        account.authenticate(scopes=['basic', 'message_all'])

mailbox = account.mailbox(resource=EMAIL_TO_WATCH)
inbox_folder = mailbox.inbox_folder()
junk_folder = mailbox.junk_folder()




def read_token(path):
    if os.path.exists(path):
        with open(path, "r") as f:
            return f.read().strip()
    return None

def load_last_delta():
    # Create the full, correct path to the token files
    inbox_token_path = os.path.join(BASE_DIR, "delta_token_inbox.txt")
    junk_token_path = os.path.join(BASE_DIR, "delta_token_junk.txt")

    # READ the content from those paths
    inbox_token = read_token(inbox_token_path)
    junk_token = read_token(junk_token_path)
    
    # RETURN the actual tokens (the content)
    return inbox_token, junk_token


def save_last_delta(inbox_token, junk_token):
    # Create the full, correct path for saving
    inbox_token_path = os.path.join(BASE_DIR, "delta_token_inbox.txt")
    junk_token_path = os.path.join(BASE_DIR, "delta_token_junk.txt")
    print("Made it", inbox_token)

    try:
        print("second")
        # CHECK if a valid token was passed in
        if inbox_token:
            print("third")
            # OPEN the correct full path
            with open(inbox_token_path, "w") as f:
                print("CHIECKEN")
                f.write(inbox_token)
        
        # CHECK if a valid token was passed in
        if junk_token:
            # OPEN the correct full path
            with open(junk_token_path, "w") as f:
                f.write(junk_token)

    except IOError as e:
        print(f"⚠️ CRITICAL: Could not save delta token to file! Error: {e}")


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
        parsed_result = json.loads(raw)
        
        # Ensure we always return a dictionary
        if not isinstance(parsed_result, dict):
            print(f"Warning: AI returned non-dict response: {type(parsed_result)}")
            return {}
        return parsed_result
    except Exception as e:
        print(f"Classification error: {e}")
        return {}


def process_folder(folder, name, delta_token_url):
    """
    Uses the Microsoft Graph /delta endpoint directly so we can:
      1) Page through all changes (via @odata.nextLink)
      2) Capture and return the final @odata.deltaLink

    `delta_token_url` should be either:
      - None (first run), or
      - The FULL @odata.deltaLink URL from a previous run (store it as-is).
    """
    import urllib.parse

    # --- Build the initial delta URL ---
    # Prefer the account's protocol base if available; fall back to the standard Graph URL.
    try:
        base_url = getattr(account.protocol, "service_url", None)
        if callable(base_url):
            base_url = account.protocol.service_url()
        if not base_url:
            base_url = "https://graph.microsoft.com/v1.0"
    except Exception:
        base_url = "https://graph.microsoft.com/v1.0"

    if delta_token_url:
        next_url = delta_token_url  # resume from last delta link (FULL URL)
        params = None
    else:
        # First sync: build a delta URL scoped to THIS folder
        # /users/{user}/mailFolders/{folder_id}/messages/delta
        folder_id = getattr(folder, 'object_id', None) or getattr(folder, 'folder_id', None) or getattr(folder, 'id', None)
        user_part = urllib.parse.quote(EMAIL_TO_WATCH)
        next_url = f"{base_url}/users/{user_part}/mailFolders/{folder_id}/messages/delta"
        # Select lean fields (Graph uses camelCase field names)
        params = {
            "$select": "id,conversationId,isRead,receivedDateTime,from,sender,subject,categories,uniqueBody,body",
            "$top": "50",
        }
        # Your START_TIME filter on first run
        if START_TIME:
            params["$filter"] = f"receivedDateTime ge {START_TIME.isoformat()}"

    final_delta_link = delta_token_url  # will be replaced when Graph returns a new one
    total_changed = 0

    try:
        # Use the O365 connection so auth headers/tokens are handled for us
        conn = account.connection

        # Accumulate changes page by page
        while next_url:
            resp = conn.get(next_url, params=params) if params else conn.get(next_url)
            if not resp or resp.status_code // 100 != 2:
                print(f" Error accessing {name}: {getattr(resp, 'status_code', 'n/a')} {getattr(resp, 'text', '')}")
                return delta_token_url or final_delta_link

            data = resp.json() or {}
            items = data.get("value", [])

            # --- DIAGNOSTIC ---
            print(f"DIAGNOSTIC [{name}]: Page returned {len(items)} item(s).")

            # Process each changed item exactly as before
            for item in items:
                total_changed += 1
                # We intentionally re-fetch the message object via SDK for your existing logic
                try:
                    msg = folder.get_message(object_id=item.get("id"))
                except Exception:
                    msg = None

                if not msg:
                    continue

                # ==== BEGIN: your existing per-item logic (unchanged) ====
                conv_id = getattr(msg, "conversation_id", None)
                if not conv_id:
                    try:
                        msg.refresh()
                    except Exception:
                        pass
                    if any((c or "").startswith("PAIRActioned") for c in (msg.categories or [])):
                        continue
                    if not msg.is_read:
                        sender_addr = (msg.sender.address or "").lower() if msg.sender else ""
                        sender_is_staff = sender_addr in [e.lower() for e in EMAILS_TO_FORWARD]
                        is_automated_reply = bool(re.search(fr"{REPLY_ID_TAG}\s*([^\s<]+)", msg.body or "", flags=re.I|re.S))
                        if sender_is_staff and is_automated_reply and not any((c or "").startswith("PAIRActioned") for c in (msg.categories or [])):
                            handle_internal_reply(msg)
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
                                    "sender": msg.sender.address if msg.sender else "",
                                    "received": msg.received.strftime('%Y-%m-%d %H:%M'),
                                    "message_obj": msg,
                                }
                                print(f"📧 Email stored for approval: {msg.subject}")
                                # Tag immediately when enqueued for approval
                                tag_email(msg, ["pending_review"], replyTag=False)
                            else:
                                print(f"⏭️  Email already in pending queue, skipping: {msg.subject}")
                        else:
                            print(json.dumps(result, indent=2))
                            handle_new_email(msg, result)
                    continue  # done with this delta item

                # Normal path: fetch unread messages in this conversation
                try:
                    unread_msgs = unread_in_conversation(folder, mailbox, conv_id)
                except Exception as e:
                    print(f"   Could not fetch unread children for {conv_id}: {e}")
                    continue

                if not unread_msgs:
                    continue

                for child in unread_msgs:
                    try:
                        child.refresh()
                    except Exception:
                        pass
                    if any((c or "").startswith("PAIRActioned") for c in (child.categories or [])):
                        continue

                    sender_addr = (child.sender.address or "").lower() if child.sender else ""
                    sender_is_staff = sender_addr in [e.lower() for e in EMAILS_TO_FORWARD]
                    is_automated_reply = bool(re.search(fr"{REPLY_ID_TAG}\s*([^\s<]+)", child.body or "", flags=re.I|re.S))
                    if sender_is_staff and is_automated_reply and not any((c or "").startswith("PAIRActioned") for c in (child.categories or [])):
                        handle_internal_reply(child)
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
                                "sender": child.sender.address if child.sender else "",
                                "received": child.received.strftime('%Y-%m-%d %H:%M'),
                                "message_obj": child,
                            }
                            print(f"📧 Email stored for approval: {child.subject}")
                            # Tag immediately when enqueued for approval
                            tag_email(child, ["pending_review"], replyTag=False)
                        else:
                            print(f"⏭️  Email already in pending queue, skipping: {child.subject}")
                    else:
                        handle_new_email(child, result)
                # ==== END: your existing per-item logic ====

            # Prepare next page or finish
            params = None  # after first request, Graph returns absolute nextLink/deltaLink
            next_url = data.get("@odata.nextLink")
            if data.get("@odata.deltaLink"):
                final_delta_link = data["@odata.deltaLink"]

        # --- DIAGNOSTIC SUMMARY ---
        print(f"DIAGNOSTIC [{name}]: Processed {total_changed} changed item(s). Final delta: {bool(final_delta_link)}")

        # Return the FULL delta URL to persist (use as-is next time)
        return final_delta_link or delta_token_url

    except Exception as e:
        print(f" Error accessing {name}: {e}")
        return delta_token_url or final_delta_link


def handle_new_email(msg, result):
    """
    Takes a message and its AI classification result, then acts on it.
    It does NOT call the AI again.
    """
    clear_pending_tag(msg)
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

def clear_pending_tag(msg): 
    try: 
        cats = set(msg.categories or [])
        msg.categories = sorted(cats)
        msg.save_message()
        print("removed the tag")
    except Exception as e: 
        print("could not remove the tag'")

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



inbox_delta, junk_delta = load_last_delta()

# Import and initialize the web interface
from web_interface import app, create_email_routes, start_web_server

# Initialize the email routes with the required dependencies
create_email_routes(pending_emails, handle_new_email, mailbox, mark_as_read)

# Start the web server
start_web_server()
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
            account = Account(credentials, auth_flow_type="authorization", token_backend=token_backend, tenant_id=TENANT_ID)
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
        print(f"🔄 Checking for new emails... (Pending: {len(pending_emails)})")
        
        # Check if it's been too long since last successful check (1 hour)
        if time.time() - last_successful_check > 3600:
            print("⚠️ No successful checks in 1 hour, forcing re-authentication...")
            reconnect_account()
        
        inbox_delta = process_folder(inbox_folder, "INBOX", inbox_delta)
        junk_delta = process_folder(junk_folder, "JUNK", junk_delta)
        print(f"DEBUG: Current directory is: {os.getcwd()}")
        #gets the new delta tokens and then saves them,
        save_last_delta(inbox_delta, junk_delta)
        
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
