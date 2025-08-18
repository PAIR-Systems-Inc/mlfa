from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta
import os, time, openai, json
import textwrap
import re
from bs4 import BeautifulSoup
from flask import Flask, jsonify, send_from_directory, request
from flask_cors import CORS
import threading


load_dotenv()

### CONSTANTS

START_TIME = datetime.now(timezone.utc) - timedelta(weeks=2)
processed_messages = set()

CLIENT_ID = "b985204d-8506-4bb3-8f54-25899e38c825"
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
TENANT_ID = os.getenv("O365_TENANT_ID")
REPLY_ID_TAG = "Pair_Reply_Reference_ID"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EMAIL_TO_WATCH = os.getenv("EMAIL_TO_WATCH")


EMAILS_TO_FORWARD = ['Mujahid.rasul@mlfa.org', 'Syeda.sadiqa@mlfa.org', 'Arshia.ali.khan@mlfa.org', 'Maria.laura@mlfa.org', 'info@mlfa.org', 'aisha.ukiu@mlfa.org', 'shawn@strategichradvisory.com', 'm.ahmad0826@gmail.com']
NONREAD_CATEGORIES = {"marketing"}  # Keep these unread
SKIP_CATEGORIES = {'spam', 'cold_outreach', 'newsletter', 'irrelevant_other'}

HUMAN_CHECK = True  # Enable human check for approval hub

# Simple storage for current email data
current_email_data = {
    "subject": "",
    "body": "",
    "classification": {},
    "sender": "",
    "received": "",
    "message_obj": None  # Store the actual message object for processing
}

# Flask app for approval hub
app = Flask(__name__, static_folder='.')
CORS(app)




###CONNECTING


openai.api_key = OPENAI_API_KEY
credentials = (CLIENT_ID, None) #delete
#credentials = (CLIENT_ID, CLIENT_SECRET)
token_backend = FileSystemTokenBackend(token_path=".", token_filename="o365_token.txt")
account = Account(credentials, auth_flow_type="authorization", token_backend=token_backend)  #Delete this
#account = Account(credentials, auth_flow_type="credentials",  tenant_id=TENANT_ID) 

if not account.is_authenticated:
    account.authenticate(scopes=['basic', 'message_all']) #Deleting this
    #account.authenticate()

mailbox = account.mailbox(resource=EMAIL_TO_WATCH)
inbox_folder = mailbox.inbox_folder()
junk_folder = mailbox.junk_folder()




def read_token(path):
    if os.path.exists(path):
        with open(path, "r") as f:
            return f.read().strip()
    return None

def load_last_delta():
    inbox_token = read_token("delta_token_inbox.txt")
    junk_token = read_token("delta_token_junk.txt")
    return inbox_token, junk_token


def save_last_delta(inbox_token, junk_token):
    if inbox_token: open("delta_token_inbox.txt", "w").write(inbox_token)
    if junk_token: open("delta_token_junk.txt", "w").write(junk_token)


#Passes the subject and body of the email to chat gpt, which figures out how to handle the email. 
#Chat-GPT nicely returns the information in json format. 

#Args:
#subject (str): The subject of the email. 
#body (str): The body of the email. 

#Returns:
#A json script that includes the category of the email, who it needs to be forwarded to, and why. 

def classify_email(subject, body):
    prompt = f"""
You are an email routing assistant for MLFA (Muslim Legal Fund of America), a nonprofit organization focused on legal advocacy for Muslims in the United States.

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

- **Organizational questions** → If the sender is asking about **MLFA’s internal operations**, such as leadership, partnerships, volunteering, employment, or collaboration, categorize as `"organizational"`. Forward to:
Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org

- **Volunteer inquiries** → If someone is **offering to volunteer** their time or skills to MLFA **or** is **asking about volunteering** (for themselves or on behalf of someone else), categorize as `"volunteer"`. Forward to:
aisha.ukiu@mlfa.org

- **Job applications** → If someone is **applying for a paid job**, sending a resume, or asking about open employment positions, categorize as `"job_application"`. Forward to:
shawn@strategichradvisory.com

- **Internship applications** → If someone is **applying for an internship** (paid or unpaid), sending a resume for an internship program, or inquiring about internship opportunities, categorize as `"internship"`. Forward to:
aisha.ukiu@mlfa.org

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
7. Use `all_recipients` only for forwarded categories: `"donor"`, `"sponsorship"`, `"fellowship"`, `"organizational"`, `"volunteer"`, `"job_application"`, `"internship"`.
8. For `"legal"`, `"marketing"`, and all `"irrelevant"` types, leave `all_recipients` empty.

PRIORITY & TIES:
- If `"legal"` applies, **still include all other relevant categories** — `"legal"` is additive, never exclusive.
- `"marketing"` vs `"cold_outreach"`: choose only one based on tailoring (see rules above).

Return a JSON object with:
- `categories`: array from ["legal","donor","sponsorship","fellowship","organizational","volunteer","job_application","internship","marketing","spam","cold_outreach","newsletter","irrelevant_other"]
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
        response = openai.chat.completions.create(
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
    qs = qs.select([
        'id', 'conversationId', 'isRead', 'receivedDateTime', 'from', 'sender', 'subject'
    ])

    try:
        msgs = folder.get_messages(query=qs)

        for msg in msgs:
            # For each conversation that changed, act ONLY on unread children
            conv_id = getattr(msg, 'conversation_id', None)
            if not conv_id:
                # Fallback: skip if no conversation id (rare)
                dedup_key = getattr(msg, 'internet_message_id', None) or msg.object_id
                if dedup_key in processed_messages:
                    continue
                try:
                    msg.refresh()
                except Exception:
                    pass
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
                    if any((c or '').startswith('PAIRActioned') for c in (msg.categories or [])):
                        processed_messages.add(dedup_key)
                        continue

                    print(f"\nNEW:  [{name}] {msg.received.strftime('%Y-%m-%d %H:%M')} | "
                          f"{msg.sender.address if msg.sender else 'UNKNOWN'} | {msg.subject}")
                    result = classify_email(msg.subject, body_to_analyze)
                    if HUMAN_CHECK: 
                        print(json.dumps(result, indent=2))
                        # Store current email data
                        current_email_data["subject"] = msg.subject
                        current_email_data["body"] = body_to_analyze
                        current_email_data["classification"] = result
                        current_email_data["sender"] = msg.sender.address
                        current_email_data["received"] = msg.received.strftime('%Y-%m-%d %H:%M')
                        current_email_data["message_obj"] = msg
                        print(f"📧 Email stored for approval: {msg.subject}")
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
                    continue

                # Make sure we have up-to-date fields on the child
                try:
                    child.refresh()
                except Exception:
                    pass

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

                # 2) Skip if already actioned
                if any((c or '').startswith('PAIRActioned') for c in (child.categories or [])):
                    processed_messages.add(dedup_key)
                    continue

                # 3) Classify using reply-only text, then handle
                body_to_analyze = get_clean_message_text(child)
                print(f"\nNEW:  [{name}] {child.received.strftime('%Y-%m-%d %H:%M')} | "
                      f"{child.sender.address if child.sender else 'UNKNOWN'} | {child.subject}")
                result = classify_email(child.subject, body_to_analyze)
                print(json.dumps(result, indent=2))
                
                if HUMAN_CHECK:
                    # Store current email data
                    current_email_data["subject"] = child.subject
                    current_email_data["body"] = body_to_analyze
                    current_email_data["classification"] = result
                    current_email_data["sender"] = child.sender.address
                    current_email_data["received"] = child.received.strftime('%Y-%m-%d %H:%M')
                    current_email_data["message_obj"] = child
                    print(f"📧 Email stored for approval: {child.subject}")
                else:
                    handle_new_email(child, result)

                # 4) Dedup remember
                processed_messages.add(dedup_key)

        # Return latest delta token (if present) to persist
        return getattr(msgs, 'delta_token', delta_token)

    except Exception as e:
        print(f" Error accessing {name}: {e}")
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
        # fwd.to.add(list(recipients_set))
        fwd.to.add('m.ahmad0826@gmail.com')  # For testing
        # Add th=e hidden tracking ID into the top of the forwarded body
        instruction_html = f"""<div style="display:none;">{REPLY_ID_TAG}{msg.object_id}</div>"""
        
        # Prepend to the auto-generated forward body
        fwd.body = "Please press 'Reply All,' and reply to info@mlfa.org. You're email will automatically be sent to the correct person. " + instruction_html
        fwd.body_type = 'HTML'
        
        # Send the forward
        fwd.send()


    if not set(categories).issubset(NONREAD_CATEGORIES):
        mark_as_read(msg)


def handle_emails(categories, result, recipients_set, msg, name_sender): 
    for category in categories:
        if category == "legal":
            reply_message = msg.reply(to_all=False)
            # Check if this email needs a personal reply based on classification
            needs_personal = result.get("needs_personal_reply", False)

            if needs_personal:
                reply_message.body = f"""
                    <p>Dear {name_sender},</p>

                    <p>Thank you for reaching out to the Muslim Legal Fund of America (MLFA). 
                    We deeply appreciate you taking the time to contact us and share your situation.</p>

                    <p>We understand that seeking legal assistance can be a challenging and emotional process, 
                    and we want you to know that we take every inquiry seriously. Your trust in MLFA to 
                    potentially help with your legal matter means a great deal to us.</p>

                    <p>If you have not already done so, please submit a formal application for legal assistance through our website: <br>
                    <a href="https://mlfa.org/application-for-legal-assistance/">https://mlfa.org/application-for-legal-assistance/</a></p>

                    <p>Our team reviews each application carefully, and we will be in touch with you regarding 
                    next steps. If you have any questions about the application process or need help 
                    completing it, please don't hesitate to reach out.</p>

                    <p>We appreciate your patience as we work through applications, and we look forward to 
                    learning more about how we might be able to help.</p>

                    <p>Warm regards,<br>
                    The MLFA Team<br>
                    Muslim Legal Fund of America</p>
                """
                reply_message.body_type = "HTML"
            else:
                reply_message.body = f"""
                    <p>Dear {name_sender},</p>

                    <p>Thank you for contacting the Muslim Legal Fund of America (MLFA).</p>

                   <p> If you have not already done so, please submit a formal application for legal assistance through our website:  
                    https://mlfa.org/application-for-legal-assistance/</p>

                    <p>This ensures our legal team has the information needed to review your case promptly.</p>

                    <p>Sincerely, <br> 
                    The MLFA Team</p>

                """
                reply_message.body_type = "HTML"
            reply_message.send()

        elif category == "donor":
            recipients_set.update([f"{EMAILS_TO_FORWARD[0]}", f"{EMAILS_TO_FORWARD[1]}"])

        elif category == "sponsorship":
            recipients_set.update([f"{EMAILS_TO_FORWARD[2]}", f"{EMAILS_TO_FORWARD[3]}"])

        elif category == "organizational":
            recipients_set.update([f"{EMAILS_TO_FORWARD[2]}", f"{EMAILS_TO_FORWARD[3]}"])

        elif category == "volunteer":
            recipients_set.update([f"{EMAILS_TO_FORWARD[5]}"])

        elif category == "internship":
            recipients_set.update([f"{EMAILS_TO_FORWARD[5]}"])

        elif category == "job_application":
            recipients_set.update([f"{EMAILS_TO_FORWARD[6]}"])

        elif category == "fellowship":
            recipients_set.update([f"{EMAILS_TO_FORWARD[5]}"])

        elif category == "marketing":
            inbox = mailbox.inbox_folder()
            sales_folder = inbox.get_folder(folder_name="Sales emails")
            print("Moving to sales emails folder.")
            msg.move(sales_folder)

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
        final_reply = original_msg.reply(to_all=False)
        final_reply.body = reply_content
        final_reply.body_type = "HTML"
        final_reply.send()
        print(f"   Sent reply to original sender: {original_msg.sender.address}")
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



# Flask routes for approval hub
@app.route('/')
def index():
    return send_from_directory('.', 'approval-hub.html')

@app.route('/api/emails')
def get_emails():
    if current_email_data["subject"]:
        # Format data for the interface
        categories = current_email_data["classification"].get('categories', [])
        category_display = ', '.join([cat.replace('_', ' ').title() for cat in categories])
        
        recipients = current_email_data["classification"].get('all_recipients', [])
        recipients_display = ', '.join(recipients) if recipients else 'None'
        
        reasons = current_email_data["classification"].get('reason', {})
        reason_display = '; '.join(reasons.values()) if reasons else 'No reason provided'
        
        email = {
            "id": "current",
            "meta": f"FROM: [INBOX] {current_email_data['received']} | {current_email_data['sender']} | {current_email_data['subject']}",
            "senderName": current_email_data["classification"].get('name_sender', 'Unknown'),
            "category": category_display,
            "recipients": recipients_display,
            "needsReply": "Yes" if current_email_data["classification"].get('needs_personal_reply', False) else "No",
            "reason": reason_display,
            "escalation": current_email_data["classification"].get('escalation_reason') or 'None',
            "originalContent": current_email_data["body"],
            "status": "pending"
        }
        return jsonify([email])
    else:
        return jsonify([])

@app.route('/api/emails/<email_id>/approve', methods=['POST'])
def approve_email(email_id):
    if current_email_data["subject"] and current_email_data["message_obj"]:
        msg = current_email_data["message_obj"]
        classification = current_email_data["classification"]
        
        print(f"✅ Email approved: {current_email_data['subject']} - Processing normally")
        
        # Process the email normally using the stored message and classification
        handle_new_email(msg, classification)
        
    # Clear the current email
    current_email_data.clear()
    current_email_data.update({"subject": "", "body": "", "classification": {}, "sender": "", "received": "", "message_obj": None})
    return jsonify({"status": "success", "message": "Email approved and will be processed normally"})

@app.route('/api/emails/<email_id>/reject', methods=['POST'])
def reject_email(email_id):
    data = request.get_json()
    reason = data.get('reason', 'No reason provided')
    
    if current_email_data["subject"] and current_email_data["message_obj"]:
        msg = current_email_data["message_obj"]
        print(f"❌ Email rejected: {current_email_data['subject']} - Reason: {reason}")
        
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
                print(f"✅ Marked email as processed")
            except Exception as e:
                print(f"⚠️ Could not mark as processed: {e}")
            
        except Exception as e:
            print(f"❌ Error handling rejected email: {e}")
    
    # Clear the current email
    current_email_data.clear()
    current_email_data.update({"subject": "", "body": "", "classification": {}, "sender": "", "received": "", "message_obj": None})
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

inbox_delta, junk_delta = load_last_delta()
print(f"Monitoring inbox + junk for: {EMAIL_TO_WATCH} … Ctrl-C to stop.")
print(f"📧 Approval hub available at: http://localhost:5000")

while True:
    inbox_delta = process_folder(inbox_folder, "INBOX", inbox_delta)
    junk_delta = process_folder(junk_folder, "JUNK", junk_delta)
    #gets the new delta tokens and then saves them,
    save_last_delta(inbox_delta, junk_delta)
    time.sleep(10)
