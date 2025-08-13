from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta
import os, time, openai, json
import textwrap
import re
from bs4 import BeautifulSoup

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
If none of the above, set `needs_personal_reply=false`.

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

- **Volunteer inquiries** → If someone is **offering to volunteer** their time or skills to MLFA, categorize as `"volunteer"`. Forward to:
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
- If `"legal"` applies, include it regardless of other categories.
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
    qs = folder.new_query()
    if delta_token:
        qs = qs.delta_token(delta_token)
    
    try:
        msgs = folder.get_messages(query=qs)
        
        for msg in msgs:
            msg_id = msg.object_id
            if msg_id in processed_messages:
                continue
            processed_messages.add(msg_id)

            is_actioned = any(cat.startswith("PAIRActioned") for cat in (msg.categories or [])) #has it been handled already
            sender_is_staff = msg.sender.address.lower() in [email.lower() for email in EMAILS_TO_FORWARD] #are they a possible replier
            is_automated_reply = REPLY_ID_TAG in msg.body #is this a reply to the forwarded email

            if sender_is_staff and is_automated_reply and not is_actioned:
                handle_internal_reply(msg) #if all three are true, then it must be a reply to a forwarded email. 
                continue 

            if not msg.is_read and not is_actioned:
                print(f"\nNEW:  [{name}] {msg.received.strftime('%Y-%m-%d %H:%M')} | {msg.sender.address} | {msg.subject}")
                
                result = classify_email(msg.subject, msg.body)
                print(json.dumps(result, indent=2))

                categories = result.get("categories", [])
                recipients_set = set(result.get("all_recipients", []))
                name_sender = result.get("name_sender")

                tag_email(msg, categories, replyTag=False)
                handle_emails(categories, result, recipients_set, msg, name_sender)

                if recipients_set:
                    original_msg_id = msg.object_id
                    forward_msg = msg.forward()

                    forward_msg.subject = f"FW: {msg.subject}"

                    # 1. Manually build the instruction block (no visual styling)
                    instruction_html = f"""
                        <div>
                            <span style="display:none;">{REPLY_ID_TAG}{original_msg_id}</span>
                        </div>
                        """
                    forward_msg.body =  "Please reply to all recipients in the message below, ensuring that you include info@mlfa.org. Replies sent to info@mlfa.org will be automatically delivered. "+ instruction_html
                    forward_msg.body_type = 'HTML'
                    
                    # Send the email to the correct recipients
                    forward_msg.to.add('m.ahmad0826@gmail.com') # For testing
                    # forward_msg.to.add(list(recipients_set))
                    forward_msg.send()

                if not set(categories).issubset(NONREAD_CATEGORIES):
                    mark_as_read(msg)
        
        return getattr(msgs, 'delta_token', delta_token)

    except Exception as e:
        print(f" Error accessing {name}: {e}")
        return delta_token
   

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

                    <p>If you haven't already, please submit a formal application through our website:<br>
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
                reply_message.body = """
                    <p>Thank you for reaching out to the Muslim Legal Fund of America.</p>

                    <p>If you are seeking legal assistance, please submit an application through our website:<br>
                    <a href="https://mlfa.org/application-for-legal-assistance/">https://mlfa.org/application-for-legal-assistance/</a></p>

                    <p>We appreciate your message.</p>
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
    prefixed = []
    for c in categories:
        if c in ['spam', 'cold_outreach', 'newsletter']:
            prefixed.append(f'PAIRActioned/irrelevant/{c}')
        elif replyTag == True: 
            prefixed.append(f'PAIRActioned/replied/{c}')
        else:
            prefixed.append(f'PAIRActioned/{c}')
    all_tags = prefixed + ['PAIRActioned']
    msg.categories = all_tags
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

    # 5. Tag both emails to prevent loops
    replier_email = msg.sender.address
    tag_email(original_msg, [replier_email], replyTag=True)
    msg.mark_as_read()
    tag_email(msg, [replier_email], replyTag=True)
    print("   Cleanup complete. Reply process finished.")




inbox_delta, junk_delta = load_last_delta()
print(f"Monitoring inbox + junk for: {EMAIL_TO_WATCH} … Ctrl-C to stop.")

while True:
    inbox_delta = process_folder(inbox_folder, "INBOX", inbox_delta)
    junk_delta = process_folder(junk_folder, "JUNK", junk_delta)
    #gets the new delta tokens and then saves them,
    save_last_delta(inbox_delta, junk_delta)
    time.sleep(10)
