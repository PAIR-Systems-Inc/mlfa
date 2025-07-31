from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone
import os, time, openai, json


load_dotenv()

### CONSTANTS

START_TIME = datetime.now(timezone.utc)
processed_messages = set()

CLIENT_ID = "b985204d-8506-4bb3-8f54-25899e38c825"
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
TENANT_ID = os.getenv("O365_TENANT_ID")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EMAIL_TO_WATCH = os.getenv("EMAIL_TO_WATCH")


EMAILS_TO_FORWARD = ['Mujahid.rasul@mlfa.org', 'Syeda.sadiqa@mlfa.org', 'Arshia.ali.khan@mlfa.org', 'Maria.laura@mlfa.org', 'info@mlfa.org']
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

    Your job is to classify incoming emails based on their **content, intent, and relevance** to MLFA’s mission. Do not rely on keywords alone. Use the rules below to assign one or more categories and determine the correct recipients if applicable.

    ROUTING RULES & RECIPIENTS:

    - **Legal inquiries** → If someone is explicitly **asking for legal help or representation**, categorize as `"legal"`. These users should be referred to MLFA’s "Apply for Help" form (no forwarding needed).

    - **Donor-related inquiries** → Categorize as `"donor"` only if the **sender is a donor** or is asking about a **specific donation**, such as issues with payment, receipts, or follow-ups. Forward to:
    Mujahid.rasul@mlfa.org, Syeda.sadiqa@mlfa.org

    - **Sponsorship requests** → If someone is **requesting sponsorship or financial support from MLFA**, categorize as `"sponsorship"`. Forward to:
    Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org

    - **Organizational questions** → If the sender is asking about **MLFA’s internal operations**, such as leadership, partnerships, volunteering, employment, or outreach, categorize as `"organizational"`. Forward to:
    Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org

    - **Email marketing/sales** → If the sender is **promoting or offering a product, service, or software**, categorize as `"marketing"` only if:
    1. The offering is clearly and specifically relevant to **MLFA’s nonprofit or legal work**, and
    2. The sender shows contextual awareness of MLFA’s mission or prior contact,
    3. The product is narrowly targeted (e.g., legal case management, Islamic nonprofit compliance tools, or court filing automation).
    These messages should be placed in the "Sales emails" folder. **Do not treat generic or cold outreach as marketing.**

    - **Spam** → Obvious scams, phishing attempts, AI-generated nonsense, or any fraudulent messages. These should be moved to the Junk folder.

    - **Cold outreach** → Generic, unsolicited B2B or sales emails, even if they mention fundraising, AI, or nonprofit topics. Typical signs include phrases like “limited-time offer,” “800% increase,” “click here,” or “boost donations.” These should be marked as read and ignored.

    - **Newsletter** → Mass email updates, PR announcements, or content digests unrelated to specific communication with MLFA. These may be routed to a Newsletters folder if available.

    - **Irrelevant (other)** → Anything not aligned with MLFA’s mission and not covered by spam, cold outreach, or newsletter — such as misdirected inquiries or off-topic messages. These should be marked as read and left unforwarded.

    IMPORTANT GUIDELINES:

    1. Focus on the **sender’s intent and relevance to MLFA’s legal mission**, not just keywords.
    2. If someone is **offering legal services or collaboration**, categorize as `"organizational"`, not `"legal"`.
    3. If someone is **selling something**, even if it mentions donors or legal tech, treat it as `"marketing"` only if it is **narrowly tailored to MLFA**. Otherwise, use a subcategory of `"irrelevant"`.
    4. **Cold emails, newsletters, SEO tools, AI apps, or sales automation platforms are not marketing** — classify them as `"cold_outreach"`, `"newsletter"`, or `"spam"` depending on content.
    5. An email can have **multiple categories** only if the sender’s role clearly spans them (e.g., a donor requesting sponsorship).
    6. For forwarding categories (`donor`, `sponsorship`, `organizational`), include all relevant email addresses in `all_recipients`.
    7. For all non-forwarded categories (`legal`, `marketing`, and the irrelevant subtypes), leave `all_recipients` empty.

    Return a JSON object with the following keys:

    - `categories`: array of applicable categories from:
    ["legal", "donor", "sponsorship", "organizational", "marketing", "spam", "cold_outreach", "newsletter", "irrelevant_other"]

    - `all_recipients`: list of relevant MLFA email addresses (may be empty)

    - `reason`: a dictionary explaining why each category was assigned

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
    if delta_token: #to start from the appropriate place. 
        qs = qs.delta_token(delta_token)
    try:
        msgs = folder.get_messages(query=qs)
        # Keeps only the emails that either have been unread AND a delta token was used or it was recieved after the start_time. 
        filtered = [
            msg for msg in msgs
            if not msg.is_read and (delta_token or (msg.received and msg.received > START_TIME))
        ]
        if not filtered:
            return getattr(msgs, 'delta_token', delta_token)

        for msg in filtered:
            msg_id = getattr(msg, 'object_id', getattr(msg, 'message_id', str(hash(msg.subject + str(msg.received)))))
            if msg_id in processed_messages:
                continue

            print(f"\nNEW:  [{name}] {msg.received.strftime('%Y-%m-%d %H:%M')} | {msg.sender.address} | {msg.subject}")

            result = classify_email(msg.subject, msg.body)
            print(json.dumps(result, indent=2))

            processed_messages.add(msg_id)

            categories = result.get("categories", [])
            recipients_set = set()

            for category in categories:
                if category == "legal":
                    reply_message = msg.reply(to_all=False)
                    reply_message.body = (
                        "Thank you for reaching out to the Muslim Legal Fund of America.\n\n"
                        "If you are seeking legal assistance, please submit an application through our website at:\n"
                        "https://mlfa.org/application-for-legal-assistance/\n\n"
                        ". We appreciate your message.\n\n"
                    )

                    reply_message.send()

                elif category == "donor":
                    recipients_set.update([f"{EMAILS_TO_FORWARD[0]}", f"{EMAILS_TO_FORWARD[1]}"])

                elif category == "sponsorship":
                    recipients_set.update([f"{EMAILS_TO_FORWARD[2]}", f"{EMAILS_TO_FORWARD[3]}"])

                elif category == "organizational":
                    recipients_set.update([f"{EMAILS_TO_FORWARD[2]}", f"{EMAILS_TO_FORWARD[3]}"])

                elif category == "marketing":
                    inbox = mailbox.inbox_folder()
                    sales_folder = inbox.get_folder(folder_name="Sales emails")
                    print("Moving to sales emails folder.")
                    msg.move(sales_folder)
                    
            tag_email(msg, categories)

            if recipients_set:
                forward_msg = msg.forward()
                #forward_msg.to.add(list(recipients_set))
                forward_msg.to.add('m.ahmad0826@gmail.com')
                forward_msg.send()
            
            if any(cat in categories for cat in SKIP_CATEGORIES):
                continue

            if not set(categories).issubset(NONREAD_CATEGORIES):
                mark_as_read(msg)
   
        return getattr(msgs, 'delta_token', delta_token)

    except Exception as e:
        print(f" Error accessing {name}: {e}")
        return delta_token

def tag_email(msg, categories):
    prefixed = []
    for c in categories:
        if c in ['spam', 'cold_outreach', 'newsletter']:
            prefixed.append(f'PAIRActioned/irrelevant/{c}')
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




inbox_delta, junk_delta = load_last_delta()
print(f"Monitoring inbox + junk for: {EMAIL_TO_WATCH} … Ctrl-C to stop.")

while True:
    inbox_delta = process_folder(inbox_folder, "INBOX", inbox_delta)
    junk_delta = process_folder(junk_folder, "JUNK", junk_delta)
    #gets the new delta tokens and then saves them,
    save_last_delta(inbox_delta, junk_delta)
    time.sleep(10)