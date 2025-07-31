from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone
import os, time, openai, json


load_dotenv()
START_TIME = datetime.now(timezone.utc)
processed_messages = set()

CLIENT_ID = "b985204d-8506-4bb3-8f54-25899e38c825"
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
TENANT_ID = os.getenv("O365_TENANT_ID")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EMAIL_TO_WATCH = "mariamahmadpear@outlook.com"

openai.api_key = OPENAI_API_KEY
credentials = (CLIENT_ID, None)
token_backend = FileSystemTokenBackend(token_path=".", token_filename="o365_token.txt")
account = Account(credentials, auth_flow_type="authorization", token_backend=token_backend)

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
    inbox_token = read_token("delta_token_inbox.txt")
    junk_token = read_token("delta_token_junk.txt")
    return inbox_token, junk_token


def save_last_delta(inbox_token, junk_token):
    if inbox_token: open("delta_token_inbox.txt", "w").write(inbox_token)
    if junk_token: open("delta_token_junk.txt", "w").write(junk_token)


def classify_email(subject, body):
    """
    Passes the subject and body of the email to chat gpt, which figures out how to handle the email. 
    Chat-GPT nicely returns the information in json format. 

    Args:
    subject (int): The ID of the user to fetch.
    body (bool): Whether to include historical data. Defaults to False.

    Returns:
    dict: A dictionary representing the user and (optionally) their history.

    Raises:
    ValueError: If the user_id is invalid.
    ConnectionError: If the database cannot be reached.
    """

    prompt = f"""
    You are an email routing assistant for MLFA (Muslim Legal Fund of America), a nonprofit organization.

    Your job is to classify incoming emails based on their content and intent, not just keywords. Use the routing rules below to assign one or more categories to each email, along with the correct recipient(s) if applicable.

    Routing Rules & Recipients:

    - Legal inquiries → If someone is **asking for legal help or representation**, categorize as "legal". These users should be directed to the "Apply for Help" website.
    - Donor-related inquiries → Only categorize as "donor" if the **sender is a donor** or is **asking about a specific donation**, such as payment issues, receipts, or donation follow-ups. Forward to: Mujahid.rasul@mlfa.org, Syeda.sadiqa@mlfa.org
    - Sponsorship requests → If someone is **requesting sponsorship or support from MLFA**, categorize as "sponsorship". Forward to: Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org
    - Organizational questions → If someone is **asking about MLFA’s internal operations, leadership, volunteering, or employment**, categorize as "organizational". Forward to: Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org
    - Email marketing/sales → If the sender is **offering or promoting a product, service, or software**, categorize as "marketing" **only if the offering is clearly relevant to nonprofit operations**, such as donor management tools, volunteer coordination, legal intake software, or nonprofit fundraising platforms. These should be moved to the "Sales emails" folder.
    - Irrelevant/spam → Categorize as "irrelevant" if the email is:
    - A **generic B2B or commercial pitch** not tailored to nonprofit or legal aid work (e.g., SEO services, AI content tools, sales automation)
    - An **obvious scam, AI-generated nonsense, or phishing**
    - **Unrelated to MLFA’s mission** of legal advocacy for Muslims in the U.S.

    IMPORTANT GUIDELINES:

    1. Focus on the sender’s **intent and relevance to MLFA**, not just topic words.
    2. If someone is offering legal services or partnership, categorize as **organizational**, not legal.
    3. If someone is **selling something**, even if related to donors or legal tech, it is still "marketing" (not "donor" or "legal").
    4. An email can belong to **multiple categories**, but only if the sender is acting in those roles (e.g., a donor or applicant).
    5. For forwarding categories (donor, sponsorship, organizational), include all relevant recipients in `all_recipients`.
    6. If the email is marketing-only or irrelevant, leave `all_recipients` empty.
    7. If an email would otherwise be marketing, but is **irrelevant to nonprofit legal advocacy**, categorize it as `"irrelevant"` instead.

    Return a JSON object with:

    - `categories`: array of applicable categories from ["legal", "donor", "sponsorship", "organizational", "marketing", "irrelevant"]
    - `all_recipients`: list of relevant email addresses (may be empty for legal, marketing, or irrelevant cases)
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
                    recipients_set.update(["Mujahid.rasul@mlfa.org", "Syeda.sadiqa@mlfa.org"])

                elif category == "sponsorship":
                    recipients_set.update(["Arshia.ali.khan@mlfa.org", "Maria.laura@mlfa.org"])

                elif category == "organizational":
                    recipients_set.update(["Arshia.ali.khan@mlfa.org", "Maria.laura@mlfa.org"])

                elif category == "marketing":
                    inbox = mailbox.inbox_folder()
                    sales_folder = inbox.get_folder(folder_name="Sales emails")
                    print("Moving to sales emails folder.")
                    msg.move(sales_folder)
                    


            if recipients_set:
                forward_msg = msg.forward()
                #forward_msg.to.add(list(recipients_set))
                forward_msg.to.add('m.ahmad0826@gmail.com')
                forward_msg.send()
                
            if "irrelevant" in categories:
                tag_email(msg, categories)
                continue

            
            
            if set(categories) != {"marketing"}:
                tag_email(msg, categories)
                mark_as_read(msg)

            
        return getattr(msgs, 'delta_token', delta_token)

    except Exception as e:
        print(f" Error accessing {name}: {e}")
        return delta_token



def tag_email(msg, categories): 
    prefixed = [f'PAIRActioned/{c}' for c in categories]
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