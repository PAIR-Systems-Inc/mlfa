from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import datetime, timezone
import os, time, openai, json


load_dotenv()
START_TIME = datetime.now(timezone.utc)
processed_messages = set()

CLIENT_ID = "b985204d-8506-4bb3-8f54-25899e38c825"
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
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

# ──────────────── DELTA TOKEN ────────────────
def load_last_delta():
    def read_token(path): return open(path).read().strip() if os.path.exists(path) else None
    return read_token("delta_token_inbox.txt"), read_token("delta_token_junk.txt")

def save_last_delta(inbox_token, junk_token):
    if inbox_token: open("delta_token_inbox.txt", "w").write(inbox_token)
    if junk_token: open("delta_token_junk.txt", "w").write(junk_token)

# ──────────────── CLASSIFICATION ────────────────
def classify_email(subject, body):
    prompt = f"""
You are an email routing assistant for MLFA (Muslim Legal Fund of America), a nonprofit organization.

Routing Rules & Recipients:
- Legal inquiries (asking for help) → Direct to "Apply for Help" website 
- Donor-related inquiries (payments, receipts) → Forward to: Mujahid.rasul@mlfa.org, Syeda.sadiqa@mlfa.org
- Sponsorship requests → Forward to: Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org  
- Organizational questions → Forward to: Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org
- Email marketing/sales → Move to "Sales emails" folder

IMPORTANT: 
1. An email can have MULTIPLE categories
2. If "legal" means someone ASKING FOR HELP → use "website_redirect" 
3. If "legal" means someone OFFERING TO HELP (volunteer, attorney offering services) → use "forward" to organizational contacts
4. Always include ALL recipients for forwarding categories

Analyze this email and return JSON with:
- categories: array of applicable categories from ["legal", "donor", "sponsorship", "organizational", "marketing"]
- primary_action: "forward" (if any forwarding category), "website_redirect" (if asking for legal help), or "folder_sales"
- all_recipients: combined list of ALL email addresses from applicable forwarding categories (empty only for website_redirect or folder_sales)
- reason: explanation of each category

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
        if raw.startswith("```json"): raw = raw[len("```json"):].strip()
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

            #handling the emails appropriately .
            if result.get('primary_action') == 'forward':
                try:
                    recipients = ['m.ahmad0826@gmail.com']
                    
                    print(f"    Forwarding email to: {', '.join(recipients)}")
                    forward_msg = msg.forward()
                    forward_msg.to.add(recipients)
                    forward_msg.send()
                    print("   Forwarded successfully.")
                except Exception as e:
                    print(f"   Could not forward email: {e}")
            elif result.get('primary_action') == "website_redirect": 
                reply_message = msg.reply(to_all=False) # if we want reply-all then it has to be True not False. 
                reply_message.body = "Thank you for reaching out. Please go ahead and Apply for Help through our website: https://mlfa.org/application-for-legal-assistance/   -AI"
                reply_message.send()
            elif result.get('primary_action') == "folder_sales": 
                target_folder = mailbox.get_folder(folder_name='Sales emails')
                msg.move(target_folder)

            
            print("   Marking email as read...")
            try:
                msg.mark_as_read()
                print("   Marked as read")
            except Exception as e:
                print(f"    Could not mark as read: {e}")

        return getattr(msgs, 'delta_token', delta_token)

    except Exception as e:
        print(f" Error accessing {name}: {e}")
        return delta_token


inbox_delta, junk_delta = load_last_delta()
print(f"Monitoring inbox + junk for: {EMAIL_TO_WATCH} … Ctrl-C to stop.")

while True:
    inbox_delta = process_folder(inbox_folder, "INBOX", inbox_delta)
    junk_delta = process_folder(junk_folder, "JUNK", junk_delta)
    #gets the new delta tokens and then saves them,
    save_last_delta(inbox_delta, junk_delta)
    time.sleep(30)
