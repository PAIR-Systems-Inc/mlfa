from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
import os, time
import openai
import json

load_dotenv()

# ENV CONFIG
CLIENT_ID = "e55dc410-9822-47e9-9c31-a17a39b30ce1"
CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EMAIL_TO_WATCH = "mariamahmadpear@outlook.com"

credentials = (CLIENT_ID, None)
openai.api_key = OPENAI_API_KEY

# AUTH - using personal account flow (interactive)
token_backend = FileSystemTokenBackend(token_path=".", token_filename="o365_token.txt")
account = Account(
    credentials,
    auth_flow_type="authorization",  # For personal Outlook accounts
    token_backend=token_backend
)

if not account.is_authenticated:
    # This will open a browser for OAuth login (first time)
    account.authenticate(scopes=['basic', 'message_all'])

# SELECT MAILBOX AND FOLDER
mailbox = account.mailbox(resource=EMAIL_TO_WATCH)
folder = mailbox.inbox_folder()

# EMAIL CLASSIFICATION FUNCTION
def classify_email(subject, body):
    prompt = f"""
You are an email routing assistant for MLFA (Muslim Legal Fund of America), a nonprofit organization.

Routing Rules:
- Legal inquiries → Direct to "Apply for Help" section on website (return "website_redirect")
- Donor-related inquiries (payments, receipts) → Forward to: Mujahid.rasul@mlfa.org, Syeda.sadiqa@mlfa.org
- Sponsorship requests → Forward to: Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org
- Organizational questions → Forward to: Arshia.ali.khan@mlfa.org, Maria.laura@mlfa.org
- Email marketing/sales → Mark for "Sales emails" folder (return "folder_sales")

Analyze this email and return JSON with:
- category: one of ["legal", "donor", "sponsorship", "organizational", "marketing"]
- action: one of ["forward", "website_redirect", "folder_sales"]
- recipients: list of email addresses (empty if not forwarding)
- reason: brief explanation of classification

Subject: {subject}

Body:
{body}
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4.1",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"⚠️ OpenAI classification error: {e}")
        return {}

# MAIN EMAIL WATCH LOOP
last_delta = None
print(f"📬 Watching inbox for: {EMAIL_TO_WATCH} … Ctrl-C to stop.")

while True:
    qs = folder.new_query().select("id", "subject", "sender", "receivedDateTime", "body")
    if last_delta:
        qs = qs.delta_token(last_delta)

    try:
        msgs = folder.get_messages(query=qs)
        for msg in msgs:
            print(f"\n🆕 {msg.received.strftime('%Y-%m-%d %H:%M')} | {msg.sender.address} | {msg.subject}")
            result = classify_email(msg.subject, msg.body)
            print(json.dumps(result, indent=2))

            # Optional: Auto-forward logic (commented out for now)
            # to_forward = msg.forward()
            # for r in result["recipients"]:
            #     to_forward.to.add(r)
            # to_forward.body = f"Auto-routed:\n\n{result['reason']}\n\nOriginal message:\n\n" + msg.body
            # to_forward.send()

        last_delta = getattr(msgs, 'delta_token', None)

    except Exception as e:
        print(f"⚠️ Email loop error: {e}")

    time.sleep(30)
