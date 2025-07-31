from dotenv import load_dotenv
from openai import OpenAI
import os
import json

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

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
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
    )
    return json.loads(response.choices[0].message.content)

# Test cases
if __name__ == "__main__":
    test_cases = [
        ('Need legal help', 'I need help with a discrimination case'),
        ('Donation receipt', 'I need a receipt for my donation last month'),
        ('Sponsorship inquiry', 'We want to sponsor your next event')
    ]
    
    for subject, body in test_cases:
        print(f"Subject: {subject}")
        result = classify_email(subject, body)
        print(json.dumps(result, indent=2))
        print("-" * 50)