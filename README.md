# Python-Bot-for-Email-Categorization
import win32com.client
import pandas as pd
import openai

# Set your OpenAI API Key
openai.api_key = "your_openai_api_key"

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Select the inbox (or another folder if needed)
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items

# List to store email data
email_data = []

# Loop through emails
for message in messages:
    try:
        subject = message.Subject
        body = message.Body
        received_time = message.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
        
        # Categorize email using GPT-4
        prompt = f"""
        Categorize the following email into one of three categories: 
        - 'New Claim Registration'
        - 'Claim Follow-up'
        - 'Complaint'

        Email Subject: {subject}
        Email Body: {body[:500]}  # Limit text to avoid long prompts
        """
        
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "system", "content": "You are an AI that classifies emails into predefined categories."},
                      {"role": "user", "content": prompt}]
        )
        
        category = response["choices"][0]["message"]["content"].strip()

        # Store email data
        email_data.append([received_time, subject, category])

    except Exception as e:
        print(f"Error processing email: {e}")

# Convert to DataFrame
df = pd.DataFrame(email_data, columns=["Received Time", "Subject", "Category"])

# Save to Excel
df.to_excel("categorized_emails.xlsx", index=False)

print("Emails categorized and saved to Excel successfully.")
