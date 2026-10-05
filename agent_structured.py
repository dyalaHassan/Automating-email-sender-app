from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv
from langchain.agents import create_agent 
from pydantic import BaseModel, Field
import win32com.client
import pythoncom


load_dotenv(override=True)



def send_email(email_address: str, subject: str, body: str):
    """Create an email using Outlook on this computer."""

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")

        mail = outlook.CreateItem(0)

        mail.To = email_address
        mail.Subject = subject
        mail.Body = body

        mail.Display()

        return f"Email prepared for {email_address}"

    finally:
        pythoncom.CoUninitialize()

# Define the structured output schema
class EmailResponse(BaseModel):
    """Response after sending an email"""
    recipient: str = Field(description="Email address of the recipient")
    subject: str = Field(description="Subject line of the email")
    body: str = Field(description="Body of the email")
    status: str = Field(description="Status of the email (sent/draft)")
    summary: str = Field(description="Brief summary of what was done")


# Initialize Gemini Flash 2.5
llm = ChatGoogleGenerativeAI(
    model="gemini-3.8-flash",
    temperature=0.7,
)

system_prompt = """
You are a helpful assistant. 
YOUR WORKFLOW:
1. You should help the user with various office tasks.
2. If the user asks you to send an email, use the send_email tool. 
3. Create a relevant body and a subject for the email.
4. Do not say the email was sent unless the tool actually sends it.
5. When the tool only opens a draft, report the status as prepared
"""

agent = create_agent(
    model=llm,
    tools=[send_email],
    system_prompt=system_prompt,
    response_format=EmailResponse  # Add structured output
)

user_query = "Prepare an email to diala37assan@gmail.com telling her that I finished the task."

response = agent.invoke(
    {"messages": [{'role': 'user', 'content': user_query}]}
)

# Access the structured response
print("Structured Response:")
print(response['structured_response'])
print("\nFull conversation:")
print(response['messages'][-1].content)