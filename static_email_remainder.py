import os
import smtplib
import datetime
import time
import pandas as pd
from email.message import EmailMessage
from email.utils import formataddr, formatdate
from pathlib import Path
from dotenv import load_dotenv


PORT = 587
EMAIL_SERVER = "smtp.gmail.com"


SHEET_ID = "1xCRgkfupWUtHqu2d3DxOPvwOLhYlujiu6HVHwICDV0g"  # Update with your sheet ID
SHEET_NAME = "Sheet1"  # Update with your sheet name
URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"


current_dir = Path(__file__).resolve().parent if "__file__" in locals() else Path.cwd()
envars = current_dir / ".env"
load_dotenv(envars)


sender_email = os.getenv("EMAIL")
password_email = os.getenv("PASSWORD")

def create_event_reminder_email(receiver_email, name, event_time, event_details):
    """Create email message for event reminder"""
    formatted_event_time = event_time.strftime("%I:%M %p on %B %d, %Y")
    subject = f"Reminder: Event at {event_time.strftime('%I:%M %p')}"
    
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Event Reminder", f"{sender_email}"))
    msg["To"] = receiver_email
    msg["Date"] = formatdate(localtime=True)
    
    msg.set_content(
        f"""
        Hi {name},
        
        This is a reminder about your upcoming event:
        
        Event Time: {formatted_event_time}
        Details: {event_details}
        
        Best regards,
        Your Reminder Service
        """
    )
    
    msg.add_alternative(
        f"""
        <html>
          <body>
            <p>Hi {name},</p>
            <p>This is a reminder about your upcoming event:</p>
            <p><strong>Event Time:</strong> {formatted_event_time}</p>
            <p><strong>Details:</strong> {event_details}</p>
            <p>Best regards,</p>
            <p>Your Reminder Service</p>
          </body>
        </html>
        """,
        subtype="html",
    )
    
    return msg

def create_final_notification_email(receiver_email, num_scheduled, last_email_time):
    """Create final notification email"""
    subject = "All Event Reminders Have Been Scheduled"
    
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Event Reminder", f"{sender_email}"))
    msg["To"] = receiver_email
    msg["Date"] = formatdate(localtime=True)
    
    msg.set_content(
        f"""
        Hi Admin,
        
        All {num_scheduled} event reminders have been scheduled. 
        The last one will be sent at {last_email_time.strftime('%I:%M %p')}.
        
        Please refresh your Google Sheet and rerun the script for new events.
        
        Best regards,
        Your Reminder Service
        """
    )
    
    msg.add_alternative(
        f"""
        <html>
          <body>
            <p>Hi Admin,</p>
            <p>All {num_scheduled} event reminders have been scheduled.</p>
            <p>The last one will be sent at {last_email_time.strftime('%I:%M %p')}.</p>
            <p>Please refresh your Google Sheet and rerun the script for new events.</p>
            <p>Best regards,</p>
            <p>Your Reminder Service</p>
          </body>
        </html>
        """,
        subtype="html",
    )
    
    return msg

def load_events_data(url):
    """Load events data from Google Sheet"""
    try:
        df = pd.read_csv(url)
        df['Date'] = df['Date'].str.strip()
        df['Time'] = df['Time'].str.strip()
        df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'], format="%m/%d/%Y %I:%M:%S %p")
        
        return df
    except Exception as e:
        print(f"Error loading Google Sheet data: {e}")
        return pd.DataFrame()

def schedule_and_send_emails(df, receiver_email=f"{sender_email}"):
    """Process events and schedule reminder emails"""
    now = datetime.datetime.now()
    min_schedule_time = now + datetime.timedelta(minutes=0)
    
    scheduled_emails = []
    ignored_events = []
    
    for _, row in df.iterrows():
        event_time = row['DateTime']
        reminder_minutes = int(row['Reminder Before (minutes)'])
        
        reminder_time = event_time - datetime.timedelta(minutes=reminder_minutes)
        
        if reminder_time >= min_schedule_time:
            scheduled_emails.append({
                'event_details': row['Details'],
                'scheduled_time': reminder_time,
                'event_time': event_time
            })
            print(f"Email will be scheduled for {reminder_time}: {row['Details']}")
        else:
            ignored_events.append({
                'event_time': event_time,
                'event_details': row['Details'],
                'reminder_time': reminder_time
            })
            print(f"Ignored event (reminder time too soon): {row['Details']} at {event_time}")
    
    scheduled_emails.sort(key=lambda x: x['scheduled_time'])
    
    if not scheduled_emails:
        print("No emails to schedule.")
        return {
            'scheduled': [],
            'ignored': ignored_events,
            'total_scheduled': 0,
            'total_ignored': len(ignored_events)
        }
    
    last_email_time = scheduled_emails[-1]['scheduled_time']
    final_notification_time = last_email_time + datetime.timedelta(minutes=5)
    
    scheduled_emails.append({
        'event_details': "Final notification",
        'scheduled_time': final_notification_time,
        'event_time': final_notification_time,
        'is_final': True
    })
    
    with smtplib.SMTP(EMAIL_SERVER, PORT) as server:
        server.starttls()
        server.login(sender_email, password_email)
        
        for i, email_info in enumerate(scheduled_emails):
            scheduled_time = email_info['scheduled_time']
            
            wait_seconds = (scheduled_time - datetime.datetime.now()).total_seconds()
            
            if wait_seconds > 0:
                print(f"Waiting {wait_seconds} seconds until {scheduled_time} to send email...")
                time.sleep(wait_seconds)
            
            if email_info.get('is_final', False):
                msg = create_final_notification_email(
                    receiver_email=receiver_email,
                    num_scheduled=len(scheduled_emails) - 1,  # Don't count the final notification
                    last_email_time=last_email_time
                )
            else:
                msg = create_event_reminder_email(
                    receiver_email=receiver_email,
                    name="Event Participant",
                    event_time=email_info['event_time'],
                    event_details=email_info['event_details']
                )
            
            server.sendmail(sender_email, receiver_email, msg.as_string())
            print(f"Email sent at {datetime.datetime.now()}: {email_info['event_details']}")
    
    return {
        'scheduled': scheduled_emails,
        'ignored': ignored_events,
        'total_scheduled': len(scheduled_emails) - 1,  # Don't count final notification
        'total_ignored': len(ignored_events)
    }

def main():
    df = load_events_data(URL)
    if df.empty:
        print("No events data loaded. Please check your Google Sheet URL.")
        return

    result = schedule_and_send_emails(df)
    
    print("\nSummary:")
    print(f"Total emails scheduled and sent: {result['total_scheduled']}")
    print(f"Total events ignored (reminder too soon): {result['total_ignored']}")
    
    if result['ignored']:
        print("\nIgnored events:")
        for event in result['ignored']:
            print(f"  - {event['event_details']} at {event['event_time']}")

if __name__ == "__main__":
    main()