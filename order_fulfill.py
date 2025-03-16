import os
import smtplib
import pandas as pd
import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from dotenv import load_dotenv
import logging
import time
import random
import csv

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("order_notifications.log"),
        logging.StreamHandler()
    ]
)

PORT = 587
EMAIL_SERVER = "smtp.gmail.com"

# Google Sheet info
SHEET_ID = "1ncMwNeHucQcmf6lgB1yT0M8PLcoJmfIfIRX9o9iRhUY"  # Update with your sheet ID
SHEET_NAME = "Orders"  # Update with your sheet name
URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

current_dir = Path(__file__).resolve().parent if "__file__" in locals() else Path.cwd()
envars = current_dir / ".env"
load_dotenv(envars)

# Read environment variables
sender_email = os.getenv("EMAIL")
password_email = os.getenv("PASSWORD")

# Initialize sent orders tracking
SENT_ORDERS_FILE = "sent_orders.csv"

def initialize_sent_orders_file():
    """Create sent orders tracking file if it doesn't exist"""
    if not os.path.exists(SENT_ORDERS_FILE):
        with open(SENT_ORDERS_FILE, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['order_id', 'customer_email', 'notification_type', 'sent_timestamp'])
        logging.info(f"Created new sent orders tracking file: {SENT_ORDERS_FILE}")

def get_sent_orders():
    """Load the list of orders that have already been processed"""
    if not os.path.exists(SENT_ORDERS_FILE):
        initialize_sent_orders_file()
        return set()
    
    sent_orders = set()
    try:
        with open(SENT_ORDERS_FILE, 'r') as f:
            reader = csv.reader(f)
            next(reader)  # Skip header
            for row in reader:
                if row:  # Make sure row is not empty
                    order_id = row[0]
                    notification_type = row[2]
                    sent_orders.add(f"{order_id}_{notification_type}")
        return sent_orders
    except Exception as e:
        logging.error(f"Error reading sent orders file: {e}")
        return set()

def record_sent_order(order_id, customer_email, notification_type):
    """Record that a notification has been sent for an order"""
    with open(SENT_ORDERS_FILE, 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([order_id, customer_email, notification_type, datetime.datetime.now()])
        
def load_orders_data(url):
    """Load orders data from Google Sheet"""
    try:
        df = pd.read_csv(url)
        
        # Convert date columns to datetime if they exist
        date_columns = ['order_date', 'ship_date', 'delivery_date']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df
    except Exception as e:
        logging.error(f"Error loading Google Sheet data: {e}")
        return pd.DataFrame()

def create_order_confirmation_email(customer_email, customer_name, order_id, order_details, total_amount):
    """Create order confirmation email"""
    subject = f"Order Confirmation #{order_id}"
    
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Amazing Store", f"{sender_email}"))
    msg["To"] = customer_email
   
    
    msg.set_content(
        f"""
        Dear {customer_name},
        
        Thank you for your order! We're pleased to confirm that we've received your order #{order_id}.
        
        Order Details:
        {order_details}
        
        Total Amount: ${total_amount:.2f}
        
        We'll send you another email once your order has been shipped.
        
        If you have any questions about your order, please contact our customer service.
        
        Thank you for shopping with Amazing Store!
        """
    )
    
    msg.add_alternative(
        f"""
        <html>
          <body style="font-family: Arial, sans-serif; line-height: 1.6;">
            <h2>Order Confirmation</h2>
            <p>Dear {customer_name},</p>
            <p>Thank you for your order! We're pleased to confirm that we've received your order <strong>#{order_id}</strong>.</p>
            
            <h3>Order Details:</h3>
            <p>{order_details}</p>
            
            <p><strong>Total Amount:</strong> ${total_amount:.2f}</p>
            
            <p>We'll send you another email once your order has been shipped.</p>
            
            <p>If you have any questions about your order, please contact our customer service.</p>
            
            <p>Thank you for shopping with Amazing Store!</p>
          </body>
        </html>
        """,
        subtype="html",
    )
    
    return msg

def create_shipping_notification_email(customer_email, customer_name, order_id, shipping_details, tracking_number):
    """Create shipping notification email"""
    subject = f"Your Order #{order_id} Has Been Shipped"
    
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Amazing Store", f"{sender_email}"))
    msg["To"] = customer_email
    
    
    msg.set_content(
        f"""
        Dear {customer_name},
        
        Great news! Your order #{order_id} has been shipped.
        
        Shipping Details:
        {shipping_details}
        
        Tracking Number: {tracking_number}
        
        You can track your package using the tracking number above at our carrier's website.
        
        If you have any questions, please contact our customer service.
        
        Thank you for shopping with Amazing Store!
        """
    )
    
    msg.add_alternative(
        f"""
        <html>
          <body style="font-family: Arial, sans-serif; line-height: 1.6;">
            <h2>Shipping Confirmation</h2>
            <p>Dear {customer_name},</p>
            <p>Great news! Your order <strong>#{order_id}</strong> has been shipped.</p>
            
            <h3>Shipping Details:</h3>
            <p>{shipping_details}</p>
            
            <p><strong>Tracking Number:</strong> {tracking_number}</p>
            
            <p>You can track your package using the tracking number above at our carrier's website.</p>
            
            <p>If you have any questions, please contact our customer service.</p>
            
            <p>Thank you for shopping with Amazing Store!</p>
          </body>
        </html>
        """,
        subtype="html",
    )
    
    return msg

def create_delivery_confirmation_email(customer_email, customer_name, order_id):
    """Create delivery confirmation email"""
    subject = f"Your Order #{order_id} Has Been Delivered"
    
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Amazing Store", f"{sender_email}"))
    msg["To"] = customer_email
  
    
    msg.set_content(
        f"""
        Dear {customer_name},
        
        Your order #{order_id} has been delivered!
        
        We hope you're enjoying your purchase. If you have a moment, we'd appreciate it if you could leave a review of the products you purchased.
        
        If you have any questions or concerns about your order, please contact our customer service.
        
        Thank you for shopping with Amazing Store!
        """
    )
    
    msg.add_alternative(
        f"""
        <html>
          <body style="font-family: Arial, sans-serif; line-height: 1.6;">
            <h2>Delivery Confirmation</h2>
            <p>Dear {customer_name},</p>
            <p>Your order <strong>#{order_id}</strong> has been delivered!</p>
            
            <p>We hope you're enjoying your purchase. If you have a moment, we'd appreciate it if you could leave a review of the products you purchased.</p>
            
            <p>If you have any questions or concerns about your order, please contact our customer service.</p>
            
            <p>Thank you for shopping with Amazing Store!</p>
          </body>
        </html>
        """,
        subtype="html",
    )
    
    return msg

def process_and_send_notifications(df):
    """Process orders and send appropriate notifications"""
    sent_orders = get_sent_orders()
    sent_count = 0
    
    try:
        # Login to SMTP server
        with smtplib.SMTP(EMAIL_SERVER, PORT) as server:
            server.starttls()
            server.login(sender_email, password_email)
            
            current_date = datetime.datetime.now().date()
            
            for _, row in df.iterrows():
                order_id = str(row['order_id'])
                customer_email = row['customer_email']
                customer_name = row['customer_name']
                
                # Generate shipping and delivery details for the email
                order_details = f"{row['product_name']} x {row['quantity']} - ${row['unit_price']:.2f} each"
                total_amount = row['quantity'] * row['unit_price']
                shipping_details = f"Carrier: {row['shipping_carrier']}\nEstimated delivery: {row['delivery_date'].strftime('%B %d, %Y') if pd.notna(row['delivery_date']) else 'Unknown'}"
                tracking_number = row['tracking_number'] if 'tracking_number' in row and pd.notna(row['tracking_number']) else f"TRK{random.randint(10000000, 99999999)}"
                
                # 1. Send order confirmation for new orders
                if row['status'] == 'new' and f"{order_id}_confirmation" not in sent_orders:
                    msg = create_order_confirmation_email(
                        customer_email,
                        customer_name,
                        order_id,
                        order_details,
                        total_amount
                    )
                    
                    server.send_message(msg)
                    record_sent_order(order_id, customer_email, "confirmation")
                    logging.info(f"Sent order confirmation email for order #{order_id} to {customer_email}")
                    sent_count += 1
                    
                # 2. Send shipping notification for shipped orders
                elif row['status'] == 'shipped' and f"{order_id}_shipping" not in sent_orders:
                    msg = create_shipping_notification_email(
                        customer_email,
                        customer_name,
                        order_id,
                        shipping_details,
                        tracking_number
                    )
                    
                    server.send_message(msg)
                    record_sent_order(order_id, customer_email, "shipping")
                    logging.info(f"Sent shipping notification email for order #{order_id} to {customer_email}")
                    sent_count += 1
                    
                # 3. Send delivery confirmation for delivered orders
                elif row['status'] == 'delivered' and f"{order_id}_delivery" not in sent_orders:
                    msg = create_delivery_confirmation_email(
                        customer_email,
                        customer_name,
                        order_id
                    )
                    
                    server.send_message(msg)
                    record_sent_order(order_id, customer_email, "delivery")
                    logging.info(f"Sent delivery confirmation email for order #{order_id} to {customer_email}")
                    sent_count += 1
                    
                # Add a small delay between emails to avoid rate limiting
                if sent_count > 0 and sent_count % 5 == 0:
                    time.sleep(2)
    
    except Exception as e:
        logging.error(f"Error processing orders and sending notifications: {e}")
    
    return sent_count

def main():
    """Main function"""
    try:
        start_time = datetime.datetime.now()
        logging.info("Starting order notification process")
        
        # Initialize tracking file if it doesn't exist
        initialize_sent_orders_file()
        
        # Load orders data from Google Sheet
        df = load_orders_data(URL)
        
        if df.empty:
            logging.warning("No orders data loaded. Please check your Google Sheet URL.")
            return
        
        # Process orders and send notifications
        sent_count = process_and_send_notifications(df)
        
        end_time = datetime.datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        logging.info(f"Order notification process completed in {duration:.2f} seconds")
        logging.info(f"Total emails sent: {sent_count}")
        
    except Exception as e:
        logging.error(f"Error in main process: {e}")

if __name__ == "__main__":
    main()