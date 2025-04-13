import os
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def mark_attendance(student_id, name, subject, date, time):
    """Mark attendance for a student and return True if successful"""
    try:
        file_path = f"AttendanceRecords/{subject}_attendance.xlsx"
        
        # Create directory if it doesn't exist
        os.makedirs("AttendanceRecords", exist_ok=True)
        
        # Create or load existing attendance file
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
        else:
            df = pd.DataFrame(columns=["ID", "Name", "Date", "Time"])
        
        # Check if student already marked attendance today
        today_records = df[(df["ID"] == student_id) & (df["Date"] == date)]
        if not today_records.empty:
            return False  # Already marked
        
        # Add new attendance record
        new_record = {"ID": student_id, "Name": name, "Date": date, "Time": time}
        df = df._append(new_record, ignore_index=True)
        
        # Save to Excel
        df.to_excel(file_path, index=False)
        
        return True
    except Exception as e:
        print(f"Error marking attendance: {str(e)}")
        return False

def send_attendance_email(student_email, student_name, subject, date):
    """Send attendance confirmation email to student"""
    try:
        # Email configuration
        sender_email = "accf40075@gmail.com"  # Replace with your Gmail
        sender_password = "qcsleeovwivqchdt"  # Use App Password for Gmail
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = student_email
        msg['Subject'] = f"Attendance Confirmation - {subject}"
        
        # Email body
        body = f"""
        Hello {student_name},
        
        This is to confirm that your attendance has been marked for {subject} on {date}.
        
        Thank you,
        MMCOE
        """
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach Excel file if needed
        file_path = f"AttendanceRecords/{subject}_attendance.xlsx"
        if os.path.exists(file_path):
            attachment = open(file_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {subject}_attendance.xlsx")
            msg.attach(part)
            attachment.close()
        
        # Connect to Gmail SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        
        # Send email
        text = msg.as_string()
        server.sendmail(sender_email, student_email, text)
        server.quit()
        
        return True
    except Exception as e:
        print(f"Error sending email: {str(e)}")
        return False

def get_student_emails():
    """Load student emails from CSV file"""
    student_emails = {}
    try:
        email_file = "TrainingImageLabels/StudentEmails.csv"
        if os.path.exists(email_file):
            df = pd.read_csv(email_file)
            for _, row in df.iterrows():
                student_emails[int(row['ID'])] = row['Email']
    except Exception as e:
        print(f"Error loading student emails: {str(e)}")
    
    return student_emails