# Automate sending emails
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from email import encoders
import os
import tkinter as tk


# Function to generate the attachment for a given recipient
def generate_attachment(recipient):
    # Generate the attachment
    attachment_path = os.path.join(
        "/", "path", "to", "folder", f"{recipient}'s Portfolio.pdf")
    # Return the path to the generated attachment
    return attachment_path


# Read a txt file that contain the names and emails of the clients and extract them to dictionary
recipients = {}
try:
    with open('/path/to/folder/clients_list.txt', 'r') as f:
        lines = f.readlines()
    for line in lines:
        try:
            if line and not line.isspace():
                name, email = line.strip().split(',')
                recipients[name.strip()] = email.strip()
        except ValueError:
            print(f"Invalid line: {line}")
except (FileNotFoundError, IOError):
    print("Could not read file: clients_list.txt")

# Email parameters
smtp_server1 = "smtp.gmail.com"
smtp_port1 = 587
smtp_port2 = 465
smtp_server2 = "smtp.office365.com"
smtp_server3 = "smtp.mail.yahoo.com"

# Function to update the feedback box


def update_feedback(message):
    feedback_box.insert('end', message)
    feedback_box.update_idletasks()


# Function to handle the "Send" button click
def send_email():
    # Get the subject and body input from the user
    subject = subject_entry.get()
    # Get all the text in the body field, except the last newline character
    body = body_entry.get("1.0", "end-1c")
    count = 0

    # Sender's email address and credentials
    sender = email_entry.get()
    password = password_entry.get()

    # Check the email domain
    domain_name = sender.split("@")[1]
    domain = domain_name.split(".")[0]
    if domain.lower() == "gmail":
        smtp_server = smtp_server1
        smtp_port = smtp_port1
    elif domain.lower() == "hotmail" or domain.lower() == "outlook":
        smtp_server = smtp_server2
        smtp_port = smtp_port1
    elif domain.lower() == "yahoo":
        smtp_server = smtp_server3
        smtp_port = smtp_port2
    else:
        pass

    # Loop over the recipients and send the email with the corresponding attachment
    for name, email in recipients.items():
        # Generate the attachment for the current recipient
        attachment_path = generate_attachment(name)

        # Create a multipart message
        message = MIMEMultipart()
        message["From"] = sender
        message["To"] = email
        message["Subject"] = subject

        # Attach the body of the email
        message.attach(MIMEText(body))

        # Attach the file to the email
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "pdf")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition",
                                f"attachment; filename={os.path.basename(attachment_path)}")
                message.attach(part)
        # Exception when pdf file is missing - skip this recipient
        except FileNotFoundError:
            print(
                f"Attachment not found for {name}.\n")
            update_feedback(
                f"Attachment not found for {name}.\n")

        # Attach the JPEG file to the email
        try:
            with open("/path/to/folder/Monthly_Chart.jpeg", "rb") as jpeg_attachment:
                part = MIMEBase("image", "jpeg")
                part.set_payload(jpeg_attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition", f"attachment; filename={os.path.basename('/path/to/folder/Monthly_Chart.jpeg')}"
                )
                message.attach(part)
        # Exception when jpeg is missing - send without it
        except FileNotFoundError:
            print('Image file not found, sending email without Monthly Chart ...\n')
            update_feedback(
                f'Image file not found, sending email without Monthly Chart ...\n')

        # Send the email
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                try:
                    server.ehlo()
                    server.starttls()
                    server.login(sender, password)
                    server.sendmail(sender, email, message.as_string())
                    print(f"Email sent to {name} ({email})")
                    update_feedback(f"Email sent to {name} ({email}).\n")
                    count += 1
                # Wrong credentials
                except smtplib.SMTPAuthenticationError as e:
                    print(f"Failed to authenticate: {e}")
                    error_message = e.args[1].decode()
                    start_index = error_message.index(
                        "Username and Password not accepted.")
                    end_index = error_message.index("Learn more at")
                    message_error = error_message[start_index:end_index].strip(
                    )
                    update_feedback(
                        f"Failed to authenticate: {message_error}\n")
        except:
            update_feedback(f"Error sending email to {name} ({email}).\n")
    # Check if all clients on the dictionary got an email
    if count == len(recipients):
        print(f"All clients received an email.\n")
        update_feedback(f"All clients received an email. Close the window.\n")
    else:
        print(f"Some clients didn't receive email. Close the window.\n")
        update_feedback(
            f"Some clients didn't receive email. Close the window.\n")

    # Close the SMTP connection
    # server.quit()
    # Close the window
    # window.destroy()


#######################################################################
# Tkinter - Create the main window
window = tk.Tk()
window.title("Email Automation")
window.geometry("1300x450")
font = ("Arial", 12)

# Add labels for inputs
email_label = tk.Label(window, text="Your Email:")
email_label.grid(row=0, column=0)
email_entry = tk.Entry(window, width=70)
email_entry.grid(row=0, column=1)

password_label = tk.Label(window, text="Password:")
password_label.grid(row=1, column=0, padx=10, pady=10)
password_entry = tk.Entry(window, show="*", width=70)
password_entry.grid(row=1, column=1, padx=10, pady=10)

# Add labels for the fields
tk.Label(window, text="Subject:").grid(row=2, column=0, padx=10, pady=10)
tk.Label(window, text="Body:").grid(row=3, column=0, padx=10, pady=10)
feedback_box = tk.Text(window, height=15, width=70)
close_button = tk.Button(window, text="CLOSE",
                         background="#FFB6C1", command=window.destroy, font=("Arial", 10, "bold"))
tk.Label(window, text="Feedback:").grid(row=2, column=2, padx=10, pady=10)

# Add a scrollbar
scrollbar = tk.Scrollbar(window)
scrollbar.grid(row=3, column=3, sticky='NS')

# Configure the feedback_box to use the scrollbar
feedback_box.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=feedback_box.yview)

# Add entry fields and Position the entry fields
subject_entry = tk.Entry(window, width=70)
subject_entry.grid(row=2, column=1, padx=10, pady=10)
body_entry = tk.Text(window, width=70, height=15)
body_entry.grid(row=3, column=1, padx=10, pady=10)
feedback_box.grid(row=3, column=2, padx=10, pady=10)
close_button.grid(row=4, column=2, padx=10, pady=10)

# Add a "Send" button
send_button = tk.Button(window, text="SEND",
                        background="light blue", command=send_email, font=("Arial", 10, "bold"))
send_button.grid(row=4, column=1, padx=10)

# Run the main loop of the GUI
window.mainloop()
