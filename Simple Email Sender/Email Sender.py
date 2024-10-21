import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import boto3
from botocore.exceptions import ClientError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import configparser
import os

# Read from the config file
config = configparser.ConfigParser()
config.read('config.ini')

# Extract configuration sets from the config file
config_sets = ['None'] + [value for key, value in config.items('ConfigSets')]

# Check if dark mode is enabled in config
dark_mode = config['Settings'].getboolean('DarkMode', fallback=False)

# Initialize SES client
ses = boto3.client('ses', region_name='us-east-1')  # Update the region if necessary

# Store selected attachments
attachments = []

# Function to select files for attachments
def select_files():
    files = filedialog.askopenfilenames(title="Select Attachment(s)")
    attachments.extend(files)  # Add selected files to the attachment list
    # Display selected attachments
    file_list_label.config(text=", ".join([os.path.basename(file) for file in attachments]))

# Function to send email using SES with attachments and configuration set
def send_email():
    from_address = from_address_entry.get()
    to_addresses = [email.strip() for email in to_address_entry.get().split(',')]
    cc_addresses = [email.strip() for email in cc_address_entry.get().split(',')] if cc_address_entry.get() else []
    bcc_addresses = [email.strip() for email in bcc_address_entry.get().split(',')] if bcc_address_entry.get() else []
    subject = subject_entry.get()
    message_body = message_entry.get("1.0", tk.END).strip()
    selected_config_set = config_set_var.get()  # Get the selected config set
    tags = tags_entry.get()
    
    # Validate inputs
    if not from_address or not to_addresses or not message_body:
        messagebox.showerror("Error", "Please fill out all required fields: From Address, To Address, and Message.")
        return

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = ', '.join(to_addresses)
    if cc_addresses:
        msg['Cc'] = ', '.join(cc_addresses)
    if bcc_addresses:
        msg['Bcc'] = ', '.join(bcc_addresses)
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(message_body, 'plain'))

    # Attach each file
    for file_path in attachments:
        try:
            attachment = MIMEBase('application', 'octet-stream')
            with open(file_path, 'rb') as f:
                attachment.set_payload(f.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(attachment)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to attach {file_path}: {e}")
            return

    # Prepare raw message
    try:
        # Prepare raw email parameters
        email_params = {
            'Source': from_address,
            'RawMessage': {'Data': msg.as_string()},
        }

        # Add Configuration Set if selected
        if selected_config_set and selected_config_set != 'None':
            email_params['ConfigurationSetName'] = selected_config_set

        # Attempt to send the email using SES
        response = ses.send_raw_email(**email_params)  # Pass the constructed email_params
        messagebox.showinfo("Success", "Email sent successfully!")

    except ClientError as e:
        messagebox.showerror("Error", f"Failed to send email: {str(e)}")


# Set up the Tkinter UI with ttk widgets
root = tk.Tk()
root.title("Email Sender with Attachments and Config Set")
root.geometry("615x550")
root.minsize(500, 550)

# Define custom styles for dark and light modes
style = ttk.Style()

if dark_mode:
    # Dark mode settings
    root.configure(bg='#2E2E2E')  # Set background color for dark mode
    style.configure("TLabel", foreground="white", background="#2E2E2E", font=("Arial", 10))
    style.configure("TButton", font=("Arial", 12), padding=6, foreground="black", background="#3A3A3A")
    style.configure("TEntry", font=("Arial", 12), foreground="black", background="#4E4E4E")
    style.configure("TOptionMenu", font=("Arial", 12), foreground="white", background="#4E4E4E")
else:
    # Light mode settings
    root.configure(bg='white')
    style.configure("TLabel", font=("Arial", 10))
    style.configure("TButton", font=("Arial", 12), padding=6)
    style.configure("TEntry", font=("Arial", 12))
    style.configure("TOptionMenu", font=("Arial", 12))

# Layout grid configuration
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)

# Create input fields and labels
ttk.Label(root, text="From Address:").grid(column=0, row=0, sticky="W", padx=10, pady=5)
from_address_entry = ttk.Entry(root, width=40)
from_address_entry.grid(column=1, row=0, padx=10, pady=5)

ttk.Label(root, text="To Address (comma separated):").grid(column=0, row=1, sticky="W", padx=10, pady=5)
to_address_entry = ttk.Entry(root, width=40)
to_address_entry.grid(column=1, row=1, padx=10, pady=5)

ttk.Label(root, text="CC (optional, comma separated):").grid(column=0, row=2, sticky="W", padx=10, pady=5)
cc_address_entry = ttk.Entry(root, width=40)
cc_address_entry.grid(column=1, row=2, padx=10, pady=5)

ttk.Label(root, text="BCC (optional, comma separated):").grid(column=0, row=3, sticky="W", padx=10, pady=5)
bcc_address_entry = ttk.Entry(root, width=40)
bcc_address_entry.grid(column=1, row=3, padx=10, pady=5)

ttk.Label(root, text="Subject:").grid(column=0, row=4, sticky="W", padx=10, pady=5)
subject_entry = ttk.Entry(root, width=40)
subject_entry.grid(column=1, row=4, padx=10, pady=5)

ttk.Label(root, text="Message:").grid(column=0, row=5, sticky="W", padx=10, pady=5)
message_entry = tk.Text(root, height=8, width=27, font=("Arial", 12))
message_entry.grid(column=1, row=5, padx=10, pady=5)

ttk.Label(root, text="Tags (optional, comma separated):").grid(column=0, row=6, sticky="W", padx=10, pady=5)
tags_entry = ttk.Entry(root, width=40)
tags_entry.grid(column=1, row=6, padx=10, pady=5)

ttk.Label(root, text="Configuration Set (optional):").grid(column=0, row=7, sticky="W", padx=10, pady=5)

# Dropdown for configuration set
config_set_var = tk.StringVar(root)
config_set_var.set('None')  # Default value
config_set_entry = ttk.OptionMenu(root, config_set_var, *config_sets)
config_set_entry.grid(column=1, row=7, padx=10, pady=5)

ttk.Label(root, text="Attachments:").grid(column=0, row=8, sticky="W", padx=10, pady=5)
file_list_label = ttk.Label(root, text="No files selected", wraplength=400)
file_list_label.grid(column=1, row=8, padx=10, pady=5)

# Button to select attachments
attachment_button = ttk.Button(root, text="Add Attachment(s)", command=select_files)
attachment_button.grid(column=1, row=9, padx=10, pady=5)

# Send button
send_button = ttk.Button(root, text="Send Email", command=send_email)
send_button.grid(column=1, row=10, padx=10, pady=20)

root.mainloop()
