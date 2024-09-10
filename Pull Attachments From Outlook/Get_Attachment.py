import os
import win32com.client


def main() -> None:
    # Get user input for email and folder name
    email = input("What is your email: ")
    folder_name = input("What folder would you like us to review: ")
    save_path = r"C:\Users\nhorn\Downloads"  # Replace with your desired save path

    # Ensure the save path exists
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    # Scan Outlook folder and save attachments
    scan_outlook(folder_name, save_path, email)

def scan_outlook(folder_name: str, save_path: str, email: str) -> None:
    # Connect to Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access the specified folder
    folder = outlook.Folders.Item(email)  # Use the email provided by the user
    target_folder = folder.Folders.Item(folder_name)

    # Iterate through each item (email) in the folder
    for item in target_folder.Items:
        pull_attachment(item, save_path)

def pull_attachment(item, save_path: str) -> None:
    # Check if the item is an email and has attachments
    if item.Attachments.Count > 0:
        # Iterate through each attachment in the email
        for attachment in item.Attachments:
            save_attachment(attachment, save_path)

def save_attachment(attachment, save_path: str) -> None:
    # Construct the full path to save the attachment
    attachment_path = os.path.join(save_path, attachment.FileName)

    # Ensure the directory exists
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    # Save the attachment to the specified path
    try:
        attachment.SaveAsFile(attachment_path)
        print(f"Attachment {attachment.FileName} saved to {save_path}")
    except Exception as e:
        print(f"Failed to save attachment {attachment.FileName}: {e}")

if __name__ == "__main__":
    main()