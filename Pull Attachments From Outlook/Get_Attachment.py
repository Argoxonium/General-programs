import os
import win32com.client
import traceback

def main() -> None:
    # Get user input for email and folder name
    email = input("What is your email: ")
    folder_path = input("Enter the folder path (e.g., 'Inbox/FIM'): ")
    
    #ensure that the save location exists and is ready
    save_path = save_location()


    # Scan Outlook folder and save attachments
    scan_outlook(folder_path, save_path, email)


def scan_outlook(folder_path: str, save_path: str, email: str) -> None:
    # Connect to Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access the root folder of the specified email
    folder = outlook.Folders.Item(email)
    
    # Split the folder path into subfolders
    folders = folder_path.split('/')
    
    # Navigate through subfolders
    target_folder = folder
    for subfolder in folders:
        target_folder = target_folder.Folders.Item(subfolder)

    # Iterate through each item (email) in the target folder
    for item in target_folder.Items:
        try:
            pull_attachment(item, save_path)
        except Exception as e:
            print(f"Skipping an item due to error: {e}")
            traceback.print_exc()  # Optional: log the full traceback for debugging


def pull_attachment(item, save_path: str) -> None:
    try:
        # Check if the item is an email and has attachments
        if item.Attachments.Count > 0:
            # Iterate through each attachment in the email
            for attachment in item.Attachments:
                save_attachment(attachment, save_path)
    except Exception as e:
        print(f"Error processing item: {e}")
        traceback.print_exc()  # Optional: log the full traceback for debugging


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
        traceback.print_exc()  # Optional: log the full traceback for debugging


def save_location() -> os.path:
    # Set the save path to the user's Downloads directory
    save_path = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Ensure the save path exists
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    return save_path

if __name__ == "__main__":
    main()