import os
import win32com.client
import traceback
import configparser

def main() -> None:
    #get config information
    config = start_config(r"C:\Users\nhorn\Documents\Program Files Storage\config get_attachment.ini")
    
    # Get user input for email and folder name
    email = config.get('Path','Email')
    partial_folder_path = config.get('Path','folder_path')
    month = input('what month would you like to download?: ')
    year = input('What year would you like to download?: ')
    month_year = f'{month} {year}'
    
    # Construct the full folder path by appending the month to the partial path
    folder_path = os.path.join(partial_folder_path, month_year)
    
    #ensure that the save location exists and is ready
    save_path = save_location(config.get('Path','save_path'))

    # Scan Outlook folder and save attachments
    scan_outlook(folder_path, save_path, email)

    print('Attachment extraction complete. Files have been extracted to the Downloads folder')
    input('Enter anything to Exit program:')

def start_config(config_path:str) -> configparser.ConfigParser:
    # Initialize the ConfigParser
    config = configparser.ConfigParser()

    # Check if the provided path exists
    if os.path.exists(config_path):
        # Read the configuration file
        config.read(config_path)
        print(f"Configuration loaded from {config_path}")
    else:
        # Raise an error if the configuration file does not exist
        raise FileNotFoundError(f"The configuration file at {config_path} was not found.")

    return config

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


def save_location(folder_location: str='Downloads') -> os.path:
    # Set the save path to the user's Downloads directory
    save_path = os.path.join(os.path.expanduser('~'), folder_location)

    # Ensure the save path exists
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    return save_path

if __name__ == "__main__":
    main()