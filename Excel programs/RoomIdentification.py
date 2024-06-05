import pandas as pd

def main():
    path = rf"C:\Users\nhorn\Documents\Lab Inspections\Hood & Glovebox Inspections\Files\CFCT_Glovebox_Hood_ID_Owner.xlsx"
    sheet = 'test'
    csv_path = rf'C:\Users\nhorn\Documents\Lab Inspections\Hood & Glovebox Inspections\Files\Owners_list.csv'
    data = pd.read_excel(path,sheet)
    print(data)
    name_df = normalize_names(data['Owner'],csv_path)
    data['H or G'].replace({'H':'Hood', 'G':'Glovebox'}, inplace=True)

    with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
    # Write your DataFrame to a new sheet in the file
        data.to_excel(writer, sheet_name='Full Lab List', index=False)
        name_df.to_excel(writer, sheet_name='Owner List', index=False)

def main1():
    path = rf"C:\Users\nhorn\Documents\Lab Inspections\Hood & Glovebox Inspections\Files\CFCT_Glovebox_Hood_ID_Owner.xlsx"
    csv_path = rf'C:\Users\nhorn\Documents\Lab Inspections\Hood & Glovebox Inspections\Files\Owners_list.csv'
    data = pd.read_excel(path,'Full Lab List')
    names = pd.read_excel(path, 'Owner List')
    name_mapping = create_name_mapping(names)
    df = update_names(data, 'Owner', name_mapping)
    with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
    # Write your DataFrame to a new sheet in the file
        df.to_excel(writer, sheet_name='Full Lab List 2', index=False)

def extract_last_name(name):
    """Extract the last name from a full name string."""
    parts = name.replace(',', ' ').split()
    if len(parts) == 1:
        return parts[0]  # Assume single part is last name
    else:
        return parts[-1]  # Assume the last part is the last name if "First Last"

def create_name_mapping(correct_names):
    """Create a dictionary mapping last names to full correct names."""
    last_name_to_full = {}
    for name in correct_names:
        last_name = extract_last_name(name)
        last_name_to_full[last_name] = name  # Assumes last names are unique
    return last_name_to_full

def update_names(df, column_name, name_mapping):
    """Update all names in the specified DataFrame column based on the last name mapping."""
    # Function to find correct name using last name extracted
    def find_correct_name(full_name):
        last_name = extract_last_name(full_name)
        return name_mapping.get(last_name, full_name)  # Default to original if no match found

    df[column_name] = df[column_name].apply(find_correct_name)
    return df

'''# Example usage:
correct_names = ['Smith, John', 'Doe, Jane', 'Brown, Sam']
data = {'Owner': ['John Smith', 'Jane Doe', 'Sammy Brown', 'S. Brown']}
df = pd.DataFrame(data)

# Create mapping and update DataFrame
name_mapping = create_name_mapping(correct_names)
df = update_names(df, 'Owner', name_mapping)
print(df)'''



def normalize_names(names_series: pd.Series, output_csv_path: str):
    unique_names = set()

    for name in names_series:
        # Remove leading/trailing whitespaces
        name = name.strip()

        # Normalize and split the name
        if ',' in name:
            last, first = name.split(',', 1)
        elif ' ' in name:
            parts = name.split()
            first = parts[0]
            last = ' '.join(parts[1:])
        else:
            first = name
            last = '_'

        # Strip spaces from the split parts
        first = first.strip()
        last = last.strip()

        # Create a consistent format for the full name
        full_name = f"{last}, {first}".title()  # Capitalize names properly

        # Add to set to avoid duplicates
        unique_names.add(full_name)

    # Create a DataFrame from the unique names set
    unique_names_df = pd.DataFrame(list(unique_names), columns=['Full Name'])

    return unique_names_df

# Example usage:
# Assume `df` is your DataFrame and 'Name' is the column with names
# normalize_names(df['Name'], 'output_names.csv')

if __name__ == '__main__':
    main1()