import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# Default Path agnostic of user (user path)

# Specify the path to the folder containing Excel files
folder_path = "C:/Users/wgranalyst/Desktop/AutomationFolder/Input"

# List all files in the folder
files = os.listdir(folder_path)

# Filter out only Excel files
excel_files = [
    file for file in files if file.endswith(".xlsx") or file.endswith(".xls")
]

# Prompt user to choose file

if excel_files:
    # Take the first Excel file found (you may modify this logic as needed)
    excel_file_path = os.path.join(folder_path, excel_files[0])

    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file_path)
else:
    print("No Excel files found in the specified folder.")

# Clears the comment column
df["Comment/Notes: C. Stud Meter, Entertainment, Promotions, Tournaments"] = ""
# Rename the columns to something more manageable and dropping irrelevant columns
df.rename(
    columns={
        "Comment/Notes: C. Stud Meter, Entertainment, Promotions, Tournaments": "comments"
    },
    inplace=True,
)
df.rename(columns={"Select the Casino": "casino"}, inplace=True)
df.rename(columns={"# @ High Stakes Area": "High Stakes"}, inplace=True)


# Function to separate date and time
def separate_date_time(date_time_str):
    date_time_obj = datetime.strptime(date_time_str, "%b %d, %Y %I:%M %p")
    date = date_time_obj.strftime("%m/%d/%Y")
    time = date_time_obj.strftime("%H:%M")
    return date, time


# Apply function to each row and create new columns
df[["date", "time"]] = df["Date and Time of Count"].apply(
    lambda x: pd.Series(separate_date_time(x))
)

# Drop the Date and Time of Count Column
df.drop(columns=["Date and Time of Count"], inplace=True)

# Combine 'First Name' and 'Last Name' into a new column 'rep'
df["rep"] = df["First Name"] + " " + df["Last Name"]

# Drop the 'First Name' and 'Last Name' columns
df.drop(columns=["First Name", "Last Name"], inplace=True)

# Dropping more columns
df.drop(columns=["Geo Stamp"], inplace=True)

# Assuming df is your DataFrame
if "Enter Your Email" in df.columns:
    df.drop(columns=["Enter Your Email"], inplace=True)
else:
    print("'Enter Your Email' column not found, skipping drop operation.")
df.drop(columns=["Submission Date"], inplace=True)

# Assuming df is your DataFrame
if "Timer" in df.columns:
    df.drop(columns=["Timer"], inplace=True)
else:
    print("'Timer' column not found, skipping drop operation.")


# Check if 'High Stakes' column exists in the DataFrame
if "High Stakes" in df.columns:
    # Replace NaN values with 0's in High Stakes Column
    df["High Stakes"].fillna(0, inplace=True)

    # Replace float values with integers in the 'High Stakes' column
    df["High Stakes"] = df["High Stakes"].astype(int)

    # Moving High Stakes Slot Area Values to the comments column
    df["comments"] = df["High Stakes"]

    # Drop the original column '# @ High Stakes Slot Area' if needed
    df.drop("High Stakes", axis=1, inplace=True)

    # Convert comments to string
    df["comments"] = df["comments"].astype(str)

    # Add string to the end of existing string in the 'comments' column
    df["comments"] = df["comments"] + " @ HIGH STAKES SLOT AREA"

    # If 0 @ High stakes slot area delete the comment
    # Check if the value in 'comments' column is equal to '0 @ HIGH STAKES SLOT AREA'
    mask = df["comments"] == "0 @ HIGH STAKES SLOT AREA"

    # Replace the values with blank where mask is True
    df.loc[mask, "comments"] = ""
else:
    print("The 'High Stakes' column does not exist in the DataFrame.")


# Rearrange column positions
import pandas as pd

# Assuming df is your DataFrame

if "Small Baccarat PLAYERS" in df.columns:
    df = df[
        [
            "casino",
            "date",
            "time",
            "rep",
            "comments",
            "Small Craps PLAYERS",
            "Small Craps TABLES",
            "High Craps PLAYERS ($25+)",
            "High Craps TABLES ($25+)",
            "Small Table PLAYERS",
            "Small TABLES",
            "High Table PLAYERS ($25+)",
            "High TABLES ($25+)",
            "Small Slots (1¢ 5¢ 10¢ 25¢ 50¢)",
            "Large Slots ($1 $5 $25 $50+)",
            "Poker PLAYERS",
            "Poker TABLES",
            "Bingo",
            "Small Baccarat PLAYERS",
            "Small Baccarat TABLES",
            "High Baccarat PLAYERS ($25+)",
            "High Baccarat TABLES ($25+)",
        ]
    ]
    pass
else:
    df = df[
        [
            "casino",
            "date",
            "time",
            "rep",
            "comments",
            "Small Craps PLAYERS",
            "Small Craps TABLES",
            "High Craps PLAYERS ($25+)",
            "High Craps TABLES ($25+)",
            "Small Table PLAYERS",
            "Small TABLES",
            "High Table PLAYERS ($25+)",
            "High TABLES ($25+)",
            "Small Slots (1¢ 5¢ 10¢ 25¢ 50¢)",
            "Large Slots ($1 $5 $25 $50+)",
            "Poker PLAYERS",
            "Poker TABLES",
            "Bingo",
        ]
    ]

# Renaming more columns to the correct format
# Search function for game category (Players, Tables)
if "Small Baccarat PLAYERS" in df.columns:
    df.rename(
        columns={
            "Small Craps PLAYERS": "craps|players|-1",
            "Small Craps TABLES": "craps|open|-1",
            "High Craps PLAYERS ($25+)": "craps|players|25",
            "High Craps TABLES ($25+)": "craps|open|25",
            "Small Table PLAYERS": "other tables games|players|-1",
            "Small TABLES": "other tables games|open|-1",
            "High Table PLAYERS ($25+)": "other tables games|players|25",
            "High TABLES ($25+)": "other tables games|open|25",
            "Small Slots (1¢ 5¢ 10¢ 25¢ 50¢)": "small slots|players|-1",
            "Large Slots ($1 $5 $25 $50+)": "large slots|players|-1",
            "Poker PLAYERS": "poker|players|-1",
            "Poker TABLES": "poker|open|-1",
            "Bingo": "bingo|players|-1",
            "Small Baccarat PLAYERS": "baccarat|players|-1",
            "Small Baccarat TABLES": "baccarat|open|-1",
            "High Baccarat PLAYERS ($25+)": "baccarat|players|25",
            "High Baccarat TABLES ($25+)": "baccarat|open|25",
        },
        inplace=True,
    )
    pass
else:
    df.rename(
        columns={
            "Small Craps PLAYERS": "craps|players|-1",
            "Small Craps TABLES": "craps|open|-1",
            "High Craps PLAYERS ($25+)": "craps|players|25",
            "High Craps TABLES ($25+)": "craps|open|25",
            "Small Table PLAYERS": "other tables games|players|-1",
            "Small TABLES": "other tables games|open|-1",
            "High Table PLAYERS ($25+)": "other tables games|players|25",
            "High TABLES ($25+)": "other tables games|open|25",
            "Small Slots (1¢ 5¢ 10¢ 25¢ 50¢)": "small slots|players|-1",
            "Large Slots ($1 $5 $25 $50+)": "large slots|players|-1",
            "Poker PLAYERS": "poker|players|-1",
            "Poker TABLES": "poker|open|-1",
            "Bingo": "bingo|players|-1",
        },
        inplace=True,
    )

# Checks to get rid of 0's for baccarat low
if "baccarat|players|-1" in df.columns:
    for index, row in df.iterrows():
        if row["baccarat|players|-1"] == 0 and row["baccarat|open|-1"] == 0:
            df.at[index, "baccarat|players|-1"] = -1
            df.at[index, "baccarat|open|-1"] = -1
else:
    pass

# Checks to get rid of 0's for baccarat low
if "baccarat|players|-1" in df.columns:
    for index, row in df.iterrows():
        if row["baccarat|players|25"] == 0 and row["baccarat|open|25"] == 0:
            df.at[index, "baccarat|players|25"] = -1
            df.at[index, "baccarat|open|25"] = -1
else:
    pass


# Performing Checks for slot or tables confirmed
df["total_players"] = df[
    [
        "craps|players|-1",
        "craps|players|25",
        "other tables games|players|-1",
        "other tables games|players|25",
    ]
].sum(axis=1, skipna=True)
df["total_tables"] = df[
    [
        "craps|open|-1",
        "craps|open|25",
        "other tables games|open|-1",
        "other tables games|open|25",
    ]
].sum(axis=1, skipna=True)
df["tables_confirmed"] = df["total_players"] / df["total_tables"]

# Converting High Craps to -1 if both players and tables = 0
for index, row in df.iterrows():
    if row["craps|players|25"] == 0 and row["craps|open|25"] == 0:
        df.at[index, "craps|players|25"] = -1
        df.at[index, "craps|open|25"] = -1

print(df.columns)

# Converts 0's in bingo to -1
df["bingo|players|-1"] = df["bingo|players|-1"].apply(lambda x: -1 if x == 0 else x)


# Replace NaN values with -1 in specified columns
columns_to_replace = [
    "craps|players|-1",
    "craps|open|-1",
    "craps|players|25",
    "craps|open|25",
    "other tables games|players|-1",
    "other tables games|open|-1",
    "other tables games|players|25",
    "other tables games|open|25",
    "small slots|players|-1",
    "large slots|players|-1",
    "poker|players|-1",
    "poker|open|-1",
    "bingo|players|-1",
]

df[columns_to_replace] = df[columns_to_replace].fillna(-1)


# Display the DataFrame with new columns
print(df)


# Where do you want to save default output

file_path = "C:/Users/wgranalyst/Desktop/AutomationFolder/Output/output.xlsx"

# Export DataFrame to Excel with specified file path
df.to_excel(file_path, index=False)
