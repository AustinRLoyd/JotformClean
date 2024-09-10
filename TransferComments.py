import pandas as pd
import re
from datetime import datetime

######################################################  This section is dedicated to updating comments from excel comment books    ############################################################
# Loading Dataframes and dropping first row (Day Row)
# Biloxi
df_biloxi_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Biloxi Comments/BiloxiComments.xlsx"
)
# Dropping First Row
df_biloxi_corrected = df_biloxi_original.drop(index=0).reset_index(drop=True)

# Laughlin
df_laughlin_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Laughlin Comments/LaughlinComments.xlsx"
)
# Dropping First Row
df_laughlin_corrected = df_laughlin_original.drop(index=0).reset_index(drop=True)

# Mesquite
df_mesquite_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Mesquite Comments/MesquiteComments.xlsx"
)
# Dropping First Row
df_mesquite_corrected = df_mesquite_original.drop(index=0).reset_index(drop=True)

# NorCal
df_norcal_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Nor Cal Comments/NorCalComments.xlsx"
)
# Dropping First Row
df_norcal_corrected = df_norcal_original.drop(index=0).reset_index(drop=True)

# NorOregon
df_nororegon_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Oregon Comments/NorOregonComments.xlsx"
)
# Dropping First Row
df_noreoregon_corrected = df_nororegon_original.drop(index=0).reset_index(drop=True)

# Phoenix
df_phoenix_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Phoenix Comments/PhoenixComments.xlsx"
)
# Dropping First Row
df_phoenix_corrected = df_phoenix_original.drop(index=0).reset_index(drop=True)

# Shreveport
df_shreveport_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Shreveport Comments/ShreveportComments.xlsx"
)
# Dropping First Row
df_shreveport_corrected = df_shreveport_original.drop(index=0).reset_index(drop=True)

# Tucson
df_tucson_original = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/Tucson Comments/TucsonComments.xlsx"
)
# Dropping First Row
df_tucson_corrected = df_tucson_original.drop(index=0).reset_index(drop=True)


# Concatenating all market comment dataframes to df_melted_corrected dataframe
df_corrected = pd.concat(
    [
        df_biloxi_corrected,
        df_laughlin_corrected,
        df_mesquite_corrected,
        df_norcal_corrected,
        df_noreoregon_corrected,
        df_phoenix_corrected,
        df_shreveport_corrected,
        df_tucson_corrected,
    ],
    ignore_index=True,
)


# Transforming the corrected dataframe to the desired structure
df_melted_corrected = df_corrected.melt(
    id_vars=["Casino"], var_name="Date", value_name="Value"
)

# Removing rows with NaN values in the 'Value' column
df_melted_clean_corrected = df_melted_corrected.dropna(subset=["Value"]).reset_index(
    drop=True
)

# Display the first few rows of the corrected and transformed dataframe
df_melted_clean_corrected.head()

# Assuming df_melted_clean_corrected is your DataFrame
# Convert 'Date' column to datetime, coerce errors
df_melted_clean_corrected["Date"] = pd.to_datetime(
    df_melted_clean_corrected["Date"], errors="coerce"
)

# Reformat the 'Date' column
df_melted_clean_corrected["Date"] = df_melted_clean_corrected["Date"].dt.strftime(
    "%m/%d/%Y"
)


# Export transformed data to Excel
output_file_path = (
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/MeltedComments.xlsx"
)
df_melted_clean_corrected.to_excel(output_file_path, index=False)
print(f"\nExported melted data to {output_file_path}")


###########################################################################################################################################################################################################

#############################################   This section is dedicated to moving comments from MeltedComments to Cleaned Jotform Excel Sheet   #########################################################


# Jotform dataframe
df2 = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Output/combined_output_highlighted.xlsx",
    sheet_name="Sheet1",
)

# Comment dataframe
df1 = pd.read_excel(
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Comments/MeltedComments.xlsx",
    sheet_name="Sheet1",
)

# Convert 'comments' column to string type
df2["comments"] = df2["comments"].astype(str)

# Check the data type of 'comments' column after conversion
comments_dtype = df2["comments"].dtype

print("Data type of df2['comments'] column after conversion:", comments_dtype)


# Function to merge promotions into comments
def merge_promotions(df1, df2):
    for index, row1 in df1.iterrows():
        for index, row2 in df2.iterrows():
            if row1["Casino"] == row2["casino"] and row1["Date"] == row2["date"]:
                df2.at[index, "comments"] += " / " + row1["Value"]
    return df2


# Merge promotions into comments
result_df = merge_promotions(df1, df2)


# Define a function to remove "nan" or "nan / " using regular expressions
def remove_nan_comments(comments):
    return re.sub(r"\bnan\s*(\/\s*)?", "", str(comments))


# Apply the function to the 'comments' column
result_df["comments"] = result_df["comments"].apply(remove_nan_comments)

# Print the DataFrame after removing "nan" or "nan / "
print(result_df)

# Specify the file path
file_path = (
    "C:/Users/wgranalyst/Desktop/AutomationFolder/Output/output_with_comments.xlsx"
)

# Export DataFrame to Excel
result_df.to_excel(
    file_path, index=False
)  # Set index=False if you don't want to include the DataFrame index
