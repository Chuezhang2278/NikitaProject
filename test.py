import pandas as pd
import pprint

# Define file paths
input_file_path = 'input2.xlsx'  # Path to the input Excel file
output_file_path = 'cleaned_output.xlsx'  # Path to the output Excel file

# Load the Excel file
try:
    # Read the data from the first sheet
    df = pd.read_excel(input_file_path, header=0)
    print("Data read successfully:")
    print(df.head())
except Exception as e:
    print(f"Error reading the Excel file: {e}")
    exit()

# Check if the necessary columns exist
required_columns = ['Item', 'test', 'Description', 'Qty. Sold']
if not all(col in df.columns for col in required_columns):
    print(f"Required columns not found. Columns present: {df.columns}")
    exit()

#This bit of code is actually completely unecessary. Just keeping it in here because im lazy to re-do some naming#

my_df = []
my_dict = {}

for index, row in df.iterrows(): 
    Item = row['Item'] 
    Test = row['test']
    Description = row['Description'] 
    Qty = row['Qty. Sold']

    my_df.append([Item, Test, Description, Qty])

#This bit of code is actually completely unecessary. Just keeping it in here because im lazy to re-do some naming#

df_test = pd.DataFrame(my_df, columns=['Item','Test','Description','Qty'])  
df2 = df_test.groupby(["Description"], sort=False, as_index=False).agg({"Item":'first', "Test":'first', "Qty":"sum"})
df2.to_excel("output.xlsx", index=False)
