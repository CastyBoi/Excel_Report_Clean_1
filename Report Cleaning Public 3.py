# This program was created with the intent of being semi-modular and dynamic. I set this up with the intention of automating other departments excel file cleaning. The best use cases I have found thus far have been in really big reports where a lot of random rows get dropped, and separated onto different tabs. 
# Usually if you're decent at Excel this stuff doesn't take that long, but as many of us know a lot of people really aren't that good at Excel. 
# I also recommend setting this up and creating an executable that can be shared, with required file paths located on shared drives or such. 

### Imports ###
import pandas as pd
import PySimpleGUI as sg
import os
import datetime
import csv
from datetime import datetime
import openpyxl
### Imports ###

### PySimple User GUI Setup Section ###


# FilePath Output PySimple
sg.theme("LightBlue1")
# Link to graphic of all themes: https://media.geeksforgeeks.org/wp-content/uploads/20200511200254/f19.jpg 

layout = [
    [sg.T("")],
    [sg.Text("Choose the input file: "),
     sg.Input(key="-IN-"),
     sg.FileBrowse(file_types=(("ALL Files", "*.*"), ))],
    [sg.Button("Continue with program")]
]
# Window Instance
window = sg.Window('My File Browser', layout, size=(600,150))

# Create empty file paths that can be changed later on.
input_file_path = ""
output_file_path = ""

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break
    elif event == "Continue with program":
        input_file_path = values['-IN-']
        print(input_file_path)
    break
window.close()


### END PySimple User GUI Setup Section ###


### File Setup Section ###

input_file_1 = f'{input_file_path}' # For production. 
# input_file_1 = r'Your_File_Path' # For development I reccomend hard coding this, and then commenting out the PySimple stuff for quicker revision and checks.

strings_to_delete_file = r"Drive:\Program References\Remove Strings.csv" # This is for the program to reference and read a csv file, then search for exact matches in a specified column to then drop from the DataFrame
# This is cool becasue you can then go in and update the list and the program will still work and search for the new items.

# Setting up the date for the output file name
current_date = datetime.now().strftime('%m-%d-%Y') # Date format can be changed if desired

# Appending date to the end of the filpath name
output_file = r'Your_Output_Path'
output_file_2 = f'{output_file}({current_date}).xlsx' # This way the file is created new in whatever location you set, and it adds the date so if its a daily report the user can identify it easily.


### END File Setup Section ###

### Create and setup Original DF Information ###


# Initial File learning and DataFrame conversion/cleaning
def convert_to_df(input_df):
    global df1
    if input_df.lower().endswith('.csv'):
        df1 = pd.read_csv(input_df)
        return(df1)
    elif input_df.lower().endswith('.xlsx'):
        df1 = pd.read_excel(input_df)
        return(df1)
    else:
        print('FileType Not Recognized')
    return(df1)
# Set up a conditional in case the file type changes from Excel or CSV. 

# Convert input file to a DataFrame for manipulation  
# print('\nInput File\n', input_file_1) # Print can remain commented out, only used for initial setup or testing.
convert_to_df(input_file_1)

# Print Orignal DataFrame
print('\nOriginal DataFrame\n') 
print(df1.head(10)) # Verify it was read correctly

# Get column names as a list 
# The functions basically all use the format col_list[#] to identify a row. Idk why but I found that better than having to write out the names for each column, but hard becasue you gotta count lol.
def get_column_names(input_df):
    global col_list
    col_list = []
    # create a list based on column names in the dataframe
    for col in input_df.columns:
        col_list.append(col)
    return(col_list)            
get_column_names(df1)

# Print Orignal Column Names
# print('\nOriginal Column Names\n')
# print(col_list)


### Create and setup Original DF Information ###

### Setting up Strings to drop Sets Section ###

# strings_to_drop sets 
strings_to_drop_1 = {"String 1", "String 2", "String 3"} # Either by hand
strings_to_drop_2 = set() # or an empty one to be filled 

# Reading through a CSV file to create a set of string to drop
# Open the csv file with all of the strings that need to be dropped
with open(strings_to_delete_file, 'r', newline = '') as csvfile:
    csvreader = csv.reader(csvfile)  
    # Iterate through each row in the CSV
    for row in csvreader:
        # Add the value from second column to the set
        strings_to_drop_2.add(row[1])


### Setting up Strings to drop Sets Section ###

### Functions Section ###


## Function Drop_Null
# Drops Null or NA values from DataFrame
def Drop_Null(input_df):
    input_df.dropna(subset=col_list[0],inplace = True) # Drop rows with empty Plant Name

## Function that drops set of strings
# Outputs a new dataframe WITHOUT whatever information found in the strings_to_drop set    
def Drop_Unused_Strings(input_df, dict_del_1, column_list, column_number_1):
    output_df = input_df[~input_df[column_list[column_number_1]].isin(dict_del_1)] # Checks column for strings then drop rows
    return output_df 

## Function that drops unused columns
# Drop columns that are not used
def drop_columns(input_df, column_list, *args):
    # Args changes based on the users input, so basically it lets the dev dynamically determine which columns to drop. Sick af I barely learned about this BS. 
    columns_to_drop = [column_list[i] for i in args] 
    return input_df.drop(columns = columns_to_drop) 
    # Insert up to whatever number required to drop

## Function to convert specific columns to numbers
# Convert Columns to numbers
def convert_to_float(input_df, column_list, *args):
    columns_to_convert = [column_list[i] for i in args]
    input_df[columns_to_convert] = input_df[columns_to_convert].apply(pd.to_numeric, errors='coerce')
    return input_df

## Function to rename columns if required
# Rename Columns if required
def rename_columns(input_df):
    input_df.rename(columns = {col_list[0]: 'User Defined Name', col_list[1]:'User Defined Name' }, inplace = True)
    # Insert up to whatever number required to change name, and change 'User Defined Name'
# rename_columns()

## Function to convert columns to numbers
# Convert Columns to strings
def convert_to_string(input_df, column_list, *args):
    columns_to_convert = [column_list[i] for i in args]
    input_df[columns_to_convert] = input_df[columns_to_convert].astype(str)
    return input_df

## Function to convert columns that either should be dates or are dates but are kinda jacked
# Set the desired columns when you run the function against the dataframe
def format_date_columns(input_df, date_columns):
    try:
        desired_format = '%m/%d/%Y'
        for column in date_columns:
            if input_df[column].dtype == 'object': 
                # If data is an object convert to datetime only and reformat
                input_df[column] = pd.to_datetime(input_df[column]).dt.date
                input_df[column] = input_df[column].apply(lambda x: x.strftime(desired_format))
            else: 
                # All else, convert to datetime and add the HH:MM:SS and then convert to desired format
                input_df[column] = pd.to_datetime(input_df[column], format='%Y-%m-%d %H:%M:%S').dt.date
                input_df[column] = input_df[column].apply(lambda x: x.strftime(desired_format))
        return input_df
    except ValueError as e:
        print(f'Did not work, error is: {e}')


### Functions Section ###

### Apply functions to DataFrame Section ###


# Apply Drop Unused Strings set 1
df2 = Drop_Unused_Strings(df1, strings_to_drop_1, col_list, 1)
# Apply Drop Unused Strings Set 2
df_strings_dropped = Drop_Unused_Strings(df2, strings_to_drop_2, col_list, 2)
# Check if values are dropped and DF is not broken
print('Dataframe with strings Dropped\n', df_strings_dropped, '\n')

# Apply Covnert to float 
df_float_converted = convert_to_float(df_strings_dropped,col_list,1,2,3)
# Check if values are converted and DF is not broken
print('Dataframe after converting to floats\n', df_float_converted, '\n')

# Apply Date convert and format Function
date_columns = [col_list[6], col_list[8], col_list[20]] # Create a list of columns to convert
df_date_formated = format_date_columns(df_float_converted, date_columns)
# Check if values are converted and DF is not broken
print('Dataframe after converting to floats\n', df_date_formated, '\n')

# Apply Convert to string function
df_string_converted = convert_to_string(df_date_formated, col_list,1,2,3)
# Check if values are converted and DF is not broken
print('Dataframe after converting to floats\n', df_date_formated, '\n')

# Apply dropped columns 
df_dropped_col = drop_columns(df_string_converted, col_list, 4,5,6)   
# Check if values are converted and DF is not broken
print('Dataframe after converting to floats\n', df_dropped_col, '\n')

# Change name
df_final_format = df_dropped_col

# Apply function to update the col_list variable
get_column_names(df_final_format)
print('\nFinal DataFrame Col_List\n', col_list, '\n')
                
# Final full dataframe
print('\nFinal DataFrame\n')
print(df_final_format.sort_values(by=[col_list[5]],ascending = False).head(10)) # Sort by whatever column you want


### Apply functions to DataFrame Section ###

### Setup a seperation of dataframes into unique DFs ###


# Make Separate Dataframes based on the plant location
# Identify each unique dataframe
unique_dfs = df_final_format[col_list[1]].unique()

# Create a dictionary to store the seprate dataframes
unique_dfs_final = {}

# Iterate through unique plant names and create the separate dataframes
for item in unique_dfs:
    unique_dfs_final[item] = df_final_format[df_final_format[col_list[1]]==item] # Sepcify what identifier will be used 

# Print new Plant specific dataframes 
print('Unique Dataframes:')
for item, unique_dfs in unique_dfs_final.items():
    print(f"Plant: {item}\n{unique_dfs.head(10)}\n")

# Assign the desired order of tabs in the Excel file
unique_tab_order = ['Tab 1', 'Tab 2', 'Tab 3']


### Setup a seperation of dataframes into unique DFs ###

### Final Function that writes to excel, re-orders the tabs, and then waits for usre input before opening the file ###


try:
    # Create an Excel file with tabs for each dataframe in a specific order
    with pd.ExcelWriter(output_file_2, engine='openpyxl') as writer:
        for item in unique_tab_order:
            unique_df_single = unique_dfs_final.get(item) 
            if unique_df_single is not None:
                # Sort and write to Excel
                unique_df_single.sort_values(by=[col_list[4]], ascending=True).to_excel(writer, sheet_name=item, index=False)
        
        print('\nOutput to Excel Successful\n')
        
        # Auto-adjust column lengths based on cell values
        workbook = writer.book
        for item in unique_tab_order:
            worksheet = workbook[item]
            
            for column_cells in worksheet.columns:
                max_length = 0 
                column = column_cells[0].column_letter
                
                for cell in column_cells:
                    cell_value = cell.value
                    if cell_value is not None and isinstance(cell_value, str):
                        # Only consider cells that are strings
                        cell_value = str(cell_value)
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
                
        # Print out some info for the user to read
        print(f'DataFrames exported to {output_file_2} on different tabs.')
        
        # Wait for user input
        input('Press Enter to continue')
        
        # Start the new file
        os.startfile(output_file_2)
except Exception as e:
    print(f'An error occurred: {e}')
    # Wait for user input
    input('Press Enter to exit')


### Final Function that writes to excel, re-orders the tabs, and then waits for usre input before opening the file ###

### Closing Notes ###
# A consisten file naming convention helps
# To create an executable some learnings I have found are:
# Setup a Virtual Environment to keep the Python Packages in one place
# I think some IDEs may have weird packages installed that may require you to omit them from the pyinstall command
# 
## pyinstaller --onefile "Drive:\Path_to_your_file.py" May need to exclude some packages to get this to work properly