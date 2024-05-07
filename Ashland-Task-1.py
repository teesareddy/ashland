import pandas as pd  
# Import the pandas library for data manipulation

df = pd.read_excel('example.xlsx')  
# Read the Excel file 'example.xlsx' and store the data in a DataFrame

column_name = input("Enter the column name: ")  
# Prompt the user to enter the column name to analyze

words_to_count = input("Enter the words to count (separated by commas): ").split(',')  
# Prompt the user to enter the words to count, split by commas

word_counts = {}  
# Initialize an empty dictionary to store the word counts

for word in words_to_count:
    word_count = df[df[column_name].str.contains(word.strip(), case=False)].shape[0]  
    # Count the occurrences of each word in the specified column
    word_counts[word.strip()] = word_count  
    # Store the word count in the dictionary

unique_words_count = df[column_name].value_counts() 
# Calculate the count of unique words in the specified column

word_counts_df = pd.DataFrame(list(word_counts.items()), columns=['Word', 'Count'])  
# Create a DataFrame from the word_counts dictionary
unique_words_count_df = unique_words_count.reset_index()  
# Create a DataFrame from the unique_words_count Series
unique_words_count_df.columns = ['Word', 'Count']  
# Rename the columns of the unique_words_count_df

output_file = 'word_counts_output.xlsx'  
# Define the output file name

with pd.ExcelWriter(output_file) as writer:
    word_counts_df.to_excel(writer, sheet_name='Word_Counts', index=False)  
    # Write the word_counts_df to the 'Word_Counts' sheet
    unique_words_count_df.to_excel(writer, sheet_name='Unique_Words_Count', index=False)  
    # Write the unique_words_count_df to the 'Unique_Words_Count' sheet

print(f"Word counts and unique words count have been saved to '{output_file}'.")  
# Print a message indicating that the output file has been saved