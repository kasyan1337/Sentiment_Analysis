import os
import ssl

import nltk
import pandas as pd
from textblob import TextBlob

# Bypass SSL verification
ssl._create_default_https_context = ssl._create_unverified_context

# Download required NLTK corpora
nltk.download('punkt')

# Define the path to your Excel file (assuming it's in the same directory as your script)
file_path = 'data/test.xlsx'

# Load the data from the Excel file
df = pd.read_excel(file_path)


# Define Sentiment Analysis Functions
def get_subjectivity(text):
    return TextBlob(text).sentiment.subjectivity


def get_polarity(text):
    return TextBlob(text).sentiment.polarity


def get_intensity(text):
    polarity = TextBlob(text).sentiment.polarity
    subjectivity = TextBlob(text).sentiment.subjectivity
    return (abs(polarity) + subjectivity) / 2


# Initialize empty columns for Subjectivity, Polarity, and Intensity
df['Subjectivity'] = ''
df['Polarity'] = ''
df['Intensity'] = ''

# Apply the Functions to Your DataFrame, skipping blanks and stopping at '~~~END~~~'
for index, row in df.iterrows():
    comment = row['COMMENT EN']  # DEFINE THE COLUMN NAME FOR WHATEVER YOU WANT TO DO SENTIMENT ANALYSIS ON
    if pd.isna(comment) or comment.strip() == '':
        continue  # Skip blank comments, but keep the index
    if comment.strip() == '~~~END~~~':  # DEFINE THE STOPPING POINT
        break  # Stop processing when reaching '~~~END~~~'

    df.at[index, 'Subjectivity'] = get_subjectivity(comment)
    df.at[index, 'Polarity'] = get_polarity(comment)
    df.at[index, 'Intensity'] = get_intensity(comment)

# Ensure the results directory exists
results_dir = 'results'
os.makedirs(results_dir, exist_ok=True)

# Create the new file name with "_results" appended
base_name = os.path.splitext(os.path.basename(file_path))[0]
new_file_name = f"{results_dir}/{base_name}_results.xlsx"

# Save the Results
df.to_excel(new_file_name, index=False)

print(f"Results saved to {new_file_name}")
