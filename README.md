### README for Sentiment Analysis Script

## Overview

This Python script performs sentiment analysis on comments in an Excel file using the TextBlob library. It calculates the subjectivity, polarity, and intensity of each comment and saves the results to a new Excel file with "_results" appended to the original file name.

## How It Works

1. **Environment Setup**: The script sets up the necessary environment, including installing the required NLTK data for TextBlob.
2. **Data Loading**: It loads comments from an Excel file.
3. **Sentiment Analysis**: The script applies sentiment analysis to each comment, skipping blank comments and stopping at a specified marker.
4. **Results Saving**: The results are saved to a new Excel file in a specified directory.

## Input

- **Excel File**: The input file should be an Excel file located in the `data` directory. The file should contain comments in a specified column.
  - Example file path: `data/test.xlsx`
  - Example column name: `COMMENT EN`

## Output

- **Excel File with Results**: The script saves the results to a new Excel file in the `results` directory, with "_results" appended to the original file name.
  - Example output file path: `results/test_results.xlsx`

## Customization

### File Name

To customize the input file, change the `file_path` variable:

```python
file_path = 'data/your_filename.xlsx'
```

### Column Name

To customize the column containing the comments, change the column name in the script:

```python
comment = row['YOUR_COLUMN_NAME']
```

### Stopping Point

To customize the stopping point, change the marker in the script:

```python
if comment.strip() == 'YOUR_STOPPING_MARKER':
```

## Contact

For further information or questions, please contact:
- **Weixin ID**: kasyan98