# YouTube Video Comment Comparison Script

This script processes YouTube video data from an Excel file, compares the comments of two videos in each row, and saves the results in CSV files. The script is designed to check if both videos have more than 100 comments before proceeding with comment extraction.

## Features

- **Excel File Reading**: Automatically finds and reads the first `.xls` or `.xlsx` file in the current directory.
- **YouTube Video Processing**: Extracts video IDs from the provided URLs and fetches video statistics and comments from the Hadzy API.
- **Comment Comparison**: Compares comments between two videos if both have more than 100 comments.
- **CSV Export**: Saves the compared comments into CSV files, with each row representing a pair of comments from the two videos.

## Prerequisites

Before running the script, ensure you have the following installed:

- **Python 3.7 or higher**
- Required Python libraries:
  - `pandas`
  - `requests`
  - `openpyxl`

You can install the necessary libraries using the following command:

```bash
pip install -r requirements.txt
```

## Step 1: Prepare Your Excel File

- **Place your Excel file in the same directory as the script**.
- **Ensure the Excel file has the following columns**:
  - **Video_URL_A: URL of the first video.**
  - **Video_URL_B: URL of the second video.**
  - **Channel_Name_A: Name of the channel associated with the first video.**

### The script will automatically find the first .xls or .xlsx file in the directory.

## Step 2: Run the Script

**Navigate to the directory containing the script and the Excel file, then run the following command:**

```bash
python script.py
```

### The script will:

1. Locate the Excel file.
2. Read the URLs from each row.
3. Fetch video statistics to check if both videos have more than 100 comments.
4. If both videos meet the criteria, the script will fetch the comments and compare them.
5. The compared comments will be saved as CSV files in the same directory.

### Output:

- For each row in the Excel file where both videos have more than 100 comments, a CSV file will be generated.
- The CSV file will be named following this pattern: ChannelName_row_number_CommentsCount.csv.

## Logs

- **The script logs its activity and errors to a script.log file.**
- **Important messages are also printed to the console.**

## Example CSV Output

Each generated CSV file contains the following columns:

| Video_Title_A   | Video_Upload_DateTime_A | Video_Comment_A     | Comment_Date_A | Comment_Time_A | Video_Title_B   | Video_Upload_DateTime_B | Video_Comment_B     | Comment_Date_B | Comment_Time_B |
| --------------- | ----------------------- | ------------------- | -------------- | -------------- | --------------- | ----------------------- | ------------------- | -------------- | -------------- |
| Example Title A | 2024-01-01T12:00:00Z    | This is a comment A | 2024-01-01     | 12:00:00       | Example Title B | 2024-01-01T12:00:00Z    | This is a comment B | 2024-01-01     | 12:00:00       |

- **Video_Title_A**: The title of the first video.
- **Video_Upload_DateTime_A**: The upload date and time of the first video.
- **Video_Comment_A**: The comment on the first video.
- **Comment_Date_A**: The date when the comment was posted.
- **Comment_Time_A**: The time when the comment was posted.
- **Video_Title_B**: The title of the second video.
- **Video_Upload_DateTime_B**: The upload date and time of the second video.
- **Video_Comment_B**: The comment on the second video.
- **Comment_Date_B**: The date when the comment was posted.
- **Comment_Time_B**: The time when the comment was posted.

## Error Handling

The script includes error handling to manage issues such as:

- **File Not Found**: Logs an error if the specified Excel file cannot be found.
- **General Exceptions**: Logs any other errors that occur during processing.

Errors are logged to `script.log`, and important messages are also printed to the console.

## Usage Instructions

### Prepare Your Excel File

Place your Excel file in the same directory as the script. The file should have the following columns:

- `Video_URL_A`: URL of the first video.
- `Video_URL_B`: URL of the second video.
- `Channel_Name_A`: Name of the channel associated with the first video.

The script will automatically find the first `.xls` or `.xlsx` file in the directory.

There is also a Sample Excel File in the repository
