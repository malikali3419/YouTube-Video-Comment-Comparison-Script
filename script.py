import pandas as pd
import requests
import re
import csv
import logging
from datetime import datetime
import os

logging.basicConfig(
    level=logging.DEBUG,  # Set the log level
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("script.log"),
        logging.StreamHandler()
    ]
)

# Suppress DEBUG logs from external libraries
logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger("selenium").setLevel(logging.WARNING)
logging.getLogger("requests").setLevel(logging.WARNING)

logger = logging.getLogger(__name__)


class YouTubeCommentProcessor:

    def __init__(self, file_path):
        """
        Initializes the YouTubeCommentProcessor with the given file path.

        Args:
            file_path (str): The path to the Excel file containing video data.
        """
        self.file_path = file_path
        try:
            self.data_frame = pd.read_excel(self.file_path)
            logger.info("File Loaded Successfully!")
            logger.info(self.data_frame.head())
        except FileNotFoundError:
            logger.error(f"Error: The file '{self.file_path}' was not found.")
            raise
        except Exception as e:
            logger.error(f"An error occurred while loading the Excel file: {e}")
            raise

    def extract_video_id(self, url):
        """
        Extracts the video ID from a YouTube URL.

        Args:
            url (str): The YouTube URL.

        Returns:
            str: The extracted video ID or None if not found.
        """
        video_id = re.findall(r'(?:v=|\/)([0-9A-Za-z_-]{11}).*', url)
        return video_id[0] if video_id else None

    def get_video_statistics(self, video_id):
        """
        Fetches video statistics from the Hadzy API.

        Args:
            video_id (str): The YouTube video ID.

        Returns:
            dict: The response JSON containing video statistics.
        """
        url = f"https://www.hadzy.com/api/videos/{video_id}"
        response = requests.get(url)
        return response.json()

    def get_comments(self, video_id, video_link, page=1, size=100, show_logs=True):
        """
        Fetches comments from the Hadzy API for a given video.

        Args:
            video_id (str): The YouTube video ID.
            page (int): The page number to retrieve.
            size (int): The number of comments per page (default is 100).

        Returns:
            dict: The response JSON containing comments.
        """
        url = f"https://www.hadzy.com/api/comments/{video_id}?page={page}&size={size}&searchTerms=&author="
        response = requests.get(url)
        if show_logs:
            logger.info(f"Getting Comments from this: {video_link} on page number: {page}")
        return response.json()
    
    def process_comments_in_hadzy(self, video_id):
        url = f"https://www.hadzy.com/api/videos/{video_id}?entity=true"
        response = requests.get(url)
        return response.json()

    def process_video(self, video_id, video_link):
        """
        Processes comments for a single video, fetching all pages up to 1000.

        Args:
            video_id (str): The YouTube video ID.

        Returns:
            tuple: A tuple containing video metadata and a list of comments.
        """
        comments_data = self.get_comments(
            video_id=video_id,
            page=0,
            size=100,
            video_link=video_link,
            show_logs=False
        )
        self.process_comments_in_hadzy(video_id)
        total_page_count = comments_data['pageInfo']['totalPages']
        
        comments_list = []

        for page in range(0, min(total_page_count, 1000) + 1):
            comments_data = self.get_comments(video_id=video_id,
                page=page,
                size=100,
                video_link=video_link
            )
            comments_list.extend(comments_data['content'])

        stats = self.get_video_statistics(video_id)
        title = stats['items'][0]['snippet']['title']
        upload_time = stats['items'][0]['snippet']['publishedAt']
        comment_count = int(stats['items'][0]['statistics']['commentCount'])

        upload_date, upload_time = self.split_datetime(upload_time)

        return title, upload_date, upload_time, comment_count, comments_list

    def split_datetime(self, datetime_str):
        """
        Splits a datetime string into date and time components.

        Args:
            datetime_str (str): The datetime string in ISO 8601 format.

        Returns:
            tuple: A tuple containing the date and time as strings.
        """
        dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
        return dt.date().isoformat(), dt.time().isoformat()

    def process_row(self, row, index):
        """
        Processes a single row of the Excel file, checking conditions and saving comments if met.

        Args:
            row (pd.Series): A row from the DataFrame.
        """
        video_id_a = self.extract_video_id(row['Video_URL_A'])
        video_id_b = self.extract_video_id(row['Video_URL_B'])
        channel_a = row['Channel_Name_A']

        if not video_id_a or not video_id_b:
            return

        stats_a = self.get_video_statistics(video_id_a)
        stats_b = self.get_video_statistics(video_id_b)

        comment_count_a = int(stats_a['items'][0]['statistics']['commentCount'])
        comment_count_b = int(stats_b['items'][0]['statistics']['commentCount'])

        if comment_count_a > 100 and comment_count_b > 100:
            result_a = self.process_video(video_id_a, row['Video_URL_A'])
            result_b = self.process_video(video_id_b, row['Video_URL_B'])

            title_a, upload_date_a, upload_time_a, comment_count_a, comments_list_a = result_a
            title_b, upload_date_b, upload_time_b, comment_count_b, comments_list_b = result_b

            upload_time_date_a = f"{upload_date_a} {upload_time_a}"
            upload_time_date_b = f"{upload_date_b} {upload_time_b}"
            rows_count = max(len(comments_list_a), len(comments_list_b))

            with open(f'{channel_a}_{index + 1}_{rows_count}.csv', 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow([
                    'Video_Title_A', 'Video_Upload_DateTime_A', 'Video_Comment_A', 'Comment_Date_A', 'Comment_Time_A',
                    'Video_Title_B', 'Video_Upload_DateTime_B', 'Video_Comment_B', 'Comment_Date_B', 'Comment_Time_B'
                ])

                max_comments_count = max(len(comments_list_a), len(comments_list_b))
                for i in range(max_comments_count):
                    row_data = []

                    if i < len(comments_list_a):
                        comment_a = comments_list_a[i]
                        comment_date_a, comment_time_a = self.split_datetime(comment_a['publishedAt'])
                        row_data.extend([
                            title_a, upload_time_date_a,
                            comment_a['textDisplay'], comment_date_a, comment_time_a
                        ])
                    else:
                        row_data.extend([''] * 6)

                    if i < len(comments_list_b):
                        comment_b = comments_list_b[i]
                        comment_date_b, comment_time_b = self.split_datetime(comment_b['publishedAt'])
                        row_data.extend([
                            title_b, upload_time_date_b,
                            comment_b['textDisplay'], comment_date_b, comment_time_b
                        ])
                    else:
                        row_data.extend([''] * 6)

                    writer.writerow(row_data)

    def process_file(self):
        """
        Processes the entire Excel file, row by row.
        """
        for index, row in self.data_frame.iterrows():
            self.process_row(row, index)


def find_excel_file():
    """
    Finds the first .xls or .xlsx file in the current directory.

    Returns:
        str: The path to the Excel file, or None if no file is found.
    """
    for file in os.listdir():
        if file.endswith(".xls") or file.endswith(".xlsx"):
            return file
    return None


if __name__ == "__main__":
    try:
        excel_file = find_excel_file()
        if excel_file:
            logger.info(f"Found Excel file: {excel_file}")
            processor = YouTubeCommentProcessor(file_path=excel_file)
            processor.process_file()
        else:
            logger.error("No .xls or .xlsx file found in the current directory.")
    except Exception as e:
        logger.error(f"An error occurred: {e}")
