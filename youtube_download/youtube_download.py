#!/usr/bin/python
# coding: utf-8

import sys
import os
import socket
import argparse
import re
import time
from tkinter import *
from tkinter import messagebox
from pytube import YouTube
from pytube.exceptions import AgeRestrictedError

import yt_dlp   # pip install yt-dlp

from youtube_transcript_api import (
    YouTubeTranscriptApi,
    TranscriptsDisabled,
    NoTranscriptFound,
    CouldNotRetrieveTranscript
)

from datetime import datetime
from openpyxl import Workbook

from openpyxl.utils.exceptions import IllegalCharacterError



# --------------------------------------------------------------------------
# Globals
# --------------------------------------------------------------------------
d = datetime.today()
todaysDate = d.strftime("%Y-%m-%d %H:%M:%S")
todaysDateTime = d.strftime("%Y-%m-%d_%H-%M-%S")

author = 'LincolnLandForensics'
description = "Download a list of Youtube videos from videos.txt, save list in xlsx file"
version = '1.2.2'

# will be set in main()
filename = None
spreadsheet = None
Row = None
workbook = None
Sheet1 = None

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    global filename, Row, spreadsheet

    filename = 'videos.txt'
    Row = 2
    spreadsheet = f'videos_{todaysDateTime}.xlsx'

    status = internet()
    if not status:
        noInternetMsg()
        print('\nCONNECT TO THE INTERNET FIRST\n')
        exit()
    else:
        create_xlsx()

    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='input file (default videos.txt)', required=False)
    parser.add_argument('-O', '--output', help='output Excel filename', required=False)
    parser.add_argument('-y', '--youtube', help='youtube download', required=False, action='store_true')
    parser.add_argument('--sleep', type=int, default=30, help='seconds to sleep between downloads')
    parser.add_argument('--headless', action='store_true', help='disable GUI popups')
    parser.add_argument('--resolution', type=str, default="360p", help='video resolution to download (e.g., 360p, 720p)')

    args = parser.parse_args()

    if args.input:
        filename = args.input
    if args.output:
        spreadsheet = args.output

    if args.youtube:
        youtube(args)
    else:
        youtube(args)

    workbook.save(spreadsheet)
    return 0

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def check_keywords_in_captions(caption, keywords):
    """Check for keywords inside transcript text."""
    found_keywords = []
    for keyword in keywords:
        try:
            if re.search(rf"\b{re.escape(keyword)}\b", caption, re.IGNORECASE):
                found_keywords.append(keyword)
        except Exception:
            pass
    return ", ".join(found_keywords)

# Remove illegal characters for Excel
def clean_excel_value(val):
    if isinstance(val, str):
        return re.sub(r'[\x00-\x1F\x7F]', '', val)
    return val


def create_xlsx():
    global workbook, Sheet1
    workbook = Workbook()
    Sheet1 = workbook.active
    Sheet1.title = 'videos'
    Sheet1.freeze_panes = 'B2'

    # Excel column widths
    widths = {
        'A': 45, 'B': 45, 'C': 11, 'D': 9, 'E': 14,
        'F': 18, 'G': 6, 'H': 30, 'I': 12, 'J': 14,
        'K': 18, 'L': 15, 'M': 60, 'N': 2, 'O': 15
    }
    for col, w in widths.items():
        Sheet1.column_dimensions[col].width = w

    # Headers
    headers = [
        'url','title','description','views','author','publish_date',
        'length','download_name','video_id','thumbnail_url','dateDownloaded',
        'error','caption','(unused)','keywords'
    ]
    for idx, h in enumerate(headers, 1):
        Sheet1.cell(row=1, column=idx, value=h)

def finalMessage(headless=False):
    if headless:
        print(f"\nSaved to current folder\n\nSaved info to {spreadsheet}\n")
    else:
        try:
            window = Tk()
            window.geometry("1x1")
            w = Label(window, text='Youtube Downloader', font="100")
            w.pack()
            print(f'\nSaved to current folder\n\nSaved info to {spreadsheet}\n')
            messagebox.showinfo('information',
                f'\nDownloaded all Youtube videos from {filename}\n\nSaved to {spreadsheet}\n')
        except Exception:
            print(f"\nSaved to {spreadsheet} (GUI disabled)\n")

def internet(host="youtube.com", port=443, timeout=3):
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except socket.error as ex:
        print(ex)
        return False

def noInternetMsg():
    try:
        window = Tk()
        window.geometry("1x1")
        Label(window, text='Youtube Downloader', font="100").pack()
        messagebox.showwarning("Warning", "CONNECT TO THE INTERNET FIRST")
    except Exception:
        print("CONNECT TO THE INTERNET FIRST")

def noVideos():
    try:
        window = Tk()
        window.geometry("1x1")
        Label(window, text='Youtube Downloader', font="100").pack()
        messagebox.showwarning("Warning", f"Missing youtube videos in {filename}")
    except Exception:
        print(f"Missing youtube videos in {filename}")

def youtube_transcript(video_id):
    """Fetch transcript text, fallback if English not available."""
    caption = ""
    try:
        # Try English transcript
        srt = YouTubeTranscriptApi.get_transcript(video_id, languages=['en'])
    except (TranscriptsDisabled, NoTranscriptFound, CouldNotRetrieveTranscript):
        try:
            # Fallback: any available transcript
            transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)
            srt = transcript_list.find_transcript([t.language_code for t in transcript_list]).fetch()
        except Exception:
            return ""   # No transcripts available
    except Exception as e:
        print(f"Transcript error for {video_id}: {str(e)}")
        return ""

    for part in srt:
        text = part.get('text', '').strip()
        if text.lower() == '[music]':
            continue
        caption += " " + text

    return caption.strip()

# --------------------------------------------------------------------------
# Main downloader
# --------------------------------------------------------------------------
def youtube(args):
    keywords_to_check = ["tax", "lambo", "drug dealer"]

    if not os.path.exists(filename):
        print(f"\n{filename} does not exist\n")
        noVideos()
        exit()

    print(f'Downloading list of Youtube videos from {filename}. This can take a while...\n')
    csv_file = open(filename)

    count = 0

    for each_line in csv_file:
        link = each_line.split(',')[0].strip()
        if "youtu" not in link.lower():
            continue

        count += 1
        (title, description, views, author, publish_date) = ('', '', '', '', '')
        (length, rating, thumbnail_url, dateDownloaded, owner, caption) = ('', '', '', '', '', '')
        (owner_id, owner_url, download_name, error) = ('', '', '', '')
        (video_id, keywords) = ('', '')

        dateDownloaded = todaysDate

        try:
            yt = YouTube(link)
            video_id = yt.video_id
            title = yt.title
            download_name = f'{title}.mp4'
            author = yt.author
            description = yt.description
            rating = yt.rating
            publish_date = yt.publish_date
            length = str(yt.length)
            thumbnail_url = yt.thumbnail_url
            views = yt.views

            stream = yt.streams.filter(file_extension="mp4", res=args.resolution).first()
            if stream:
                stream.download()
            else:
                yt.streams.first().download()
            caption = youtube_transcript(video_id)

        except AgeRestrictedError:
            print(f"Video is age restricted: {link}")
            error = f"Age restricted (requires login). {download_name}"
        except Exception as e:
            print(f"An error occurred while processing video {link}: {str(e)}")
            error = f"Error processing video {download_name}: {str(e)}"
            title, description, views, author, publish_date, length, download_name, error, caption, tags = yt__dlp(link)   # test


        # Keyword search
        keywords = check_keywords_in_captions(caption, keywords_to_check)

        print(f'{link}  {title}')
        time.sleep(args.sleep)
        write_xlsx(link, title, description, views, author, publish_date, length,
                   download_name, rating, thumbnail_url, dateDownloaded, error,
                   caption, video_id, owner, owner_id, owner_url, keywords)

    if count == 0:
        noVideos()
    else:
        finalMessage(headless=args.headless)


def yt__dlp(video_url):
    title = description = views = creator = release_date = length = display_id = caption = tags = ''
    error = ''
    
    ydl_opts = {
        'format': 'mp4',
        'outtmpl': '%(title)s.%(ext)s',
        'quiet': True,
    }

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(video_url, download=True)

            # Extract desired metadata
            title = info.get('title')
            description = info.get('description')
            views = info.get('view_count')
            creator = info.get('uploader')
            release_date = info.get('upload_date')
            length = info.get('duration')
            display_id = info.get('display_id')
            
    except Exception as e:
        print(f"An error occurred while processing video {video_url}: {str(e)}")
        title = description = views = creator = release_date = length = display_id = caption = tags = ''
        error = str(e)

    return title, description, views, creator, release_date, length, display_id, error, caption, tags


# --------------------------------------------------------------------------
# Excel writer
# --------------------------------------------------------------------------
def write_xlsx(link, title, description, views, author, publish_date, length,
               download_name, rating, thumbnail_url, dateDownloaded, error,
               caption, video_id, owner, owner_id, owner_url, keywords):

    global Row, workbook, Sheet1, spreadsheet

    values = [
        link, title, description, str(views), author, str(publish_date),
        length, download_name, str(video_id), thumbnail_url, dateDownloaded,
        error, caption, '', keywords
    ]

    # for idx, val in enumerate(values, 1):
        # Sheet1.cell(row=Row, column=idx, value=val)

    for idx, val in enumerate(values, 1):
        try:
            Sheet1.cell(row=Row, column=idx, value=clean_excel_value(val))
        except IllegalCharacterError as e:
            print(f"Illegal character in column {idx}, row {Row}: {e}")
            Sheet1.cell(row=Row, column=idx, value="ERROR")


    Row += 1
    workbook.save(spreadsheet)  # Autosave after each row

# --------------------------------------------------------------------------
if __name__ == '__main__':
    main()



# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""
1.2.1 - re-written by chatGPT   
1.0.7 - works
1.0.5 - transcript saved as caption, look for keywords like tax
1.0.4 - functional copy
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>


"""
download a list of all urls the user has
parse comments

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>


"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
