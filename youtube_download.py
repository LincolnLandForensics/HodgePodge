#!/usr/bin/python
# coding: utf-8

import sys
import os
import os.path
import socket
import argparse
from tkinter import *
from tkinter import messagebox
from pytube import YouTube
from datetime import date, datetime
from openpyxl.styles import Font 
from openpyxl import Workbook

d = datetime.today()
Day = d.strftime("%d")
Month = d.strftime("%m")
Year = d.strftime("%Y")
todaysDate = d.strftime("%m/%d/%Y")
todaysDateTime = d.strftime("%m_%d_%Y_%H-%M-%S")

author = 'LincolnLandForensics'
description = "Download a list of Youtube videos from videos.txt, save list in xlsx file"
version = '1.0.4'

def main():
    global filename
    filename = 'videos.txt'
    global Row
    Row = 1
    global spreadsheet
    spreadsheet = f'videos_{todaysDateTime}.xlsx'
    global sheet_format
    sheet_format = ''

    status = internet()
    if not status:
        noInternetMsg()
        print('\nCONNECT TO THE INTERNET FIRST\n')
        exit()
    else:
        create_xlsx()

    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-y', '--youtube', help='youtube download', required=False, action='store_true')

    args = parser.parse_args()

    if args.input:
        filename = args.input
    if args.output:
        outputFile = args.output

    if args.youtube:
        youtube()
    else:
        youtube()

    workbook.save(spreadsheet)
    return 0

def create_xlsx():
    global workbook
    workbook = Workbook()
    global Sheet1
    Sheet1 = workbook.active
    Sheet1.title = 'videos'
    # header_format = openpyxl.styles.Font(bold=True)
    # header_formatUrl = openpyxl.styles.Font(bold=True, color='FFc000')  # orange Case items

    Sheet1.freeze_panes = 'B2'
    # Sheet1['B2'].selected = True

    # Excel column width
    Sheet1.column_dimensions['A'].width = 45
    Sheet1.column_dimensions['B'].width = 45
    Sheet1.column_dimensions['C'].width = 50
    Sheet1.column_dimensions['D'].width = 11
    Sheet1.column_dimensions['E'].width = 12
    Sheet1.column_dimensions['F'].width = 19
    Sheet1.column_dimensions['G'].width = 7
    Sheet1.column_dimensions['H'].width = 30
    Sheet1.column_dimensions['I'].width = 8
    Sheet1.column_dimensions['J'].width = 40
    Sheet1.column_dimensions['K'].width = 15

    # Write column headers
    Sheet1['A1'] = 'url'
    Sheet1['B1'] = 'title'
    Sheet1['C1'] = 'description'
    Sheet1['D1'] = 'views'
    Sheet1['E1'] = 'author'
    Sheet1['F1'] = 'publish_date'
    Sheet1['G1'] = 'length'
    Sheet1['H1'] = 'download_name'
    Sheet1['I1'] = 'rating'
    Sheet1['J1'] = 'thumbnail_url'
    Sheet1['K1'] = 'dateDownloaded'

    # Sheet1['A1'].font = header_formatUrl
    # Sheet1['B1'].font = header_format
    # Sheet1['C1'].font = header_format
    # Sheet1['D1'].font = header_format
    # Sheet1['E1'].font = header_format
    # Sheet1['F1'].font = header_format
    # Sheet1['G1'].font = header_format
    # Sheet1['H1'].font = header_format
    # Sheet1['I1'].font = header_format
    # Sheet1['J1'].font = header_format
    # Sheet1['K1'].font = header_format

def finalMessage():
    window = Tk()
    window.geometry("1x1")
    w = Label(window, text='Youtube Downloader', font="100")
    w.pack()
    print('\nSaved to current folder\n\nSaved info to %s  \n' % (spreadsheet))
    messagebox.showinfo('information', '\nDownloaded all Youtube videos from %s  \n\nSaved to current folder\n\nSaved info to %s  \n' % (filename, spreadsheet))

def internet(host="youtube.com", port=443, timeout=3):
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except socket.error as ex:
        print(ex)
        return False

def noInternetMsg():
    window = Tk()
    window.geometry("1x1")
    w = Label(window, text='Youtube Downloader', font="100")
    w.pack()
    messagebox.showwarning("Warning", "CONNECT TO THE INTERNET FIRST")

def noVideos():
    window = Tk()
    window.geometry("1x1")
    w = Label(window, text='Youtube Downloader', font="100")
    w.pack()
    messagebox.showwarning("Warning", f"you are missing youtube videos in {filename}")

def youtube():
    if not os.path.exists(filename):
        print(f"\n\n\n\t {filename} does not exist\n\n\n\t")
        noVideos()
        exit()
    else:
        print('Downloading a list of Youtube videos from %s.\n\nThis can take a bit' % (filename))
        csv_file = open(filename)

    count = 0

    for each_line in csv_file:
        url, title, description, views, author, publish_date = '', '', '', '', '', ''
        length, rating, thumbnail_url, dateDownloaded, owner = '', '', '', '', ''
        owner_id, owner_url, captions, caption_tracks, note = '', '', '', '', ''

        each_line = each_line.split(',')
        each_line = each_line[0].strip()

        link = each_line

        if "youtu" in link.lower():
            count += 1
            yt = YouTube(link)
            title = yt.title
            note = f'{title}.mp4'
            author = yt.author
            description = yt.description
            rating = yt.rating
            publish_date = yt.publish_date
            length = str(yt.length)
            thumbnail_url = yt.thumbnail_url
            views = yt.views

            yt.streams.first().download()
            mp4_files = yt.streams.filter(file_extension="mp4")
            mp4_369p_files = mp4_files.get_by_resolution("360p")
            mp4_369p_files.download("")  # <Download folder path>

            write_xlsx(link, title, description, views, author, publish_date, length, rating, thumbnail_url,
                       dateDownloaded, owner, owner_id, owner_url, captions, caption_tracks, note)
        else:
            write_xlsx(link, title, description, views, author, publish_date, length, rating, thumbnail_url,
                       dateDownloaded, owner, owner_id, owner_url, captions, caption_tracks, note)

    if count == 0:
        noVideos()
        exit()
    else:
        finalMessage()

def write_xlsx(link, title, description, views, author, publish_date, length, rating, thumbnail_url,
               dateDownloaded, owner, owner_id, owner_url, captions, caption_tracks, note):
    global Row

    Sheet1.cell(row=Row, column=1, value=link)
    Sheet1.cell(row=Row, column=2, value=title)
    Sheet1.cell(row=Row, column=3, value=description)
    Sheet1.cell(row=Row, column=4, value=str(views))
    Sheet1.cell(row=Row, column=5, value=author)
    Sheet1.cell(row=Row, column=6, value=str(publish_date))
    Sheet1.cell(row=Row, column=7, value=length)
    Sheet1.cell(row=Row, column=8, value=note)
    Sheet1.cell(row=Row, column=9, value=str(rating))
    Sheet1.cell(row=Row, column=10, value=thumbnail_url)
    if title != '':
        Sheet1.cell(row=Row, column=11, value=todaysDate)

    Row += 1

def usage():
    file = sys.argv[0].split('\\')[-1]
    print("\nDescription: " + description)
    print(file + f" Version: {version} by {author}")
    print("\nExample:")
    print(f"\t{file} -y\t\t")
    print(f"\t{file} -y -I videos.txt")

if __name__ == '__main__':
    main()
