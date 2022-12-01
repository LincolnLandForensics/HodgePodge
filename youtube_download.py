#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import sys
import os.path
import socket
import argparse  # for menu system
from tkinter import * 
from tkinter import messagebox
try:
    from pytube import YouTube  # pip install pytube
    import xlsxwriter
except:
    print('pip install -r requirements_youtube.txt')

from datetime import date
from datetime import datetime

d = datetime.today()

Day    = d.strftime("%d")
Month = d.strftime("%m")    # %B = October
Year  = d.strftime("%Y")        
todaysDate = d.strftime("%m/%d/%Y")
todaysDateTime = d.strftime("%m_%d_%Y_%H-%M-%S")

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Download a list of Youtube videos from videos.txt, save list in xlsx file"
version = '1.0.3'


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():

    # global variables
    global filename
    filename = ('videos.txt')
    global Row
    Row = 1  # defines arguments
    global spreadsheet
    spreadsheet = ('videos_%s.xlsx' %(todaysDateTime)) # uniq naming for xlsx output
    global sheet_format
    sheet_format = ('')

    # check internet status
    status = internet()
    if status == False:
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

    if args.input:  # defaults to index.txt
        filename = args.input
    if args.output: # defaults to output.txt
        outputFile = args.output

    if args.youtube:
        # print('reading from %s\n\n' %(filename))
        youtube()
    else:
        # print('Reading from %s\n\n' %(filename))
        youtube()

    workbook.close()
    return 0    

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def create_xlsx():  
    '''
    Creates an xlsx file with formatting
    '''
  
    global workbook
    workbook = xlsxwriter.Workbook(spreadsheet)
    global Sheet1
    Sheet1 = workbook.add_worksheet('videos')
    header_format = workbook.add_format({'bold': True, 'border': 1})
    header_formatUrl = workbook.add_format({'bold': True, 'border': 1, 'bg_color':'#FFc000'})   # orange Case items

    Sheet1.freeze_panes(1, 1)  # Freeze cells
    Sheet1.set_selection('B2')

    # Excel column width
    Sheet1.set_column(0, 0, 45) # url
    Sheet1.set_column(1, 1, 45) # title
    Sheet1.set_column(2, 2, 50) # description
    Sheet1.set_column(3, 3, 11) # views
    Sheet1.set_column(4, 4, 12) # author
    Sheet1.set_column(5, 5, 19) # publish_date
    Sheet1.set_column(6, 6, 7) # length
    Sheet1.set_column(7, 7, 30) # download_name
    Sheet1.set_column(8, 8, 8) # rating
    Sheet1.set_column(9, 9, 40) # thumbnail_url
    Sheet1.set_column(10, 10, 15) # dateDownloaded
    # Sheet1.set_column(11, 11, 15) # owner
    # Sheet1.set_column(12, 12, 15) # owner_id
    # Sheet1.set_column(13, 13, 15) # owner_url
    # Sheet1.set_column(14, 14, 15) # captions
    # Sheet1.set_column(15, 15, 15) # caption_tracks

    # hidden columns
    Sheet1.set_column(15, 15, None, None, {'hidden': 1}) # caption_tracks
    
    # Write column headers
    Sheet1.write(0, 0, 'url', header_formatUrl)
    Sheet1.write(0, 1, 'title', header_format)
    Sheet1.write(0, 2, 'description', header_format)
    Sheet1.write(0, 3, 'views', header_format)
    Sheet1.write(0, 4, 'author', header_format)
    Sheet1.write(0, 5, 'publish_date', header_format)
    Sheet1.write(0, 6, 'length', header_format)
    Sheet1.write(0, 7, 'download_name', header_format)
    Sheet1.write(0, 8, 'rating', header_format)
    Sheet1.write(0, 9, 'thumbnail_url', header_format)
    Sheet1.write(0, 10, 'dateDownloaded', header_format)
    # Sheet1.write(0, 11, 'owner', header_format)
    # Sheet1.write(0, 12, 'owner_id', header_format)
    # Sheet1.write(0, 13, 'owner_url', header_format)
    # Sheet1.write(0, 14, 'captions', header_format)
    # Sheet1.write(0, 15, 'caption_tracks', header_format)

def FormatFunction(bg_color = 'white'):
	global Format
	Format=workbook.add_format({
	'bg_color' : bg_color
	}) 

def finalMessage():
    '''
    prints a pop-up that says "done downloading videos from videos.txt to *.xlsx"
    '''
    window = Tk()
    window.geometry("1x1")
      
    w = Label(window, text ='Youtube Downloader', font = "100") 
    w.pack()
    print('\nSaved to current folder\n\nSaved info to %s  \n' %(spreadsheet))
    messagebox.showinfo('information', '\nDownloaded all Youtube videos from %s  \n\nSaved to current folder\n\nSaved info to %s  \n' %(filename, spreadsheet))
    
    
def internet(host="youtube.com", port=443, timeout=3):
    """
    Host: youtube.com
    OpenPort: 443/tcp
    Service: domain (DNS/TCP)
    """
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except socket.error as ex:
        print(ex)
        return False

def noInternetMsg():
    '''
    prints a pop-up that says "CONNECT TO THE INTERNET FIRST"
    '''
    window = Tk()
    window.geometry("1x1")
      
    w = Label(window, text ='Youtube Downloader', font = "100") 
    w.pack()
    messagebox.showwarning("Warning", "CONNECT TO THE INTERNET FIRST") 

def noVideos():
    '''
    prints a pop-up that says "you are missing youtube videos"
    '''
    window = Tk()
    window.geometry("1x1")
      
    w = Label(window, text ='Youtube Downloader', font = "100") 
    w.pack()
    messagebox.showwarning("Warning", ("you are missing youtube videos in %s" %(filename)))
    
def youtube():
    '''
    read youtube url's, download them and create an xlsx of details
    of the videos
    '''

    print('Downloading a list of Youtube videos from %s.\n\nThis can take a bit'%(filename))
    try:
        csv_file = open(filename)
    except:
        print("\n\n\n\t", filename, " does not exist\n\n\n\t")
        noVideos()
        exit()
    (count) = (0)

    for each_line in csv_file:
        (url, title, description, views, author, publish_date) = ('', '', '', '', '', '')
        (length, rating, thumbnail_url , dateDownloaded, owner) = ('', '', '', '', '')
        (owner_id, owner_url, captions, caption_tracks, note) = ('', '', '', '', '')

        
        each_line = each_line.split(',')  # splits by comas
        each_line = (each_line[0].strip())
        
        link = each_line

        # check if it's a youtube url
        if "youtu" in link.lower():
            count +=1
            yt = YouTube(link)
            title = yt.title
            note = ('%s.mp4' %(title))
            author = yt.author
            description = yt.description
            # captions = yt.captions    # list?
            # caption_tracks = yt.caption_tracks  # list not compatible with xlsx
            # metadata = yt.metadata
            rating = yt.rating
            publish_date = yt.publish_date
            # check_availability = yt.check_availability
            # owner = yt.owner
            # owner_id = yt.owner_id
            # owner_url = yt.owner_url
            # playlist_url = yt.playlist_url
            # video_urls = yt.video_urls
            # videos = yt.videos
            length = str(yt.length) # 
            thumbnail_url = yt.thumbnail_url
            views = yt.views
            
            yt.streams.first().download()   # Get the first Stream in the results.  # Write the media stream to disk.
            mp4_files = yt.streams.filter(file_extension="mp4")
            mp4_369p_files = mp4_files.get_by_resolution("360p")
            mp4_369p_files.download("") # <Download folder path>

        # Write excel
            write_xlsx(link, title, description, views, author, publish_date, length, rating, thumbnail_url , dateDownloaded, owner, owner_id, owner_url, captions, caption_tracks, note)
        else:
            write_xlsx(link, title, description, views, author, publish_date, length, rating, thumbnail_url , dateDownloaded, owner, owner_id, owner_url, captions, caption_tracks, note)
      
    if count == 0:
        noVideos()
        exit()
    else:
        finalMessage()


def write_xlsx(link, title, description, views, author, publish_date, length, rating, thumbnail_url , dateDownloaded, owner, owner_id, owner_url, captions, caption_tracks, note):
    '''
    write out_log_.xlsx
    '''
    global Row

    Sheet1.write_string(Row, 0 , link) 
    Sheet1.write_string(Row, 1 , title) 
    Sheet1.write_string(Row, 2 , description) 
    Sheet1.write_string(Row, 3 , str(views)) # errors was integer
    Sheet1.write_string(Row, 4 , author) 
    Sheet1.write_string(Row, 5 , str(publish_date))
    Sheet1.write_string(Row, 6 , length)    
    Sheet1.write_string(Row, 7 , note) 
    Sheet1.write_string(Row, 8 , str(rating))  
    Sheet1.write_string(Row, 9 , thumbnail_url)
    if title != '':
        Sheet1.write_string(Row, 10 , todaysDate) # dateDownloaded

    # Sheet1.write_string(Row, 11 , owner)
    # Sheet1.write_string(Row, 12 , owner_id)
    # Sheet1.write_string(Row, 13 , owner_url)
    # Sheet1.write_string(Row, 14 , captions)
    # Sheet1.write_string(Row, 15 , caption_tracks)
     
    Row += 1


def usage():
    file = sys.argv[0].split('\\')[-1]
    print("\nDescription: " + description)
    print(file + " Version: %s by %s" % (version, author))
    print("\nExample:")
    print("\t" + file + " -y\t\t")
    print("\t" + file + " -y -I videos.txt")

  
if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
1.0.1 - export details to videos_<datestamp>.xlsx
1.0.0 - gui error messages. works without any switches
0.9.0 - you can specify a different input file with -I (ex. -I input.txt)
0.8.0 - internet check. It shows an error window if you aren't connected
0.1.0 - read vidoes.txt, downloud youtube videos and output to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
skip the bonus 1x1 tkinter window
don't download the .3gpp
list all videos for a username, export to .xlsx

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
insert a list of youtube videos into videos.txt, it will save them to the same folder.

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2022 LincolnLandForensics
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License version 2, as published by the
# Free Software Foundation
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details (http://www.gnu.org/licenses/gpl.txt).

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
