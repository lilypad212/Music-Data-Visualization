# Lily Wang
# 8/22/18
# Music data visualization pt 1: Creates an Excel sheet of song titles and artists of songs in your specific genre

import os
import eyed3 #external package for reading metadata of audio files
import pandas as pd
import xlsxwriter

eyed3.log.setLevel("ERROR") #hides warning messages


def show_info():
    audio = eyed3.load("E:\\Music\Power Up.mp3")
    print audio.tag.artist
    print audio.tag.album
    print audio.tag.title
    print audio.tag.genre

show_info()

title = []
artist = []
gender = []

dirname = "E:/Music"
for name in os.listdir(dirname):
    path = os.path.join(dirname, name) # Path name for subfolders
    if path.endswith('mp3'):
        song = eyed3.load(path)
        if str(song.tag.genre) == ("K-Pop"):
            if song.tag.artist not in artist:
                #artistsList.write(song.tag.artist)
            title.append(song.tag.title)
            artist.append(song.tag.artist)

out_path = "C:/Users/Lily/Desktop/CS Programs/Data.xlsx"
df = pd.DataFrame({'Artist': artist, 'Song Title': title})
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()