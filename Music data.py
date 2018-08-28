# Lily Wang
# 8/22/18
# Music data visualization pt 1: Creates an Excel sheet of song titles and artists of songs in a specific genre
# as well as a text file of a list of unique artists

import os
import eyed3 #external package for reading metadata of audio files
import pandas as pd
import xlsxwriter

eyed3.log.setLevel("ERROR") #hides warning messages


def show_info(): #initial test
    audio = eyed3.load("E:\\Music\Power Up.mp3")
    print audio.tag.artist
    print audio.tag.album
    print audio.tag.title
    print audio.tag.genre

show_info()

title = []
artist = []
#year = []
gender = []
dirname = "E:/Music"
artistsList = open("C:\Users\Lily\Desktop\CS Programs\List of Artists.txt","w") #List of all unique artists in collection

def add_title_artist():
    for name in os.listdir(dirname):
        path = os.path.join(dirname, name) # Path name for subfolders
        if os.path.exists(path): 
            if path.endswith('mp3'): #Find all mp3 files
                song = eyed3.load(path)
                if str(song.tag.genre) == ("K-Pop"): #Only adding K-Pop songs
                    if song.tag.artist not in artist:
                        artistsList.write(song.tag.artist.encode('utf-8')) #Makes a list of unique artists
                        artistsList.write('\n')
                    #print song.tag.title
                    title.append(song.tag.title)
                    artist.append(song.tag.artist)
                    #year.append(song.tag.getBestDate())

add_title_artist()

#Defines each song as by a "male", "female", or "other" artist. Other means both female and male (i.e. duet or co-ed group)
def add_gender():
    for name in artist:
        if name.encode('utf-8').upper() in open('C:\Users\Lily\Desktop\CS Programs\Female Artists.txt').read():
            gender.append("Female")
        elif name.encode('utf-8').upper() in open('C:\Users\Lily\Desktop\CS Programs\Male Artists.txt').read():
            gender.append("Male")
        else:
            gender.append("Other")

add_gender()

out_path = "C:/Users/Lily/Desktop/CS Programs/Data.xlsx"
df = pd.DataFrame({'Song Title': title, 'Artist': artist, 'Gender': gender})
df = df[['Song Title','Artist','Gender']]
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
artistsList.close()