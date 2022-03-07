from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
import xlsxwriter

#find the specified occurence of a substring from a string
def findnth(word, subword, n):
    parts = word.split(subword, n + 1)
    if len(parts) <= n + 1:
        return -1
    return len(word) - len(parts[-1]) - len(subword)

#store bb top 10 ranks in rankings list
def get_rankings(html):
    rankings_index = []
    for k in range(10):
        rankings_index.insert(k, (findnth(html, "ye-chart-item__rank", k) + 24)) 
        rankings_index[k] = int(rankings_index[k])
    rankings = []
    #retrieve ranks and format properly in list 
    for k in range(10):
        rankings.insert(k, html[rankings_index[k] - 2:rankings_index[k]])        
        rankings[k] = rankings[k].rstrip()
    return rankings

#store bb top 10 artists in artists list
def get_artists(html):
    artists_whole_index = []
    artists_end_index = []
    for k in range(10):
        artists_end_index.insert(k, (findnth(html, "ye-chart-item__expand-caret", k) - 32))
        artists_end_index[k] = int(artists_end_index[k])
        artists_whole_index.insert(k, (findnth(html, "ye-chart-item__artist", k))) 
        artists_whole_index[k] = int(artists_whole_index[k])
    artists = []
    #retrieve artists and format properly in list 
    for k in range(10):
        artists.insert(k, html[artists_whole_index[k]:artists_end_index[k]])        
        escape_char = artists[k].find("\n")        
        while escape_char > 0:
            escape_char = artists[k].find("\n") + 1
            artists[k] = artists[k][escape_char:]            
    return artists

#store bb top 10 songs in songs list
def get_songs(html):
    songs_whole_index = []
    songs_end_index = []
    for k in range(10):
        songs_end_index.insert(k, (findnth(html, "ye-chart-item__artist", k) - 20))
        songs_end_index[k] = int(songs_end_index[k])
        songs_whole_index.insert(k, (findnth(html, "ye-chart-item__title", k))) 
        songs_whole_index[k] = int(songs_whole_index[k])
    songs = []
    #retrieve songs and format properly in list 
    for k in range(10):
        songs.insert(k, html[songs_whole_index[k]:songs_end_index[k]])        
        
        escape_char = songs[k].find("\n")
        
        while escape_char > 0:
            escape_char = songs[k].find("\n") + 1
            songs[k] = songs[k][escape_char:]            
    return songs

#manual musical data input --> retrieve and process

def get_tonic(html):
    songs_whole_index = []
    songs_end_index = []
    for k in range(10):
        songs_end_index.insert(k, (findnth(html, "ye-chart-item__artist", k) - 20))
        songs_end_index[k] = int(songs_end_index[k])
        songs_whole_index.insert(k, (findnth(html, "ye-chart-item__title", k))) 
        songs_whole_index[k] = int(songs_whole_index[k])
    songs = []
    #retrieve songs and format properly in list 
    for k in range(10):
        songs.insert(k, html[songs_whole_index[k]:songs_end_index[k]])        
        
        escape_char = songs[k].find("\n")
        
        while escape_char > 0:
            escape_char = songs[k].find("\n") + 1
            songs[k] = songs[k][escape_char:]            
    return songs


def write_data_xlsx(spreadsheet, dictionary):
    row = 0
    col = 0

    order=sorted(dictionary.keys())
    for key in order:
        row += 1
        spreadsheet.write(row, col, key)
        i = 1
        for item in dictionary[key]:
            spreadsheet.write(row, col + i, item)
            i += 1

#obtain data from billboard top 100 
year_num = [2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020]
url = []
workbook = xlsxwriter.Workbook('billboard_top_10_data.xlsx')

#write data to spreadsheet for each year
for x in range(len(year_num)):
    url.insert(x, 'https://www.billboard.com/charts/year-end/' + str(year_num[x]) + '/hot-100-songs')

    bb_website = Request(url[x], headers={'User-Agent': 'Mozilla/5.0'})
    bb_website_html = urlopen(bb_website).read()

    #html parse & obtain rankings under main tag
    soup = BeautifulSoup(bb_website_html, 'html.parser')
    html_obj = soup.find("main")
    html = str(html_obj)

    #test functions and data grabbing 
    html_txt_doc = open("bb2.txt", "w")
    rankings_list = get_rankings(html)
    html_txt_doc.write(str(rankings_list))
    html_txt_doc.close()
    html_txt_doc = open("bb2.txt", "a")
    artist_list = get_artists(html)
    songs_list = get_songs(html)
    html_txt_doc.write("\n" + str(artist_list))
    html_txt_doc.write("\n" + str(songs_list))
    html_txt_doc.write("\n\n")

    #record data in dictionary --> xlsx 
    bb_top_10_dict = {
        "Rank" : rankings_list,
        "Song Title" : songs_list,
        "Artist" : artist_list#,
       #"Chords" : chord_list,
       #"Tempo" : tempo_list,
       #"Tonic" : tonic_list,
       #"Key" : key_list
    }
   
    worksheet = workbook.add_worksheet(str(year_num[x]) + " Charts")
    write_data_xlsx(worksheet, bb_top_10_dict)
workbook.close()