import urllib.request, urllib.parse, urllib.error
from bs4 import BeautifulSoup
import openpyxl

sheet_file = openpyxl.load_workbook("itunes_charts.xlsx")
sheet = sheet_file["Sheet1"]

link_dict = {
    "USA" : "https://rss.itunes.apple.com/api/v1/us/apple-music/top-songs/all/100/explicit.atom?at=10l9W2",
    "United Kingdom" : "https://rss.itunes.apple.com/api/v1/gb/apple-music/top-songs/all/100/explicit.atom?at=10l9W2",
    "Ukraine" : "https://rss.itunes.apple.com/api/v1/ua/apple-music/hot-tracks/all/100/explicit.atom?at=10l9W2",
    "France" : "https://rss.itunes.apple.com/api/v1/fr/apple-music/hot-tracks/all/100/explicit.atom?at=10l9W2",
    "Germany" : "https://rss.itunes.apple.com/api/v1/de/apple-music/hot-tracks/all/100/explicit.atom?at=10l9W2"
    }


artist_count = 1
song_count = 1
category_count = 1
rel_count = 1
album_count = 1


for country in link_dict:
    #print(country, link_dict[country])

    chart = urllib.request.urlopen(link_dict[country]).read()
    soup = BeautifulSoup(chart, "xml")

    artist_col = "A"
    
    pos_col = "G"
    position = 0

    date_col = "H"
    date_time = soup.find("updated")
    pos_date = date_time.get_text().find("T")
    date = date_time.get_text()[ : pos_date]

    country_col = "E"

    artists = soup.find_all("artist")
    for artist in artists:
        artist_count = artist_count + 1
        cell_number = artist_col + str(artist_count)
        sheet[cell_number] = artist.get_text()
        position = position + 1
        pos_cell_number = pos_col + str(artist_count)
        sheet[pos_cell_number] = position
        date_cell_number = date_col + str(artist_count)
        sheet[date_cell_number] = date
        country_cell_number = country_col + str(artist_count)
        sheet[country_cell_number] = country


    song_col = "B"
    songs = soup.find_all("name")
    for song in songs[1:]:
        song_count = song_count + 1
        cell_number = song_col + str(song_count)
        sheet[cell_number] = song.get_text()

    category_col = "D"
    categorys = soup.find_all("entry")
    for category in categorys:
        cat = category.find_all("category")
        for style in cat:
            if int(style.attrs["im:id"]) > 33:
                continue
            else:
                category_count = category_count + 1
                cell_number = category_col + str(category_count)
                sheet[cell_number] = style.attrs["term"]
            break    


    rel_col = "F"
    releases = soup.find_all("releaseDate")
    for rel in releases:
        rel_count = rel_count + 1
        cell_number = rel_col + str(rel_count)
        sheet[cell_number] = rel.get_text()


    album_col = "C"
    albums = soup.find_all("content")
    for content in albums:
        album_count = album_count + 1
        cell_number = album_col + str(album_count)
        lpos_alb = content.get_text().find("10l9W2>")
        rpos_alb = content.get_text().find("</a><br/")
        alb = content.get_text()[lpos_alb + 7 : rpos_alb]
        sheet[cell_number] = alb
    

sheet_file.save("itunes_charts.xlsx")

print("Done")