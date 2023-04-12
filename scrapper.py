import requests, openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'IMDB top 250 movies'
sheet.append(["Movie Rank", "Movie Name", "Year of Release", "IMDB Rating"])

try:
    imdb = requests.get("https://www.imdb.com/chart/top/")
    imdb.raise_for_status()

    soup = BeautifulSoup(imdb.text, 'html.parser')
    movies = soup.find('tbody', class_='lister-list').find_all('tr')

    for movie in movies:
        name = movie.find('td', class_='titleColumn').a.text
        rank = movie.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]
        year = movie.find('td', class_='titleColumn').span.text.strip('()')
        rating = movie.find('td', class_='ratingColumn imdbRating').strong.text
        sheet.append([rank, name, year, rating])

except Exception as e:
    print("page not found", e)

excel.save("IMDB Movie Ratings")
