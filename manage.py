from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name'])


try:
    source = requests.get('https://www.imdb.com/chart/boxoffice')
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text,'html.parser')
    
    movies = soup.find('td', class_="titleColumn").find_all('a')
    
    for movie in movies:
        name = movie.find('td',class_ = "titleColumn").a.text
        
        rank = movie.find('td',class_="ratingColumn").get_text(Strip=True).split()
        
       # year = movie.find('td',class_ = "titleColumn").span.text.strip('()')
        
       # rating = movie.find('td',class_ = "ratingCOlumn imdbRating")
        
        print(rank,name)
       # sheet.append([rank,name,year,rating])
    
    
except Exception as e:
    print(e)

#excel.save('IMDB Movie Ratings.xlsl')

