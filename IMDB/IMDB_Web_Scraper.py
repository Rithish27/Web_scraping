# Import the required modules for requesting web data and parsing.
import requests,openpyxl
from bs4 import BeautifulSoup

# requests --> Used to make http calls to request web contents.
# BeautifulSoup --> Parsing the response from web content.
# openpxl --> To load data into excel.

excel = openpyxl.Workbook() # Creating a excel workbook
sheet = excel.active # Making the excel sheet active to allow adding values/ modifications
sheet.title = 'Top Rated Movies' # Naming the excel sheet 
print(excel.sheetnames) # verify the changes made


# To add headings of values in Excel we need to pass it in a list as below
sheet.append(['Movie Rank','Movie Name','Year of Release','Movie Runtime','IMDB Rating','Total Ratings'])


# resource location
base_url = 'https://www.imdb.com/chart/top/'

try:

    # This a optional params that needed to pass when we encountered 403 Forbidden Error
    headers = {'User-Agent':
               'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36'
               ,'Referer':'https://www.imdb.com/'}
    

    # requesting data from url
    result = requests.get(base_url,headers=headers)

    # Using BeautifulSoup to parsethe web contents
    soup = BeautifulSoup(result.content.decode(),'html.parser')

    # Identify the exact tag/class name in HTML, where does the data that we required is located
    movies = soup.find('ul',class_ = "ipc-metadata-list ipc-metadata-list--dividers-between sc-a1e81754-0 eBRbsI compact-list-view ipc-metadata-list--base").find_all('li',class_='ipc-metadata-list-summary-item sc-10233bc-0 iherUv cli-parent')

    for movie in movies:
        # find is a method in bs4 to get the first element with the given tag name or class
        movie_name = movie.find('h3',class_='ipc-title__text').text.split('.')

        name = movie_name[1]

        rank = movie_name[0]

        movie_details = movie.find('div',class_='sc-b0691f29-7 hrgukm cli-title-metadata').text

        release_year = movie_details[:4]
        run_time = movie_details[4:-1]

        movie_rating = movie.find('span',class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').text.split()

        imdb_rating = movie_rating[0]
        total_ratings_given = movie_rating[1].strip('()')

        print(rank,name,release_year,run_time,imdb_rating,total_ratings_given)

        # loading data row by row into excel file
        sheet.append([rank,name,release_year,run_time,imdb_rating,total_ratings_given])
        
        
except Exception as e:
    print(e)
    
# save the Excel file by passing a name
excel.save('IMDB Movie Ratings.xlsx')

print('Done')