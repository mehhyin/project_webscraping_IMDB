from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
#see how many excel sheets it has
print(excel.sheetnames)
#make sure we are working on active sheet
sheet = excel.active
sheet.title = 'Top Movies Based on IMDB'
print(excel.sheetnames)

#create headings for the excel
sheet.append(['Movie rank','Movie name','Year of release','IMDB rating'])



try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status() #<=== it will print an error if the website is wrong or have issue

    soup = BeautifulSoup(source.text, 'html.parser')
    #print(soup)

    #now we need to get the rank, title and etc from the website
    #Since <tr> tag is inside <tbody>, we need to access to <tbody> first 
    movies = soup.find('tbody', class_="lister-list").find_all('tr')       #use find_all() to get all <tr> tags
    #print(len(movies))

    #Need to access to <td> tags now
    #so now we will open a loop to loop through all the td tags
    for x in movies:
        name = x.find('td', class_="titleColumn").a.text
        rank = x.find('td', class_="titleColumn").get_text(strip=True).split('.')[0] #need to add index as it will return all in list
        year = x.find('td', class_="titleColumn").span.text.strip('()')
        rating = x.find('td', class_="ratingColumn imdbRating").strong.text
        print(rank, name, year, rating)
        #everytime the scraper extract value, we want it to load it into the excel
        sheet.append([rank, name, year, rating])


except Exception as e:
    print(e) 


#save the excel 
excel.save('IMDB Movie Ratings.xlsx')