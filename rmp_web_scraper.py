from datetime import datetime
import scrape_data
import filter_data
import requests
import time
import csv
import os

#place the url  for the professor you want to get data about
#URL= 'https://www.ratemyprofessors.com/professor/######' 
#type in chrome or firefox for whichever one you have
#browser = "chrome"

print("Enter the url of the professor:")
URL = input()
print()
print("What browser will you be using?")
print("Type chrome or firefox")
browser = input()
print()

Soup = scrape_data.scrape_data(browser, URL)

# Gets information about the professor
professor_name = filter_data.get_professor_name(Soup)
feedback_numbers = Soup.find_all('div',{'class': 'FeedbackItem__FeedbackNumber-uof32n-1 kkESWs'})
rating = Soup.find('div',{'class': 'RatingValue__Numerator-qw8sqy-2 liyUjw'}).get_text().strip()
takeAgain = feedback_numbers[0].text.strip()[:2]
Difficulty= feedback_numbers[1].text.strip()
scrapedate =  datetime.today()
scrapedate = scrapedate.strftime("%Y-%m-%d")
reviews = Soup.find_all('div',{'class': 'Rating__StyledRating-sc-1rhvpxz-1 jcIQzP'})

#creates the CSV for professor reviews with the following header
header = ['professor_name','TakeAgain%','Difficulty','Rating','Date','','course', 'Review Date', 'Quality', 'Difficulty', 'For Credit', 'Attendance', 'Would Take Again', 'Grade', 'Textbook', 'Online Class', 'Comment']
with open(f'{professor_name} Reviews.csv','w',newline='',encoding='UTF8') as f:
    writer = csv.writer(f)
    writer.writerow(header)

counter = 0
print("Creating csv...")
#reads review data, removes unnecessary tags, exports to CSV
for i in range(len(reviews)):
    #replaces previous value with NULL incase of no response in review
    review_for_credit = 'NULL'
    review_attendance = 'NULL'
    review_takeAgain  = 'NULL'
    review_grade      = 'NULL'
    review_textbook   = 'NULL'
    review_online     = 'NULL'
    review_comment    = 'NULL'
    reviews_process = reviews[i].get_text().strip()
    reviews_process = reviews_process.split('\n')

    new_list = filter_data.whitelist(reviews_process)
    
    # removes empty space in list
    new_list = filter_data.strip(new_list)
    
    # removes repeated course and date values in list
    new_list = filter_data.remove_duplicates(new_list)
    
    for i,j in enumerate(new_list):
        # ignores cases where coursename was not written properly
        if len(new_list[0]) > 9 or len(new_list[0]) < 6: 
            continue
        course = new_list[0]
        review_quality = new_list[3]
        review_difficulty = new_list[5]
        review_date = filter_data.convert_date(new_list[1])
        if 'For Credit' in j:
            review_for_credit = new_list[i+1]
        if 'Attendance' in j:
            review_attendance = new_list[i+1]
        if 'Would Take Again' in j:
            review_takeAgain  = new_list[i+1]
        if 'Grade' in j:
            review_grade      = new_list[i+1]
        if 'Textbook' in j:
            review_textbook   = new_list[i+1]
        if 'Online Class' in j:
            review_online     = new_list[i+1]
        review_comment    = new_list[i-1]
    
    review_dict = {'professor_name':professor_name,'takeAgain%' : takeAgain,'Total Difficulty' : Difficulty,'Total Rating' : rating,
    'Scrape Date' : scrapedate,'blank' : ' ','course': new_list[0], 'review_date': review_date,'review_quality': review_quality,
    'review_difficulty' : review_difficulty, 'review_for_credit':review_for_credit, 'review_attendance':review_attendance, 
    'review_takeAgain':review_takeAgain, 'review_grade':review_grade, 'review_textbook':review_textbook, 'review_online': review_online,'review_comment':review_comment}

    if counter == 1:
        review_dict = {'professor_name':'','takeAgain%' : '','Total Difficulty' : '','Total Rating' : '','Scrape Date' : '','blank' : ' ','course': course, 'review_date': review_date,'review_quality': review_quality,'review_difficulty' : review_difficulty, 'review_for_credit':review_for_credit, 
    'review_attendance':review_attendance, 'review_takeAgain':review_takeAgain, 'review_grade':review_grade, 'review_textbook':review_textbook, 'review_online': review_online,'review_comment':review_comment}
        

    #appends data to csv
    data = review_dict.values()
    with open(f'{professor_name} Reviews.csv','a+', newline='', encoding='UTF8') as f:
        writer = csv.writer(f)
        writer.writerow(data)
        counter = 1

print(f"{professor_name} Reviews.csv saved")