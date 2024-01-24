# Rate My Professors Web Scraper 
I started this project with the motivation to learn web scraping. The website https://www.ratemyprofessors.com is a helpful website that allows you to view a professors rating given to them by students who have left a review along with a quality and difficulty score. 

This app allows you to collect all of the reviews given to a professor and neatly export them to a csv file. From this cvs file you can look at all the collective data and make an informed decision on whether you want to take their class in a following semester.

This app can also be used like I have to give yourself a dataset that I can use to practice data sorting in excel. (see last image)

## Setup

1) This was written in Python 3.12.1. You will need to install it from here: https://www.python.org/downloads/release/python-3121/

2) This uses the Selenium module to automate web browser interactions. Currently I have only implemented chrome and firefox browsers. Make sure you have one of those installed:
    - https://www.google.com/chrome/
    - https://www.mozilla.org/en-US/firefox/new/

2) Go to the directory of the python files and run "pip install -r requirements.txt" in command prompt to install all modules.

## App Usage

1) Open the rmp_web_scraper.py file and provide the url and browser.

4) The csv will be saved as the professors "Firstname Lastname Reviews.csv" in the current directory.

## CSV Output

![csv example](/assets/images/csv_example.jpg?raw=true "csv example")


## Data Analysis Example
(Not Provided)

![analysis example](/assets/images/analysis_example.png?raw=true "analysis example")