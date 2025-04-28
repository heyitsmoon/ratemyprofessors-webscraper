from datetime import datetime

def get_professor_name(Soup):
    title = Soup.title.string
    title = title.strip()
    title = title.split(' ')
    professor = ''
    for i in title:
        if i == "at":
            break
        professor += i + " "
    return professor.strip()


def convert_date(date_string):
    day = date_string.split(' ')[1][:-3]
    month_abbr = date_string.split(' ')[0][:3]
    year = date_string.split(' ')[2]
    date_str = f"{month_abbr}-{day}-{year}"
    date_obj = datetime.strptime(date_str, '%b-%d-%Y')
    date_obj = date_obj.strftime("%Y-%m-%d")
    return date_obj

def getYear(date_string):
    day = date_string.split(' ')[1][:-3]
    month_abbr = date_string.split(' ')[0][:3]
    year = date_string.split(' ')[2]
    date_str = f"{month_abbr}-{day}-{year}"
    date_obj = datetime.strptime(date_str, '%b-%d-%Y')
    date_obj = date_obj.strftime("%Y")
    return date_obj

def whitelist(reviews_process):
    # Use list comprehension to filter out undesired items
    filtered_list = [item for item in reviews_process if item.strip() != '' and 
        item.strip() not in ['😎','😖','😐','awesome','awful','average',':',
        'Participation matters','Group projects','GROUP PROJECTS','PROJECTS','CARES ABOUT STUDENTS',
        'So many papers','Amazing lectures','Caring','Inspirational','Respected',
        'Clear grading criteria','Hilarious',"Skip class? You won't pass.",'Gives good feedback',
        'Graded by few things','GRADED BY FEW THINGS','Accessible outside class',
        'EXTRA CREDIT', 'Online Savvy', 'LECTURE HEAVY','ACCESSIBLE OUTSIDE CLASS', 'Test heavy',
        'Get ready to read','Would take again','TEST HEAVY','PARTICIPATION MATTERS',
        'LOTS OF HOMEWORK','Lots of homework','Tough grader','Tough Grader',
        'Lecture heavy','BEWARE OF POP QUIZZES','Reviewed'] and 
        'Reviewed: ' not in item.strip()]
    
    return filtered_list
    
def remove_duplicates(new_list):
    deduplicated_list = [item for i, item in enumerate(new_list) if i != 6 and i != 7]
    #print(deduplicated_list)
    return deduplicated_list
        
def strip(new_list):
    for i,j in enumerate(new_list):
        new_list[i] = j.strip()

    new_list = [item for item in new_list if item.strip()]
    return new_list 