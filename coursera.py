import sys
import os
import random
import requests
import openpyxl
from openpyxl.styles import Font
from lxml import html
from bs4 import BeautifulSoup


def get_courses_list():
    response = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    content = response.text.encode('utf-8')
    parsed_document = html.fromstring(content)
    courses_list = parsed_document.xpath('//loc/text()')
    return courses_list


def get_course_info(course_slug):
    response = requests.get(course_slug)
    content = response.content.decode('utf-8')
    soup = BeautifulSoup(content, 'html.parser')
    title = soup.find('h1', {'class': 'title display-3-text'}).text
    lang = soup.find('div', {'class': 'rc-Language'}).text
    startdate = soup.find('div', {
        'class': 'startdate rc-StartDateString caption-text'
    }).text
    weeks = len(soup.findAll('div', {'class': 'week'}))
    rating_tags = soup.findAll('div', {'class': 'ratings-text headline-2-text'})
    rating = ''
    for rating_tag in rating_tags:
        rating += rating_tag.text
    course_info = [title, lang, startdate, weeks, rating]
    return course_info


def output_courses_info_to_xlsx(filepath, courses_info_list):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    title_list = ['TITLE', 'LANGUAGE', 'DATE TO START', 'WEEKS', 'RATING']
    worksheet.append(title_list)
    for title in [
        worksheet['A1'],
        worksheet['B1'],
        worksheet['C1'],
        worksheet['D1'],
        worksheet['E1']
    ]:
        title.font = Font(bold=True)
    for course_info_list in courses_info_list:
        worksheet.append(course_info_list)
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column].width = length
    workbook.save(filepath)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        exit('need path to file as argument')
    dir_path = sys.argv[1]
    file_name = 'coursera_courses.xlsx'
    if not os.path.isdir(dir_path):
        exit('incorrect directory to save')
    file_path = os.path.join(dir_path, file_name)
    courses = get_courses_list()
    max_length = 20
    courses_info_list = []
    try:
        while len(courses_info_list) < max_length:
            index = random.randint(0, len(courses))
            courses_info_list.append(get_course_info(courses[index]))
    except(IndexError, AttributeError):
        pass
    output_courses_info_to_xlsx(file_path, courses_info_list)
