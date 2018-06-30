import sys
import os
import random
import requests
import openpyxl
from openpyxl.styles import Font
from lxml import html
from bs4 import BeautifulSoup


def get_data_from_url(url):
    response = requests.get(url)
    content = response.content.decode('utf-8')
    return content


def get_courses_urls_list(xml_data):
    parsed_urls_document = html.fromstring(xml_data)
    courses_urls_list = parsed_urls_document.xpath('//loc/text()')
    return courses_urls_list


def get_course_info(page):
    title = page.find('h1', {'class': 'title display-3-text'}).text
    lang = page.find('div', {'class': 'rc-Language'}).text
    startdate = page.find('div', {
        'class': 'startdate rc-StartDateString caption-text'
    }).text
    weeks = len(page.findAll('div', {'class': 'week'}))
    rating_tags = page.findAll('div', {'class': 'ratings-text headline-2-text'})
    rating = ''
    for rating_tag in rating_tags:
        rating += rating_tag.text
    course_info = {
        'title': title,
        'language': lang,
        'startdate': startdate,
        'weeks': weeks,
        'rating': rating
    }
    return course_info


def output_courses_info_to_xlsx(courses_info_list):
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
    for course_info_dic in courses_info_list:
        course_info = [
            course_info_dic['title'],
            course_info_dic['language'],
            course_info_dic['startdate'],
            course_info_dic['weeks'],
            course_info_dic['rating']
        ]
        worksheet.append(course_info)
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column].width = length
    return workbook


def save_data_to_xlsx(workbook, filepath):
    workbook.save(filepath)


if __name__ == '__main__':
    if len(sys.argv) == 1 or not os.path.isdir(sys.argv[1]):
        exit('need path to file as argument / incorrect directory to save')
    dir_path = sys.argv[1]
    file_name = 'coursera_courses.xlsx'
    file_path = os.path.join(dir_path, file_name)
    courses_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    courses_data_from_url = get_data_from_url(courses_url).encode('utf-8')
    courses_urls_list = get_courses_urls_list(courses_data_from_url)
    number_of_courses = 20
    courses_info_list = []
    try:
        short_urls_list = random.sample(courses_urls_list, number_of_courses)
        for course_url in short_urls_list:
            course_data = get_data_from_url(course_url)
            soup = BeautifulSoup(course_data, 'html.parser')
            course_info = get_course_info(soup)
            courses_info_list.append(course_info)
    except(IndexError, AttributeError):
        pass
    excel_workbook = output_courses_info_to_xlsx(courses_info_list)
    save_data_to_xlsx(excel_workbook, file_path)
    print('File "coursera_courses.xlsx" was successfully created')
