import datetime
from pathlib import Path
import re
import logging
# import pprint
import months_cz
import openpyxl

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

EXP_DIR = Path('export')
ATTENDANCE = EXP_DIR / 'DOCHAZKA.TXT'

staff_data = {}


# prepare input data
def get_content_data():
    content = ''
    with open(ATTENDANCE, 'r') as a_file:
        line = True
        while line:
            line = a_file.readline()
            content += line.replace('  ', '')
    return content


c = get_content_data()


# split chunks of data for each person
# return array of people
def split_to_chunks(content):

    delimiter_re = re.compile(r'--+')
    delimiter_mo = delimiter_re.findall(content)
    delimiter = '-' * 80
    if delimiter_mo:
        delimiter = delimiter_mo[0]

    staff = content.split(delimiter)
    staff_members = []
    for i in range(1, len(staff), 2):
        staff_member = staff[i-1] + staff[i]
        member_data = staff_member.split('\n')
        while '' in member_data:
            member_data.remove('')
        if member_data[0].startswith('Odpraco') or member_data[0].startswith('Nepří'):
            member_data.remove(member_data[0])
        staff_members.append(member_data)
    return staff_members


all_staff = split_to_chunks(c)


name_num_RE = re.compile(r'(\d{4,})+"\s/\s([\w\sěščřžýáíéúůĚŠČŘŽÝÁÍÉÚŮóÓ,-.]+)')
# date_RE = re.compile(r'(^\d{2}\.\d{2}\.\d{2})\s(\d{2}:\d{2}).*(\d{2}:\d{2})\s*\S*\s(\w)?.*(\w{2})')
date_RE = re.compile(r'(^\d{2}\.\d{2}\.\d{2})\s(\d{2}:\d{2}).*(\d{2}:\d{2})\s*\S*\s([A-Z])?.*(\w{2})')

month_RE = re.compile(r'\d{2}\.\d{2}\.\d{4}')
month_mo = month_RE.search(c)
if month_mo:
    time_period = month_mo.group().split('.')
    month_num = int(time_period[1]) - 1
    year = time_period[2]
    month = months_cz.months_cz[month_num]
else:
    year = datetime.datetime.now().year
    month = 'mesic'

member_id = False
for person in all_staff:
    for data in person:
        name_id_mo = name_num_RE.findall(data)
        if name_id_mo:
            member_id, member_name = name_id_mo[0]
            staff_data.setdefault(member_id, {
                    'name': member_name,
                    'attendance': []
                }
            )

        att_mo = date_RE.findall(data)
        if att_mo and member_id:
            # logging.info(att_mo[0]) # ('03.01.22', '07:00', '15:00')
            staff_data[member_id]['attendance'].append(att_mo[0])

template = Path('template.xlsx')
if template.exists():
    wb = openpyxl.load_workbook(template)
    ws = wb.active

    for k, v in staff_data.items():
        new_worksheet = wb.copy_worksheet(ws)
        new_worksheet.title = v['name']
        new_worksheet.cell(1, 3).value = v['name']
        new_worksheet.cell(2, 3).value = f'{month} {year}'
        new_worksheet.cell(3, 3).value = k

        for workday in v['attendance']:

            row = int(workday[0][:2]) + 5
            new_worksheet.cell(row, 2).value = workday[4]
            new_worksheet.cell(row, 3).value = workday[1]
            new_worksheet.cell(row, 3).number_format = 'H:mm;@'
            new_worksheet.cell(row, 4).value = workday[2]
            new_worksheet.cell(row, 4).number_format = 'H:mm;@'
            new_worksheet.cell(row, 5).value = f'=(D{row}-C{row})*24'
            new_worksheet.cell(row, 6).value = workday[3]

    wb.remove(ws)
    wb.save(f'{month} {year}.xlsx')
