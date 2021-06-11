import pandas as pd
import numpy as np
import shutil
import os
import zipfile
from datetime import date, datetime


def update_dept(headers):
    dept = []
    dept.append(headers)
    dept.append(['department','UPDATE','dept_BrightspaceCourses_d2l','Brightspace Courses','','','','','','','',''])
    df = pd.DataFrame(dept)
    df.to_csv('1-Department.csv', header=False, index=False, sep=',')


def create_district(doc_path, temp_list, grades_list, headers, unc):
    course_index = []
    WS = pd.read_excel(doc_path)
    WS_np = np.array(WS)
    district_list = list(WS_np[0])
    district_data = []
    district_data.append('District')
    district_data.append('CREATE')
    district_data.append(aca_year + unc + 'D_FY_' + str(district_list[1]).replace(' ','_'))
    district_data.append(str(district_list[1]))
    district_code = aca_year + unc + 'D_FY_' + str(district_list[1]).replace(' ','_')
    district_data = district_data + ['','']
    district_data.append('1')
    district_data = district_data + ['','','','']
    district_list.append('6606')
    district = [headers, district_data]
    df = pd.DataFrame(district)
    df.to_csv('1-District.csv', header=False, index=False, sep=',')
    create_schools(district_code, doc_path, temp_list, grades_list, headers, unc)
    course_index = create_schools(district_code, doc_path, temp_list, grades_list, headers, unc)
    return course_index


def main():
    doc_path = 'upload.xlsx'
    aca_year = '2021'
    ucode = 'BSD'
    headers = ['type','action','code','name','start_date','end_date','is_active','department_code','template_code','semester_code','offering_code','custom_code']
    # temp_list = [['3gm','3rd Grade Master'],['4gm','4th Grade Master'],['5gm','5th Grade Master']]
    temp_list = [
        ['PreK-K_GRD_T','PreK-K Template (7017)'],
        ['1_GRD_T','1st Grade Template (7015)'],
        ['2_GRD_T','2nd Grade Template (7019)'],
        ['3_GRD_T','3rd Grade Template (6692)'],
        ['4_GRD_T','4th Grade Template (6687)'],
        ['5_GRD_T','5th Grade Template (7022)'],
        ['6_GRD_T','6th Grade Template (6689)'],
        ['7_GRD_T','7th Grade Template (7020)'],
        ['8_GRD_T','8th Grade Template (7023)'],
        ['9-10_GRD_T','9-10th Grade Template (7026)'],
        ['11-12_GRD_T','11th-12th Grade Template (7005)']
    ]
    grades_list  = [['PK','K','1','2','3','4','5'],['6','7','8'],['9','10','11','12']]
    update_dept(headers)
