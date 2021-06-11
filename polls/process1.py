import pandas as pd
import numpy as np
import shutil
import os, boto3, logging
import zipfile
from datetime import date, datetime
from botocore.exceptions import ClientError

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


def upload_file(file_name, bucket, object_name=None):
    """Upload a file to an S3 bucket

    :param file_name: File to upload
    :param bucket: Bucket to upload to
    :param object_name: S3 object name. If not specified then file_name is used
    :return: True if file was uploaded, else False
    """

    # If S3 object_name was not specified, use file_name
    if object_name is None:
        object_name = file_name

    # Upload the file
    s3_client = boto3.client('s3')
    try:
        response = s3_client.upload_file(file_name, bucket, object_name)
    except ClientError as e:
        logging.error(e)
        return False
    return True


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


def create_schools(district_code, doc_path, temp_list, grades_list, headers, unc):
    schools_list = []
    course_index = []
    WS = pd.read_excel(doc_path)
    WS_np = np.array(WS)
    temp = []
    for item in WS_np:
        temp.append(list(item))
    temp.pop(0)
    temp.pop(0)
    temp.pop(0)
    for item in temp:
        if item[3] not in schools_list:
            schools_list.append([item[3],item[4]])
    schools = [[]]
    schools[0] = headers
    school_codes = []
    for school in schools_list:
        schools.append(['School','CREATE',aca_year + unc + 'S-FY'+school[0].replace(' ','_'),school[0],'','',1,'','','','',district_code])
        school.append(aca_year + unc + 'S-FY'+school[0].replace(' ','_'))
        df = pd.DataFrame(schools)
        df.to_csv('2-Schools.csv', header=False, index=False, sep=',')
    update_templates(temp_list, schools_list, grades_list, headers, unc)
    course_index = update_templates(temp_list, schools_list, grades_list, headers, unc)
    return course_index


def create_semester(aca_year, headers, unc):
    semester = []
    semester.append(headers)
    semester.append(['semester','CREATE',aca_year + unc + '_AY_Sem','AY ' + aca_year,'','','','','','','',''])
    df = pd.DataFrame(semester)
    df.to_csv('3-Semester.csv', header=False, index=False, sep=',')
    sem_code = aca_year + unc + '_AY_Sem'
    return sem_code


def update_templates(temp_list, schools_list, grades_list, headers, unc):
    templates = []
    course_index = []
    templates.append(headers)
    schools = []
    i_start = 0
    i_end = 0
    for school in schools_list:
        if school not in schools:
            schools.append(school)
    for school in schools:
        if school[1] == 'Elementary':
            i_start = 0
            i_end = 6
        if school[1] == 'Middle':
            i_start = 6
            i_end = 9
        if school[1] == 'High':
            i_start = 9
            i_end = 11
        print('Start: %s, End: %s' % (i_start, i_end))
        for i in range(i_start, i_end):
            print(i)
            templates.append(['course template','UPDATE',temp_list[i][0],temp_list[i][1],'','','',school[2],'','','',''])
    df = pd.DataFrame(templates)
    df.to_csv('4-Templates.csv', header=False, index=False, sep=',')
    create_offerings(temp_list, schools_list, grades_list, headers, unc)
    course_index = create_offerings(temp_list, schools_list, grades_list, headers, unc)
    return course_index



def create_offerings(temp_list, schools_list, grades_list, headers, unc):
    offerings = []
    course_index = []
    offerings.append(headers)
    temp = []
    sem_code = create_semester(aca_year, headers, unc)
    for school in schools_list:
        if school not in temp:
            temp.append(school)
    for school in temp:
        holder = school[2]
        if school[1] == 'Elementary':
            for level in grades_list[0]:
                if grades_list[0].index(level) <= 1:
                    x = 0
                else:
                    x = grades_list[0].index(level) - 1
                offerings.append(['course offering','CREATE',school[0].replace(' ','_') + '_' + level,school[0] + ' - Grade ' + level,'','','1','',temp_list[x][0],sem_code,'',holder])
        if school[1] == 'Middle':
            for level in grades_list[1]:
                x = grades_list[1].index(level) + 6
                offerings.append(['course offering','CREATE',school[0].replace(' ','_') + '_' + level,school[0] + ' - Grade ' + level,'','','1','',temp_list[x][0],sem_code,'',holder])
        if school[1] == "High":
            for level in grades_list[2]:
                if grades_list[2].index(level) >= 2:
                    x = 10
                else:
                    x = 9
                offerings.append(['course offering','CREATE',school[0].replace(' ','_') + '_' + level,school[0] + ' - Grade ' + level,'','','1','',temp_list[x][0],sem_code,'',holder])
    df = pd.DataFrame(offerings)
    for offering in offerings[1:]:
        try:
            off_code = offering[3].split(' - ')[1].split(' ')[1]
            off_name = offering[3].split(' - ')[0]
            course_index.append([off_name, off_code, offering[2], offering[11]])
        except:
            print('offering has a value not in the list range.')
    df.to_csv('5-Offerings.csv', header=False, index=False, sep=',')
    return course_index


def create_users():
    users = []
    headers = ['type','action','username','org_defined_id','first_name','last_name','password','is_active','role_name','email','relationships','pref_first_name','pref_last_name']
    WS = pd.read_excel(doc_path)
    WS_np = np.array(WS)
    user_list = list(WS_np)
    user_list.pop(0)
    user_list.pop(0)
    user_list.pop(0)
    users.append(headers)
    for user in user_list:
        users.append(['user','CREATE',user[2],user[2],user[0],user[1],'demos123!','1',user[5],user[2],'','',''])
    df = pd.DataFrame(users)
    df.to_csv('6-Users.csv', header=False, index=False, sep=',')
    return user_list


def create_enrollments(users, course_index):
    enrollments = []
    headers = ['type','action','child_code','role_name','parent_code']
    enrollments.append(headers)
    for user in users:
        user_grades = str(user[6]).split(',')
        for grade in user_grades:
            print(grade)
            print(course_index)
            for course in course_index:
                if course[1] == grade:
                    enrollments.append(['enrollment','CREATE',user[2],user[8],course[2]])
    df = pd.DataFrame(enrollments)
    df.to_csv('7-Enrollments.csv', header=False, index=False, sep=',')


def main(unc):
    if os.path.exists('polls/static/polls/downloads/download.zip'):
        print('path exists')
        os.remove('polls/static/polls/downloads/download.zip')
    else:
        print('path does not exist')
    course_index = []
    users = []
    update_dept(headers)
    create_semester(aca_year, headers, unc)
    create_district(doc_path, temp_list, grades_list, headers, unc)
    course_index = create_district(doc_path, temp_list, grades_list, headers, unc)
    create_users()
    users = create_users()
    create_enrollments(users, course_index)
    this_day = str(date.today())
    now = datetime.now()
    ct = str(now.strftime("%H:%M:%S").replace(':',''))
    zft = this_day + '-' + ct + '_' + ucode +'_IPSIS.zip'
    zf = zipfile.ZipFile('download.zip', mode='w')
    try:
        zf.write('manifest.json')
        zf.write('1-Department.csv')
        zf.write('1-District.csv')
        zf.write('2-Schools.csv')
        zf.write('3-Semester.csv')
        zf.write('4-Templates.csv')
        zf.write('5-Offerings.csv')
        zf.write('6-Users.csv')
        zf.write('7-Enrollments.csv')
    finally:
        zf.close()

        S3_BUCKET = os.environ.get('S3_BUCKET')
        #file_name = request.args.get('file_name')
        #file_type = request.args.get('file_type')

        s3 = boto3.client('s3')

        presigned_post = s3.generate_presigned_post(
          Bucket = S3_BUCKET,
          Key = file_name,
          Fields = {"acl": "public-read", "Content-Type": file_type},
          Conditions = [
            {"acl": "public-read"},
            {"Content-Type": file_type}
          ],
          ExpiresIn = 3600
        )

        return json.dumps({
          'data': presigned_post,
          'url': 'https://mtwbucket.s3.us-east-2.amazonaws.com/%s' % (S3_BUCKET, file_name)
        })

        '''
        filelist=['1-Department.csv','1-District.csv','2-Schools.csv','3-Semester.csv','4-Templates.csv','5-Offerings.csv','6-Users.csv','7-Enrollments.csv','upload.xlsx']
        for file in filelist:
            os.remove(file)
        shutil.move('download.zip', 'polls/')
        '''
