import xlrd
import sys
import json
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_schedule(workbook_path, group):
    workbook = None
    try:
        workbook = xlrd.open_workbook(workbook_path)
    except:
        print('Could not find: ', workbook_path)
        print('Aborting...')
        exit()
    subs = {}
    schedule = []
    daysOfWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']
    m = 0  # day index
    time = 1
    col = -1

    a = 4  # row containing group names used to iterate through schedule rows
    row_ending = 114  # ending of schedule rows
    ws = None  # open first tab

    # get col no. acc to group
    for sheet in workbook.sheets():
        for i in range(sheet.ncols):
            g = sheet.cell_value(4, i)
            if g != '' and g == group:
                col = i
                ws = sheet
                break
    if ws is None:
        print('Group ', group, ' Not Found...\nAborting.....')
        exit()
    first = col - 2 * (int(group[-1]) - 1)
    last = col + 2 * (8 - int(group[-1])) - 1
    # extract schedule
    while a < row_ending:
        a += 1
        data = ws.cell_value(a, col)
        dic = {}
        # logic for calculating time and m for day
        if time > 11:
            time = 1
            m += 1
            if m > 4:
                m = 4
        # if data='' then could be lecture (for non 1 groups eg h2) and data='L' for groups like h1
        if data == '' or data[-1] == 'L':
            val = str(ws.cell_value(a, first))  # cell(a,18) gives val for lecture
            # if val='' or not "L" =>break
            if val == '' or val[-1] != 'L':
                dic['time'] = time
                dic['day'] = daysOfWeek[m]
                dic['sub'] = 'break'
                a += 1
                time += 1
                # schedule.append(dic)
                continue
            # if last letter of val ='L' =>lecture
            if val[-1] == 'L':
                dic['sub'] = val

                dic['time'] = time
                dic['day'] = daysOfWeek[m]
                time += 1
                a += 1
                dic['location'] = ws.cell_value(a, first)
                dic['teacher'] = ws.cell_value(a, last)

        # if last letter of val ='P' =>lab
        elif data[-1] == 'P':
            dic['sub'] = data
            dic['time'] = time
            time += 2
            dic['day'] = daysOfWeek[m]
            a += 1
            z = ws.cell_value(a, col)
            a += 1
            z += " "
            z += ws.cell_value(a, col)
            dic['location'] = z
            a += 1
            z = ws.cell_value(a, col)
            dic['teacher'] = z


        # if last letter of val ='T' =>tut
        elif data[-1] == 'T':

            dic['sub'] = data
            dic['time'] = time
            time += 1
            dic['day'] = daysOfWeek[m]
            a += 1
            z = ws.cell_value(a, col)
            dic['location'] = z
            z = ws.cell_value(a, col + 1)
            dic['teacher'] = z

        val=dic['sub']
        val=val[:-2]
        if val in subs:
            t = subs[val]
            t.append(dic)
            subs[val] = t
        else:
            temp = [dic]
            subs[val] = temp
        schedule.append(dic)

    return subs


def main(args):
    try:
        excel_path = args[1]
    except:
        excel_path = input('Enter path of excel sheet: ')

    try:
        group = args[2]
    except:
        group = input('Enter group: ')
    subs = get_schedule(excel_path, group)
    create_schedule(subs,group)


def create_schedule(subs,group):
    print('Creating schedule')
    color = ["#C8F7C5","#99CCCC","#FFE37D","#E08283","#C5EFF7","#CC99CC"]
    subjects = []
    i=0

    for sub_code in subs:
        list_meeting=subs[sub_code]
        meeting_times = []

        for e in list_meeting:

            start_time = e['time']+7
            end_time = start_time+1
            # subtype = e['sub']
            if e['sub'][-1] =="P":
                subtype = "LAB"
                end_time+=1
            elif e['sub'][-1]=="T" :
                subtype = "TUT"
            else:
                subtype = "Lecture"

            days = {"monday": False, "tuesday": False, "wednesday": False, "thursday": False, "friday": False,
                    "saturday": False, "sunday": False}
            days[e['day']] = True
            meet = {
                "uid": "2993ac0d-115c-4524-8371-c2e9106cb9f4",
                "courseType": subtype,
                "instructor": e['teacher'],
                "location": e['location'],
                "startHour": start_time,
                "endHour": end_time,
                "startMinute": 0,
                "endMinute": 0,
                "days": days
            }

            meeting_times.append(meet)
        subject={
            "uid": "50c386ad-f7aa-49e7-a214-b455f9d245a6",
            "type": "Course",
            "title": sub_code,
            "meetingTimes": meeting_times,
            "backgroundColor": color[i]
        }
        i+=1

        subjects.append(subject);
    title=group
    csmo_output = {
        "dataCheck": "69761aa6-de4c-4013-b455-eb2a91fb2b76",
        "saveVersion": 4,
        "schedules": [{
            "title": title,
            "items": subjects
    }],
   "currentSchedule":0
    }
    f = open(title+'.csmo' ,'w')
    f.write(json.dumps(csmo_output))
    f.close()
    get_schedule_as_png(title)

def get_schedule_as_png(title):
    driver = webdriver.Chrome('./chromedriver.exe')
    driver.get('https://www.freecollegeschedulemaker.com/')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div/div[1]/div/div[7]/a"))).click()
    time.sleep(1)
    a = driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div/div/input')
    a.send_keys(os.path.abspath('./'+title+'.csmo'))
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div/div/div[2]/div[2]/a"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div/div[1]/div/div[4]/a"))).click()
    print('File is downloaded as '+title+'.png')
    time.sleep(3)
    driver.close()

if __name__ == "__main__":
    main(sys.argv)
