#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import csv
import os
import shutil
from time import sleep

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

import DRCredentials


def modify_khan_csv(df, currentquarter):
    for name in range(1, len(df)):
        if df["Student name"][name] in studentMap:
            df["Student name"][name] = studentMap[df["Student name"][name]]
    df = df.replace("Not started", 0)

    for date in df.columns[2:]:
        if df[date][0][9:11] not in quarter[0]:
            df = df.drop(date, axis=1)
        elif df[date][0][9:11] == quarter[0][0] and int(df[date][0][12:14]) < int(quarter[1][0]):
            df = df.drop(date, axis=1)
        elif df[date][0][9:11] == quarter[0][-1] and int(df[date][0][12:14]) > int(quarter[1][-1]):
            df = df.drop(date, axis=1)

    # renames Khan Unit Test Assignments to Condensed form
    for assn in range(2, len(df.columns)):
        if df.columns[assn][:2] == "Un":
            name = ""
            for letter in df.columns[assn][11:15]:
                if letter == " ":
                    break
                else:
                    name += letter
            if f"Khan Q{currentquarter} {name}" + f" UT" in df.columns:
                for i in range(1, 10):
                    if f"Khan Q{currentquarter} {name}" + f" UT{i}" not in df.columns:
                        df.rename(columns={df.columns[assn]: f"Khan Q{currentquarter} {name}" + f" UT{i}"},
                                  inplace=True)
                        break
            else:
                df.rename(columns={df.columns[assn]: f"Khan Q{currentquarter} {name}" + f" UT"},
                          inplace=True)

    # recalculate exercise assignment grades to my grading criteria
    for x in range(2, len(df.columns)):
        for y in range(1, len(df)):
            if df.columns[x][:2] != "Kh":
                if df[df.columns[x]][y] == '100' or df[df.columns[x]][y] == '75' or df[df.columns[x]][y] == '2':
                    df[df.columns[x]][y] = '2'
                elif df[df.columns[x]][y] == '50' or df[df.columns[x]][y] == '1':
                    df[df.columns[x]][y] = '1'
                elif df[df.columns[x]][y] == '25' or df[df.columns[x]][y] == '.5':
                    df[df.columns[x]][y] = '.5'
                else:
                    df[df.columns[x]][y] = '0'
            else:
                df[df.columns[x]][y] = float(df[df.columns[x]][y]) / float(10)

    # months needed
    monthsUsed = set()
    for month in range(2, len(df.columns)):
        monthsUsed.add(months[df[df.columns[month]][0][9:11]])

    newScores = {}
    for month in monthsUsed:
        newScores[month] = [0] * len(df)

    # for each student, calculate their score earned in the months category and total it up
    for x in range(2, len(df.columns)):
        for y in range(0, len(df)):
            if df.columns[x][:2] != "Kh" and y != 0:
                newScores[months[df[df.columns[x]][0][9:11]]][y] += float(df[df.columns[x]][y])
            elif df.columns[x][:2] != "Kh":
                newScores[months[df[df.columns[x]][0][9:11]]][0] += 2

    # remove individual assignments now that they've been totaled
    for assn in df.columns:
        if assn[0:2] == "Ex" or assn == "Percent completed":
            df = df.drop(assn, axis=1)

    # add the monthly cumulative score of assignments
    for x in newScores:
        if newScores[x][0] != 0:
            df[x + " Khans"] = newScores[x]

    return df


options = Options()
options.add_experimental_option("detach", True)

pd.set_option("mode.chained_assignment", None)

months = {'01': "Jan", "02": "Feb", "03": "Mar",
          "04": "Apr", "05": "May", "06": "Jun",
          "07": "Jul", "08": "Aug", "09": "Sep",
          "10": "Oct", "11": "Nov", "12": "Dec"}

currentQuarter = 2
quarter = DRCredentials.academicCalendar[currentQuarter]
studentMap = DRCredentials.khanToSkyStudentMap

for classPeriod in DRCredentials.khanClasses[2:]:

    # go to khan academy and get a particular class periods assignment scores.

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                              options=options)
    driver.get("https://www.khanacademy.org/login")
    sleep(3)
    driver.find_element("xpath", "//input").send_keys(DRCredentials.khanLogin[0])
    driver.find_elements("xpath", "//input")[1].send_keys(DRCredentials.khanLogin[1])
    driver.find_elements("xpath", "//input")[1].send_keys(Keys.ENTER)
    sleep(3)

    driver.get(f"https://www.khanacademy.org/teacher/class/{classPeriod[1]}/assignment-scores")
    sleep(6)
    driver.find_element("xpath", "//button[.//span[text()[contains(.,'Download CSV')]]]").click()
    sleep(4)
    driver.find_element("xpath", "/html/body/div[3]/div/div/div/div/div[3]/div/button[2]").click()
    sleep(10)

    files = [f for f in os.listdir(f"/Users/{DRCredentials.computerUser}/Downloads") if f[4:10] == "Period"]

    for file in files:
        new_path = f"/Users/{DRCredentials.computerUser}/Desktop/KhanUpdater/" + file[0:3] + ".csv"
        shutil.move(f"/Users/{DRCredentials.computerUser}/Downloads/" + file, new_path)

    df = pd.read_csv(f"{classPeriod[0]}.csv")

    df = modify_khan_csv(df, currentQuarter)

    # get the names, total scores, and dates for the assignments that are needed to be graded

    assnToEnter = []
    scoreToEnter = []
    for assnName in df.columns[1:len(df.columns)]:
        assnToEnter.append(assnName)
        if str(df[assnName][0])[0] == "D":
            scoreToEnter.append("8")
        else:
            scoreToEnter.append(str(df[assnName][0])[:2])

    dateToEnter = []
    for x in df.columns:
        if x[:2] == 'St':
            pass
        elif x[:2] == 'Kh':
            dateToEnter.append([df[x][0][9:11], df[x][0][12:14]])
        else:
            key = [k for k, v in months.items() if v == x[:3]][0]
            if int(key) == int(quarter[0][0]):
                dateToEnter.append([key, quarter[1][0], key, '30'])
            elif int(key) == int(quarter[0][1]):
                if key == '02':
                    dateToEnter.append([key, '1', key, '28'])
                else:
                    dateToEnter.append([key, '1', key, '30'])
            elif int(key) == int(quarter[0][2]):
                dateToEnter.append([key, '1', key, quarter[1][1]])

    driver.close()

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                              options=options)

    # go to gradebook to upload grades
    driver.get(DRCredentials.gradeBookSite)

    original_window = driver.current_window_handle

    driver.find_element("xpath", "//*[@id='login']").send_keys(DRCredentials.skyLogin[0])
    driver.find_element("xpath", "//*[@id='password']").send_keys(DRCredentials.skyLogin[1])
    driver.find_element("xpath", "//*[@id='bLogin']").click()

    sleep(8)

    driver.close()

    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)

    driver.find_element("xpath", "//*[@id='nav_EAPTeacherAccess']/span").click()
    sleep(1)
    driver.find_element("xpath", "//*[@id='nav_EAPMyGradebook']/span").click()
    sleep(3)
    driver.find_element("xpath", f"//*[@id='classes']/tbody/tr[{classPeriod[2]}]/td[9]/a").click()
    sleep(4)
    driver.find_element("xpath", "//*[@id='menuAssign']/li/a/span").click()
    sleep(3)


    # get hold of all assignments currently listed on gradebook classperiod to find what needs to be added
    assn = [x for x in driver.find_elements("xpath", "//td[text()[contains(.,'Khan')]]") if x.text in assnToEnter]
    assnNames = [x.text for x in assn]

    # enter new assignments, or modify score of the monthly assignment as more problems have been added
    for x in assnToEnter:
        if x not in assnNames and "UT" in x:
            driver.find_element("xpath", "//*[@id='AltSortLine']/tbody/tr/td[1]/a").click()
            sleep(10)
            driver.find_element("xpath", "//select[@name='CATEGORY']/option[text()[contains(.,'Quiz')]]").click()
            driver.find_element("xpath", "//input[@name='DESC']").send_keys(f"{x}")
            driver.find_element("xpath", "//input[@name='ASSIGNGROUP']").send_keys("Khan")
            for y in range(5):
                driver.find_element("xpath", "//input[@name='MAXSCORE']").send_keys(Keys.BACKSPACE)
            driver.find_element("xpath", "//input[@name='MAXSCORE']").send_keys("8")
            for y in range(5):
                driver.find_element("xpath", "//input[@name='WEIGHT']").send_keys(Keys.BACKSPACE)
            driver.find_element("xpath", "//input[@name='WEIGHT']").send_keys("1.25")
            driver.find_element("xpath",
                                f"//select[@name='month1']/option[@value='{dateToEnter[assnToEnter.index(x)][0]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='day1']/option[@value='{dateToEnter[assnToEnter.index(x)][1]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='month2']/option[@value='{dateToEnter[assnToEnter.index(x)][0]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='day2']/option[@value='{dateToEnter[assnToEnter.index(x)][1]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='month3']/option[@value='{dateToEnter[assnToEnter.index(x)][0]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='day3']/option[@value='{dateToEnter[assnToEnter.index(x)][1]}']").click()
            driver.find_element("xpath", f"//td[text()='{classPeriod[0][0]}']/../td/input").click()
            sleep(10)
            driver.find_element("xpath", "//a[text()[contains(.,'Save and Back')]]").click()
            sleep(10)
        elif x not in assnNames:
            driver.find_element("xpath", "//*[@id='AltSortLine']/tbody/tr/td[1]/a").click()
            sleep(10)
            driver.find_element("xpath", "//select[@name='CATEGORY']/option[text()[contains(.,'Quiz')]]").click()
            driver.find_element("xpath", "//input[@name='DESC']").send_keys(f"{x}")
            driver.find_element("xpath", "//input[@name='ASSIGNGROUP']").send_keys("Khan")
            for y in range(5):
                driver.find_element("xpath", "//input[@name='MAXSCORE']").send_keys(Keys.BACKSPACE)
            driver.find_element("xpath", "//input[@name='MAXSCORE']").send_keys(str(scoreToEnter[assnToEnter.index(x)]))
            driver.find_element("xpath",
                                f"//select[@name='month1']/option[@value='{dateToEnter[assnToEnter.index(x)][0]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='day1']/option[@value='{dateToEnter[assnToEnter.index(x)][1]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='month2']/option[@value='{dateToEnter[assnToEnter.index(x)][2]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='day2']/option[@value='{dateToEnter[assnToEnter.index(x)][3]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='month3']/option[@value='{dateToEnter[assnToEnter.index(x)][2]}']").click()
            driver.find_element("xpath",
                                f"//select[@name='day3']/option[@value='{dateToEnter[assnToEnter.index(x)][3]}']").click()
            driver.find_element("xpath", f"//td[text()='{classPeriod[0][0]}']/../td/input").click()
            sleep(10)
            driver.find_element("xpath", "//a[text()[contains(.,'Save and Back')]]").click()
            sleep(10)
        elif x in assnNames and driver.find_element("xpath", f"//tr/td[text()[contains(.,'{x}')]]/../td[8]").text \
                != str(int(scoreToEnter[assnToEnter.index(x)])):
            driver.find_element("xpath", f"//tr/td[text()[contains(.,'{x}')]]").click()
            driver.find_element("xpath", "//a[text()='Edit']").click()
            sleep(10)
            for y in range(5):
                driver.find_element("xpath", "//input[@name='MAXSCORE']").send_keys(Keys.BACKSPACE)
            driver.find_element("xpath", "//input[@name='MAXSCORE']").send_keys(str(scoreToEnter[assnToEnter.index(x)]))
            driver.find_element("xpath", "//a[text()[contains(.,'Save')]]").click()
            sleep(10)

    driver.find_element("xpath", "//*[@id='nav_EAPTeacherAccess']/span").click()
    sleep(3)
    driver.find_element("xpath", "//*[@id='nav_EAPMyGradebook']/span").click()
    sleep(3)
    driver.find_element("xpath", f"//*[@id='classes']/tbody/tr[{classPeriod[2]}]/td[9]/a").click()
    sleep(4)
    a = ActionChains(driver)
    m = driver.find_element("xpath", "//*[@id='menuAssign']")
    a.move_to_element(m).perform()
    driver.find_element("xpath", "//*[@id='menuAssign']/li/ul/li[4]/a").click()
    sleep(4)


    driver.find_element("xpath", "//*[@id='showGraded']").click()
    driver.find_element("xpath", "//*[@id='showFuture']").click()

    # click all current quarter assignments to receive the csv needed for grade entry
    assn = [x for x in driver.find_elements("xpath", "//td[text()[contains(.,'Khan')]]") if x.text in assnToEnter]
    for x in assn:
        if x.text in assnToEnter:
            driver.find_element("xpath", f"//td[text()='{x.text}']/../td/input").click()

    driver.find_element("xpath", "//*[@id='exportBtn']").click()

    sleep(10)

    files = [f for f in os.listdir(f"/Users/{DRCredentials.computerUser}/Downloads") if classPeriod[3] in f]


    for file in files:
        new_path = f"/Users/{DRCredentials.computerUser}/Desktop/KhanUpdater/" + file
        shutil.move(f"/Users/{DRCredentials.computerUser}/Downloads/" + file, new_path)

    sleep(3)

    # for each assignment in the current quarter, enter the corresponding score the student earned found in the
    # other file. Once all grades have been entered to the csv that was downloaded, upload it so the online
    # gradebook will update with the current scores.

    with open(f"{classPeriod[3]}.csv", "r") as x:
        grades = csv.reader(x)
        grades = list(grades)

    assnNames = []
    for x in df.columns[1:len(df.columns)]:
        assnNames.append(x)

    current = None
    for x in grades:
        if x != [] and x[1] in assnNames:
            current = x[1]
        if current is not None and x != []:
            if x[1] in list(df["Student name"]):
                x[3] = float(df[df["Student name"] == x[1]][current])
                if x[3] == 0:
                    x[6] = "X"
                elif x[3] != 0 and x[6] == "X":
                    x[6] = ""

    with open(f"{classPeriod[0]}Complete.csv", "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(grades)

    driver.find_element("xpath", "//*[@id='importDiv']/table/tbody/tr/td[1]/form/input[1]").send_keys(
        f"/Users/{DRCredentials.computerUser}/Desktop/KhanUpdater/{classPeriod[0]}Complete.csv")
    driver.find_element("xpath", "//*[@id='importDiv']/table/tbody/tr/td[2]/a").click()
    sleep(15)

    driver.close()
    files = [f for f in os.listdir(f"/Users/{DRCredentials.computerUser}/Desktop/KhanUpdater/") if
             f[-3:len(f)] == "csv"]

    for x in files:
        os.remove(f"/Users/{DRCredentials.computerUser}/Desktop/KhanUpdater/" + x)
