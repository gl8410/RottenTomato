import csv
import os
import random
import re
import time

import xlwings as wing
from getTomato import findElement, \
    get1X, \
    getunfinished, \
    getindexandurls, \
    slowDown
from selenium import webdriver

# workbook path
databook = "./data/Moviedata1021_2000_2551.xlsx"

# dataset to store, can self define but the suffix should be ".csv" since we use csv package,
# if the file doen't exist, it will be created automatically based on your path.
critcstxt = "./data/Critics_2000_2551.csv"

# this is slow down for open pages in seconds waiting for data loading,
# if your connection is good, turn down to speed up
pageslow = 3

# this is base url setting
baseurl = "https://www.rottentomatoes.com/m/"


# Columns are critic name, title, fresh/rotten,

def store2CSV(path, data):
    if os.path.isfile(path):
        with open(path, mode="a+", encoding="utf-8", newline="") as f:
            fcsv = csv.writer(f, delimiter="|")
            for r in data:
                fcsv.writerow(r)
    else:
        with open(path, mode="w", encoding="utf-8", newline="") as f:
            fcsv = csv.writer(f, delimiter="|")
            for r in data:
                fcsv.writerow(r)


def hastest(r, x, positive: str = "Y", negative: str = "N"):
    try:
        r.find_element_by_xpath(x)
        return positive
    except:
        return negative


def getRows(br, xpath):
    try:
        rows = br.find_elements_by_xpath(xpath)
        return rows
    except:
        return False


def getelement(br, xpath):
    try:
        button = br.find_element_by_xpath(xpath)
        return button
    except:
        return None


def getCritic(urls):
    br = webdriver.Edge()
    for url in urls:
        try:
            br.get(url)
            slowDown()
            if findElement(br, '//div[@id="mainColumn"]/h1'):
                yield "404"
            else:
                rowsx = '//div[contains(@class,"review_table")]//div[contains(@class,"row")]'
                rows = getRows(br, rowsx)
                if rows == False:
                    yield "404"
                critics = []
                total = 1
                movie = re.search(r"(?<=/m/).*(?=/)", url).group()
                print("Starting on " + movie)
                nextpage = '//div[contains(@style,"bottom")]/a[contains(@class,"btn")]/span[contains(@class,"chevron-right")]'
                while (findElement(br, nextpage)):
                    for i in range(1, len(rows) + 1, 1):
                        reviewer = get1X(br, rowsx + '[' + str(
                            i) + ']' + '//div[contains(@class,"name")]/a[contains(@class,"bold")]')
                        if reviewer == "NA":
                            continue
                        print("Getting review " + str(total) + " on " + movie)
                        l = []
                        l.append(movie)
                        l.append(reviewer)
                        l.append(get1X(br, rowsx + '[' + str(i) + ']' + '//em[contains(@class,"subtle")]'))
                        l.append(hastest(br, rowsx + '[' + str(i) + ']' + '//div[contains(@class,"fresh")]', "F", "R"))
                        l.append(get1X(br, rowsx + '[' + str(i) + ']' + '//div[@class="the_review"]'))
                        score = get1X(br, rowsx + '[' + str(i) + ']' + '//div[contains(@class,"review-link")]')
                        score = re.search(r"(?<=Score: ).*\b", score)
                        l.append(getResearch1(score))
                        l.append(get1X(br, rowsx + '[' + str(i) + ']' + '//div[contains(@class,"review-date")]'))
                        l.append(hastest(br, rowsx + '[' + str(
                            i) + ']' + '/div[1]/div[2]/div[1]/span[contains(@class,"star")]'))
                        critics.append(l)
                        total += 1
                    botton = getelement(br, nextpage)
                    if botton == None:
                        break
                    else:
                        try:
                            botton.click()
                            time.sleep(pageslow + random.random() * 2)
                        except:
                            break
            yield critics
        except:
            yield critics
    br.close()


def getResearch1(score):
    try:
        return score.group()
    except:
        return "NA"


if __name__ == "__main__":
    wp = wing.App(visible=True, add_book=False)
    wp.display_alerts = True
    wp.screen_updating = True
    workingbook = wp.books.open(databook)
    jobsheet = workingbook.sheets["joblist"]
    joblist_whole = jobsheet.used_range.value
    unfinished = getunfinished(joblist_whole, 3)
    index, urls = getindexandurls(unfinished, "/reviews")
    i = 0
    critics = getCritic(urls)
    for m in critics:
        if m == "404":
            print("Get reviews on " + urls[i] + " failed index: " + str(int(index[i])))
        else:
            jobsheet.range("D" + str(int(index[i]))).value = 1
            store2CSV(critcstxt, m)
            print("Get reviews on " + urls[i] + " successful index: " + str(int(index[i])))
        i += 1
    workingbook.save()
    wp.quit()
