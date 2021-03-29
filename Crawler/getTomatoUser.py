import random
import re
import time
import xlwings as wing
from getTomato import findElement, get1X, getunfinished, getindexandurls, slowDown
from getTomatoCritic import store2CSV, getRows, getelement
from selenium import webdriver

# workbook path
databook = "./data/Moviedata1021_400_500.xlsx"

# dataset to store, can self define but the suffix should be ".csv" since we use csv package,
# if the file doen't exist, it will be created automatically based on your path.
usertxt = "./data/Users_400_500.csv"

pageslow = 2
# this is base url setting
baseurl = "https://www.rottentomatoes.com/m/"


def getuserstar(dr, regax):
    try:
        names = dr.find_elements_by_xpath(regax)
        star = 0
        for n in names:
            if n.get_attribute("class") == "star-display__filled ":
                star += 1
            elif n.get_attribute("class") == "star-display__half ":
                star += 0.5
        return star
    except:
        return "NA"


def getCritic(urls):
    br = webdriver.Edge()
    for url in urls:
        try:
            br.get(url)
            slowDown()
            if findElement(br, '//div[@id="mainColumn"]/h1'):
                yield "404"
            else:
                rowsx = '//li[contains(@class,"audience")]'
                rows = getRows(br, rowsx)
                if rows == False:
                    yield "404"
                critics = []
                total = 1
                movie = re.search(r"(?<=/m/).*(?=/)", url).group()
                print("Starting user reviews on " + movie)
                nextpage = '//button[@data-direction="next"]'
                while (findElement(br, nextpage)):
                    for i in range(1, len(rows) + 1, 1):
                        reviewer = get1X(br, rowsx + '[' + str(i) + ']' + '/div/div/a')
                        if reviewer == "NA":
                            continue
                        print("Getting user review " + str(total) + " on " + movie)
                        l = []
                        l.append(movie)
                        l.append(reviewer)
                        l.append(getuserstar(br, rowsx + '[' + str(i) + ']' + '/div[2]/span/span//span'))
                        l.append(get1X(br, rowsx + '[' + str(i) + ']' + '/div[2]/span[2]'))
                        l.append(get1X(br, rowsx + '[' + str(i) + ']' + '/div[2]/p[1]'))
                        critics.append(l)
                        total += 1
                    botton = getelement(br, nextpage)
                    if botton == None:
                        break
                    else:
                        try:
                            if total > 1501:
                                break
                            botton.click()
                            time.sleep(pageslow + random.random() * 2)
                        except:
                            break
            yield critics
        except:
            yield critics
    br.close()


if __name__ == "__main__":
    wp = wing.App(visible=True, add_book=False)
    wp.display_alerts = True
    wp.screen_updating = True
    workingbook = wp.books.open(databook)
    jobsheet = workingbook.sheets["joblist"]
    joblist_whole = jobsheet.used_range.value
    unfinished = getunfinished(joblist_whole, 4)
    index, urls = getindexandurls(unfinished, "/reviews?type=user")
    i = 0
    critics = getCritic(urls)
    for m in critics:
        if m == "404":
            print("Get user reviews on " + urls[i] + " failed index: " + str(int(index[i])))
        else:
            jobsheet.range("E" + str(int(index[i]))).value = 1
            store2CSV(usertxt, m)
            print("Get user reviews on " + urls[i] + " successful index: " + str(int(index[i])))
        i += 1
    workingbook.save()
    wp.quit()
