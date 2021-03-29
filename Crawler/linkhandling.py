import os
import csv
import pandas as pd
import xlwings as wing

origin="./data/moviesOriginal.csv"
current="./data/current.csv"
data="./data/Moviedata.xlsx"

currentcolumns=["name","moviestatus","criticstatus","userstatus","link"]
tomato="https://www.rottentomatoes.com/m/"
#columns = ["id", "name", "url", "urlname"]

def makingCurrent(origin):
    dfo=pd.read_csv(origin)
    dfn=pd.DataFrame()
    dfn["name"]=dfo["name"]
    dfn["moviestatus"]=0
    dfn["criticstatus"] = 0
    dfn["userstatus"] = 0
    dfn["link"]=dfo["urlname"]
    dfn.columns=currentcolumns
    dfn.to_csv(current,mode="w",index=False)

def handlingWorking(wpath):
    wp=wing.App(visible=True,add_book=False)
    wp.display_alerts=True
    wp.screen_updating=True
    workingbook=wp.books.open(wpath)
    jobsheet=workingbook.sheets["joblist"]
    joblist_whole=jobsheet.used_range.value

def getunfinished(wholelist):
    unfinished=[]
    for l in wholelist:
        if l[2]==0:
            unfinished.append(l)
    return unfinished

if __name__=="__main__":
    # if os.path.exists(current)==False:
    #     makingCurrent(origin)
    handlingWorking(data)





