import re
import time
from urllib.parse import urlencode
import requests
from Pglobal import store2CSV

cookie = "__cfduid=d990f65504606b013929cbaa43331db6a1602832519; yidio_user_country_code=CN; se_language_autodetected=1; PHPSESSID=53eu67ctkb7hci2ufjmvoi7lv2; all_utm_params=DIRMKT_Bing_Organic_yidiocn_; yidio_utm_source=Bing; yidio_utm_medium=Organic; yidio_utm_campaign=yidiocn; yidio_ga_id=331547450739; _ga=GA1.2.1804139429.1602832535; _gid=GA1.2.534635472.1602832535; __gads=ID=77ac42146f9b199e-2263a66d1bc400e2:T=1602832787:RT=1602832787:S=ALNI_MZjvAFdnflFPoSzEsT0IIaY-1XgSQ; _bti=%7B%22app_id%22%3A%22yidio%22%2C%22attributes%22%3A%5B%7B%22name%22%3A%22created_at%22%2C%22value%22%3A%222020-10-16T07%3A19%3A58%2B00%3A00%22%7D%2C%7B%22name%22%3A%22last_updated%22%2C%22value%22%3A%222020-10-16T07%3A19%3A58%2B00%3A00%22%7D%5D%2C%22bsin%22%3A%22ZDHJhRZxT2U%2BXNuZ3XhCpuZh6j0frkKHidR1orsCvnLin3MeqMFyxP%2BWD28UAbCJebpxQetieQH9Aj71r%2FlmNg%3D%3D%22%2C%22created_at%22%3A%222020-10-16T07%3A19%3A58%2B00%3A00%22%2C%22last_updated%22%3A%222020-10-16T07%3A19%3A58%2B00%3A00%22%7D; MAIN_RANDOM_VARIABLE=41; dir-movie-state={%22genre%22:[%22science-fiction%22]}"
filepath = "./data/"
baseurl = "https://www.yidio.com/redesign/json/browse_results.php?"
columns = ["id", "name", "url", "urlname"]
headers = {
    "method": "GET",
    "accept-language": "en-GB,en-US;q=0.9,en;q=0.8,zh-CN;q=0.7,zh;q=0.6",
    "cache-control": "no-cache",
    "cookie": cookie,
    "pragma": "no-cache",
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "none",
    "sec-fetch-user": "?1",
    "upgrade-insecure-requests": "1",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36"
}

queryparams = {
    "genre": "science-fiction",
    "type": "movie",
    "index": "0",
    "limit": "100"
}

def getMovieRow(json):
    if json:
        items = json.get("response")
        for item in items:
            movie = []
            movie.append(str(item.get("id")))
            name = str(item.get("name"))
            movie.append(name)
            movie.append(str(item.get("url")))
            movie.append(getUrlname(name))
            yield movie

def getUrlname(name):
    if name:
        name = name.lower()
        names = re.findall(r"[0-9a-z]+", name, re.S)
        if len(names) == 1:
            return names[0]
        else:
            rn = ""
            for n in names:
                rn += n + "_"
            return rn[:-1]
    else:
        return ""

def getJson(url="", header={}, paras={}):
    url = url + urlencode(paras)
    try:
        req = requests.get(url, headers=header)
        if req.status_code == 200:
            req.encoding = "utf-8"
            return req.json()
    except requests.ConnectionError as e:
        return None
        print("Connection error", e.args)

def getMovie2csv(filepath):
    time.sleep(10)
    timestamp = int(time.time())
    filepath = filepath + "movie" + str(timestamp) + ".csv"
    for index in range(0, 2800, 100):
        data = []
        queryparams["index"] = index
        json = getJson(baseurl, headers, queryparams)
        print("正在处理" + str(index) + "~" + str(index+100)+"数据.")
        rows = getMovieRow(json)
        for r in rows:
            data.append(r)
        store2CSV(filepath, columns, data)

if __name__ == "__main__":
    getMovie2csv(filepath)
