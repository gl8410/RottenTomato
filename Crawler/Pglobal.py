import csv
import os

def store2CSV(path, colnames, data):
    if os.path.isfile(path):
        with open(path, mode="a+", encoding="utf-8", newline="") as f:
            fcsv = csv.writer(f)
            for d in data:
                fcsv.writerow(d)
    else:
        with open(path, mode="w", encoding="utf-8", newline="") as f:
            fcsv = csv.writer(f)
            fcsv.writerow(colnames)
            for d in data:
                fcsv.writerow(d)