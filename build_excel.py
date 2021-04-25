from win32com.client import Dispatch
from datetime import datetime
import os
import csv

'''
    Uses pywin32 library to fetch all possible properties of file.
'''
def fetch_properties(shell_ns):

    MAX_ITERATIONS = 350
    properties_dict = {}

    for i in range(MAX_ITERATIONS):
        _prop = shell_ns.GetDetailsOf(None, i)
        if(_prop):
            properties_dict[_prop] = i

    return properties_dict

def parse_date(date):
    date_formats = ('\u200e%d-\u200e%m-\u200e%Y \u200f\u200e%H:%M %p', '\u200e%d-\u200e%m-\u200e%Y \u200f\u200e%H:%M')
    for fmt in date_formats:
        try:
            return datetime.strptime(date, fmt).date()
        except ValueError as e:
            pass

    return ""


def build_excel(ns, dirctory_path, properties, writer):
    writer.writerow(["file_name", "directory", "found_date"])
    
    for dirpath, _, files in os.walk(directory_path):
        for f in files:
            x = ns.GetDetailsOf(ns.ParseName(f), properties['Date taken'])
            y = ns.GetDetailsOf(ns.ParseName(f), properties['Media created'])
            if(x):
                dateTime = parse_date(x)
                writer.writerow([str(f), directory_path, dateTime])
            elif(y):
                dateTime = parse_date(y)
                writer.writerow([str(f), directory_path, dateTime])
            else:
                writer.writerow([str(f), directory_path, ""])


if __name__ == '__main__':

    download_directory = input("Please input relative path to the folder: ")
    os.chdir(download_directory)
    directory_path = str(os.getcwd())

    shell = Dispatch("Shell.Application")
    namespace = shell.NameSpace(str(os.getcwd()))

    properties = fetch_properties(namespace)

    with open('../description.csv', 'w', newline='') as csvFile:
        writer = csv.writer(csvFile)
        build_excel(namespace, directory_path, properties, writer)

    print("Saved description.csv in parent folder")