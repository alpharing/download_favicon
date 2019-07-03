import favicon
import requests
import json
import cloudconvert
from openpyxl import *

# xlsx to Favicon
def download_favi_xlsx():
    wb = load_workbook('example.xlsx')
    ws = wb.active

    name_list = []
    uri_list = []

    # r[2] field = web site Name in Excel
    # r[3] field = web site URI in Excel 
    for r in ws.rows:
        if "http" in str(r[3].value):
            name_list.append(r[2].value)
            uri_list.append(r[3].value)

    for i in range(0, len(uri_list)):

        icons = favicon.get(uri_list[i])

        # icons(list) may not have .ico file
        for j in icons:
            if ".ico" in j.url:
                icon = j

        response = requests.get(icon.url, stream=True)

        with open('icon/' + name_list[i] + '.{}'.format(icon.format), 'wb') as image:
            for chunk in response.iter_content(1024):
                image.write(chunk)

    print("Favicon ICO Download Finished!")

    return name_list

# json to Favicon
def download_favi_json():
    with open('app.json') as json_file:
        json_data = json.load(json_file)

    for i in range(0, len(json_data)-1):
        uri_list.append(json_data[i]['uri'])
        name_list.append(json_data[i]['app_name'])

    for i in range(0, len(json_data)-1):

        icons = favicon.get(uri_list[i])

        # icons(list) may not have .ico file
        for j in icons:
            if ".ico" in j.url:
                icon = j

        response = requests.get(icon.url, stream=True)

        with open('icon/' + name_list[i] + '.{}'.format(icon.format), 'wb') as image:
            for chunk in response.iter_content(1024):
                image.write(chunk)

    print("Favicon ICO Download Finished!")

    return name_list

def ico_2_any(name_list):

    # Private API key - (https://cloudconvert.com/)
    api = cloudconvert.Api('PUT_YOUR_API_KEY')

    print("ICO TO GIF Converting...")
    cnt = 1

    for i in name_list:

        # If you want different format, Change 'outputformat' : 'PNG, GIF .. etc'
        process = api.convert({
            'inputformat': 'ico',
            'outputformat': 'gif',
            'input': 'upload',
            'file': open("icon/" + i + ".ico", 'rb')
        })
        process.wait()  
        process.download("icon/" + i + ".gif")  # download output file(Check your format)

        if cnt != len(name_list):
            print(str(cnt) + "/" + str(len(name_list)) + " progressed... ")
        else:
            print(str(cnt) + "/" + str(len(name_list)) + " finished!")

        cnt += 1

    print("Convert Finished!")

def main():
    name_list = []  # App Name List

    # If you don't have a json file, make one first.
    # name_list = download_favi()

    # If you don't have a xlsx file, make one first with your MS Excel.
    name_list = download_favi_xlsx()
    print(name_list)

    # ICO to Any Image Format(If you want a different format, Change output format in this fuction!
    # Using cloudconvert API (https://cloudconvert.com/)
    ico_2_any(name_list)

if __name__ == '__main__':
    main()
