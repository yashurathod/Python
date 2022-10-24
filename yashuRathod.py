# %%
import json
from tqdm import tqdm
import xlsxwriter
file = open("file.txt", "r")

# %%
dataDict = {}
prevKey = ""
lines = file.readlines()
for line in lines:
    if "SCHEDULE" in line:
        print(line)
        dataDict[line.split("#")[1].strip()] = {}
        dataDict[line.split("#")[1].strip()]["start_times"] = []
        prevKey = line.split("#")[1].strip()
    elif "END" in line:
        prevKey = ""
        continue
    elif "DESCRIPTION" in line.split(" ")[0]:
        dataDict[prevKey]["description"] = " ".join(line.split(
            " ")[1:]).strip().removeprefix("\"").removesuffix("\"")
    elif "EXCEPT RUNCYCLE" in line:
        dataDict[prevKey]['exculde_calendar'] = " ".join(
            line.split(" ")[2:]).strip().removeprefix("\"").removesuffix("\"")
    elif "ON RUNCYCLE" in line:
        dataDict[prevKey]['date_condition'] = 1
        if "MO,TU,WE,TH,FR" in line:
            dataDict[prevKey]['days'] = "BUSINESS_WEEK"
        else:
            dataDict[prevKey]['days'] = "SOME_WEEK"
    elif "( AT" in line:
        dataDict[prevKey]['start_times'].append(line.split(" ")[2].strip())

# %%
dataDict

# %%

# %%

columnHeadings = ["KeyJob Name", "date_conditions",
                  "run_calendar", "start_times", "exclude_calendar"]
workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()

for columnHeading in columnHeadings:
    worksheet.write(0, columnHeadings.index(columnHeading), columnHeading)

for i, key in enumerate(tqdm(dataDict)):
    print(key)
    i += 1
    worksheet.write(i, 0, key)
    if "date_condition" in dataDict[key]:
        worksheet.write(i, 1, dataDict[key]['date_condition'])
    if "days" in dataDict[key]:
        worksheet.write(i, 2, dataDict[key]['days'])
    if "start_times" in dataDict[key]:
        worksheet.write(i, 3, ",".join(dataDict[key]['start_times']))
    if "exculde_calendar" in dataDict[key]:
        worksheet.write(i, 4, dataDict[key]['exculde_calendar'])

# %%


json.dump(dataDict, open("data.json", "w"))
workbook.close()

# %%
