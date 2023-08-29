from PyPDF2 import PdfReader
import pandas as pd
import glob
import os
import re
import time
import datetime

# FUNCTIONS

# Compare start times and return earliest
def compare_starts(time1,time2):
    hour1 = int(time1.split(':')[0])
    hour2 = int(time2.split(':')[0])
    if hour1 <= 3:
        hour1 += 12
    if hour2 <= 3:
        hour2 += 12

    if hour1 < hour2:
        return time1
    else:
        return time2


# Capitalizes first letters and lowercases the rest, for proper splitting later
def cap_words(student):
    student_chunks = student.split(' ')
    fixed_chunks = []
    for chunk in student_chunks:
        chunked = list(chunk)
        # Cap letter if it is first letter of word and not capped
        if len(chunked) > 0:
            if not chunked[0].isupper() and chunked[0].isalpha():
                chunked[0] = chunked[0].upper()
        # Go through each letter after first, if letters are capped but not first of word, lowercase
        capped = False
        for x in range(len(chunked)):
            if chunked[x].isupper() and x != len(chunked)-1: # if this is upper
                if chunked[x+1].isupper():
                    chunked[x+1] = chunked[x+1].lower()
                    if x != len(chunked)-2:
                        if chunked[x+2].isupper():
                            chunked[x+2] = chunked[x+2].lower()

        fixed_chunks.append("".join(chunked))
    return ' '.join(fixed_chunks)
# Removes Form fields that mess up splitting later (some rosters do not have this)
def remove_forms(student):
    cleaned = ""
    if 'Form A' in student:
        cleaned = re.sub('Form A','',student)
    elif 'Form B' in student:
        cleaned = re.sub('Form B','',student)
    elif 'Form C' in student:
        cleaned = re.sub('Form C','',student)
    else:
        cleaned = student
    return cleaned
# Teases out 8-digit Student IDs from details
def get_id(student):
    chunks = list(student)
    stu_id = ""
    count = 0
    for letter in chunks:
        if letter.isalnum():
            stu_id += letter
            count += 1
        if count == 8:
            break
    return stu_id

# Determine most recent roster from Downloads folder
list_of_files = glob.glob('C:/Users/Pearson/Downloads/ShortFormAllCDs*.pdf')
if (len(list_of_files) > 1):
    latest = max(list_of_files,key=os.path.getctime)
else:
    latest = list_of_files.pop()

# Set up PDF raster
reader = PdfReader(latest)
text_list = []

# Extract text from PDF
for page in reader.pages:
    text_list.append(page.extract_text())

students_list = []
for text in text_list:
    # Split document at headers (don't need them)
    headers, entries = text.split('STREAMNAME')
    # Shorten each entry into its own row
    students_list.append(entries.replace("\n","").split('Computer Based Test'))

# For storing test taker details
cleaned_entries = []

start_times = {}

# Parsing
for students in students_list:
    for student in students:
        if len(student) > 30:
            # get rid of new lines and cap first letters (lowercase the rest)
            student = cap_words(student.replace("\n"," "))

            # remove whitespace, split on capital letters and rejoin with spaces
            student_cleaned = re.sub(' +', ' ',(' '.join(re.findall('[A-Z][^A-Z]*', student))))
            student_cleaned = remove_forms(student_cleaned)
            student_name = student_cleaned.split(',')[0:2] # Breaking for first and last names, expect two last names
            print(student_name)
            first_name = student_name[1].strip().split(' ')[0]
            last_name = student_name[0].strip()

            # Strip names from details
            student_details = re.sub(first_name,'',student_cleaned)
            student_details = re.sub(last_name,'',student_details)

            # Tease out ID number
            student_id = get_id(student_details)

            # Strip ID number from details, remove comma
            student_details = re.sub(student_id,'',student_details)
            student_details = re.sub(',','',student_details).strip()

            # Split on ':' then split on space and grab trailing and leading numbers
            student_time = student_details.split(':')
            student_time[0] = student_time[0].split(' ')[-1]
            student_time[-1] = student_time[-1].split(' ')[0]

            # Bring time back together and clean up
            student_time = ':'.join(student_time)
            student_time = re.sub(':00 |:00$','',student_time)
            test_time = re.sub(' +',' ',re.sub('-',' - ',re.sub('PM','',(re.sub('AM','',' '.join(student_time.split(' ')[0:4]).upper())))))
            test_start = test_time.split('-')[0].strip()
            student = cap_words(first_name)+' '+cap_words(last_name)
            # Grab earliest start time for each tester
            if student in start_times:
                start_times[student] = compare_starts(test_start,start_times[student])
            else:
                start_times[student] = test_start

            # Split on Hi Set to get test type
            test_type = student_details.split('Hi Set')[-1].split('-')[0:-1]
            test_type[0] = test_type[-1].strip() if test_type[0].strip() == 'Language Arts' else test_type[0].strip()

            # Prepare entry to excel
            entry = [' '.join([first_name.title(),last_name.title()]),student_id.upper(),test_time.strip(),test_type[0]]
            cleaned_entries.append(entry)

# Build Roster Sign In Sheet

pd_array = [["Name","ID","Time","Test"]]
pd_dict = {"Name": [], "ID": [],"OTP": [], "Time": [], "Test": [], "Sign In":[], "Time In": [], "Sign Out": [], "Time Out": []}
for entry in cleaned_entries:
    pd_dict["Name"].append(entry[0])
    pd_dict["ID"].append(entry[1])
    pd_dict["OTP"].append(" ")
    pd_dict["Time"].append(entry[2])
    pd_dict["Test"].append(entry[3])
    pd_dict["Sign In"].append(" ")
    pd_dict["Time In"].append(" ")
    pd_dict["Sign Out"].append(" ")
    pd_dict["Time Out"].append(" ")

df = pd.DataFrame(pd_dict)
df.style.set_caption("HiSET Sign In")
df.to_excel(excel_writer= "C:/Users/Pearson/Desktop/HiSET/HiSET-signin.xlsx")

# Build Front Office Excel Sheet

office = {"Name": [],"# Tests": [],"ID": [], "Lock": [], "Paid": [],"Address": [], "Start": []}
deduped = {}

# Enumerate student tests and deduplicate list
for entry in cleaned_entries:
    if entry[0] not in deduped.keys():
        deduped[entry[0]] = 1
    else:
        deduped[entry[0]] += 1

for k,v in deduped.items():
    office["Name"].append(k)
    office["# Tests"].append(v)
    office["Start"].append(start_times[k])
    office["Address"].append(' ')
    office["Lock"].append(" ")
    office["Paid"].append(" ")
    office["ID"].append(" ")


df = pd.DataFrame(office)
df.to_excel(excel_writer= "C:/Users/Pearson/Desktop/HiSET/HiSET-front.xlsx")

tester_count = len(office["Name"])
test_count = 0
for num in office["# Tests"]:
    test_count += num
print("You have "+str(tester_count)+" testers taking "+str(test_count)+" tests")
print("Roster sheets added to HiSet folder on Desktop.")
input("Press any key to exit")