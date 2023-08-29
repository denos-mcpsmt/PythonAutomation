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
    #print(str(hour1)+"--"+str(hour2))
    if hour1 <= 3:
        hour1 += 12
    if hour2 <= 3:
        hour2 += 12

    if hour1 < hour2:
        return time1
    else:
        return time2


def extract_names(text):
    # Regular expression to find patterns of the form "Name1, Name2"
    # \b is a word boundary, which ensures we're getting whole words
    # [a-zA-Z]+ matches one or more alphabetic characters
    pattern = r"\b([a-zA-Z]+),\s*([a-zA-Z]+)\b"
    matches = re.findall(pattern, text)
    # Convert the matches into a list of strings
    names = [", ".join(match) for match in matches]
    return names

def extract_time_ranges(text):
    # Regular expression to find patterns of the form "HH:MM:SS Am/Pm -HH:MM:SS Am/Pm"
    # The pattern captures hours, minutes, seconds and am/pm in both start and end times
    pattern = r"(\d{1,2}:\d{2}:\d{2}\s*[AaPp][Mm])\s*-\s*(\d{1,2}:\d{2}:\d{2}\s*[AaPp][Mm])"
    matches = re.findall(pattern, text)
    # Convert the matches into a list of strings in the format "start_time - end_time"
    time_ranges = [" - ".join(match) for match in matches]
    return time_ranges

def extract_special_ids(text):
    # Regular expression to find patterns that start with a digit
    # followed by a combination of letters and digits.
    pattern = r"\b\d[A-Z]+\d+\b"
    ids = re.findall(pattern, text)
    print(ids)
    return ids

def extract_ids(text):
    # Regular expression to find patterns of the form "HXXXXXXXX"
    # The pattern captures an 'H' followed by several numeric digits
    pattern = r"\bH\d+\b"
    ids = re.findall(pattern, text)
    print(ids)
    return ids

def extract_test_type(text):
    # Regular expression to find patterns after "Montana HiSET" followed by a space
    # and ending just before a hyphen or newline.
    pattern = r"Montana HiSET ([A-Za-z\s]+)(?=\s*-|\n)"
    test_types = re.findall(pattern, text)
    return [test.strip() for test in test_types]

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


            student = remove_forms(student)
            print(student)
            student_name = extract_names(student.replace("\n", " "))
            last_name = student_name[0].split(',')[1].strip()
            first_name = student_name[0].split(',')[0]

            student_id = extract_ids(student)
            if student_id == []:
                student_id = extract_special_ids(student)
            time_range = extract_time_ranges(student)
            print(str(', '.join(student_name))+'--'+''.join(student_id)+'--'+str(time_range))
            student_name = first_name.title() + ' ' + last_name.title()
            test_type = extract_test_type(student)

            # Bring time back together and clean up
            test_start = time_range[0].split('-')[0].strip()
            # Grab earliest start time for each tester
            if student_name in start_times:
                start_times[student_name] = compare_starts(test_start,start_times[student_name])
            else:
                start_times[student_name] = test_start

            # Prepare entry to excel
            entry = [' '.join([first_name.title(),last_name.title()]),student_id[0],time_range[0],test_type[0]]
            cleaned_entries.append(entry)

# Build Roster Sign In Sheet

print(start_times)

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