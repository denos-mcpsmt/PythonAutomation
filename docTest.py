import os

def find_files(filename, search_path):
   result = []

# Wlaking top-down from the root
   for root, dir, files in os.walk(search_path):
      for file in files:
         if filename in file:
            result.append(os.path.join(root, file))

   return result

def check_file(results):
   file = results.pop(0).split("\\")[-1]
   check = input("Found Roster: "+file+"\nUse this file? (Y/n)")
   if check.capitalize() == "Y":
      return file
   elif check.capitalize() == "N":
      if len(results) > 0:
         check_file(results)
      else:
         return None



result = check_file(find_files("Pearson VUE","C:/Users/Pearson/Downloads"))
if result:
   # Kick off next program
   print("Roster Selected.")
else:
   print("could not find a Roster file!")