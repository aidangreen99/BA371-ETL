# Extract, Transorm, and Load Program for BA371
This program is intended to extract data from a folder of .xlsx files and load it into a database. Names are hardcoded because of the static nature of the project, if this was meant to be a more dynamic program that can account for faculty coming or going I'd insert a function to insert the faculty names into the faculty list that would look something like this: 
```
def fac_list():
  for folder in os.listdir(spreadsheet_root):
    if folder not in fac_list:
      name_string = folder.replace('_', ' ')
      name_string = name_string.title()
      fac_list += name_string
```
      
      
