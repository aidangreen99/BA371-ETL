import os
import openpyxl
import sys
import getopt
import sqlite3

#py ETL.py "C:\Users\Aidan\Documents\BA371\Group\testdb - Copy.SQLITE" "C:\Users\Aidan\Documents\BA371\Group\research_productivity\research_productivity"

#setting up gloabal variables
role_dict = {}
activity_type_dict = {}
target_type_dict = {}
dept_dict = {}
db_file = ''
spreadsheet_root = ''

#lists babey!!
targets_data = []
target_types_data = []
co_auth_data = []
depts_list = ["Accounting", "Marketing", "Finance", "BIS", "Management", "Entrepreneurship"]
target_types_list = ["journal", "conference"]
activity_types_list = ["submitted", "accepted", "r&r", "rejected"]
roles_list = ["Contributor", "Lead", "Co_lead"]
#I'm sorry, this has to be done :(
fac_list = ["Josie Ross", "Tallulah Stewart", "Bridget Cox", "Sally Miller", "Bella Foster", 
          "Sapphire Carter", "Amelia Lewis", "Mindy Garcia", "Ebony Taylor", "Carrie Perry", 
          "Anaya Kelly", "Ayat Reed", "Ria Green", "Klaudia Jenkins", "Myah Watson", 
          "Alma Hughes", "Hannah Turner", "Henna Baker", "May Jackson", "Jo Barnes", 
          "Tayla Sanchez", "Bailey Howard", "Louis Bennett", "Caitlan Butler", "Darcey Davis", 
          "Alfred White", "Shayne Thomas", "Melvin Richardson", "Marvin Cook", "Leyton Ward", 
          "Jesse Wright", "Barney Torres", "Casper Brown", "Aarav Roberts", "Alexander Morgan", 
          "Wesley Adams", "Carwyn Rogers", "Simone Clark", "Wilbur Parker", "Kinsley Thompson",
          "Zakaria Cooper", "Garfield Anderson", "Pierce Bailey", "Dan Allen", "Rick Wilson",
          "Colin Jones", "Ayub Myers", "Robert Brooks", "Darcy Sullivan", "Howard Jones"]

class paper_data:
     "This helps keep track of what belongs to what paper"
     paper_title = ""
     coauthors = []
     target = ""
     fac_role = ""
     activity_dates = []
     activity = []
     def __init__(self, name):
          self.paper_title = name



class sheetData:
     "This class allows easy storage and retreival of data per faculty"
     faculty_name = ""
     dept_name = ""
     dept_id = 999
     papers = []
     target_name = []
     target_type = []
     activity_type = []
     activity_date = []
     




def clear_transaction_tables():
  target_query = "delete from target;"
  authors_query = "delete from fac_paper;"
  coauthors_query = "delete from co_auth_paper;"
  papers_query = "delete from papers;"
  activities_query = "delete from activities;"
  try:
    cursor.execute(target_query)
    cursor.execute(authors_query)
    cursor.execute(coauthors_query)
    cursor.execute(papers_query)
    cursor.execute(activities_query)
  except Exception as err:
    print("Error clearing out transaction tables...\n", err)
    exit(1)
  connection.commit()

#Fill the depts table, specifically to generate an ID for dict matching
def dept_fill():
     cursor.execute("DELETE FROM depts;")
     for dep in depts_list:
          query = "INSERT INTO depts (dept_name) VALUES ('"
          query += dep
          query += "');"
          #print(query)
          try:
               cursor.execute(query)       
          except Exception as err:
               print("QUERY EXECUTION ERROR: " + str(err))
               break
     connection.commit()        

#Fill the target_type table, specifically to generate an ID for dict matching
def target_types_fill():
     cursor.execute("DELETE FROM target_type;")
     for target_type in target_types_list:
          query = "INSERT INTO target_type (target_type_name) VALUES ('"
          query += target_type
          query += "');"
          #print(query)
          try:
               cursor.execute(query)       
          except Exception as err:
               print("QUERY EXECUTION ERROR: " + str(err))
               break
     connection.commit()        

#Fill the activity_type table, specifically to generate an ID for dict matching
def activity_types_fill():
     cursor.execute("DELETE FROM activity_type;")
     for activity_type in activity_types_list:
          query = "INSERT INTO activity_type (activity_type) VALUES ('"
          query += activity_type
          query += "');"
          #print(query)
          try:
               cursor.execute(query)       
          except Exception as err:
               print("QUERY EXECUTION ERROR: " + str(err))
               break
     connection.commit()        

#Fill the roles table, specifically to generate an ID for dict matching
def roles_fill():
     cursor.execute("DELETE FROM roles;")
     for role in roles_list:
          query = "INSERT INTO roles (role) VALUES ('"
          query += role
          query += "');"
          #print(query)
          try:
               cursor.execute(query)       
          except Exception as err:
               print("QUERY EXECUTION ERROR: " + str(err))
               break
     connection.commit()     

#Fill the faculty table, specifically to generate an ID for dict matching
def faculty_fill():
     cursor.execute("DELETE FROM faculty;")
     for name in fac_list:
          query = "INSERT INTO faculty (faculty_name) VALUES ('"
          query += name
          query += "');"
          #print(query)
          try:
               cursor.execute(query)       
          except Exception as err:
               print("QUERY EXECUTION ERROR: " + str(err))
               break
     connection.commit()     

def dict_fill():
     #Set up the look-up dicts for role_type, activity_type and target_type ids
     #Role types:
     query = "select role_id, role from roles;"
     try:
          cursor.execute(query)
     except Exception as err:
          print("Error executing query...\n", err)
          exit(1)
     for record in cursor.fetchall():
          role_dict[record[1]] = record[0]
          #print(record)

     #Activity types:
     query = "select activity_type_id, activity_type from activity_type;"
     try:
          cursor.execute(query)
     except Exception as err:
          print("Error executing query...\n", err)
          exit(1)
     for record in cursor.fetchall():
          activity_type_dict[record[1]] = record[0]
          #print(record)

     #Target types:
     query = "select target_type_id, target_type_name from target_type;"
     try:
          cursor.execute(query)
     except Exception as err:
          print("Error executing query...\n", err)
          exit(1)
     for record in cursor.fetchall():
          target_type_dict[record[1]] = record[0]
          #print(record)

     #Department IDs:
     query = "select dept_id, dept_name from depts;"
     try:
          cursor.execute(query)
     except Exception as err:
          print("Error executing query...\n", err)
          exit(1)
     for record in cursor.fetchall():
          dept_dict[record[1]] = record[0]
          #print(record)

def processsheets():
     #Variable 'spreadsheet_root' holds the name of the folder containing all the faculty folders
     for folder in os.listdir(spreadsheet_root):
          excel_file = spreadsheet_root + "\\" + folder + "\\" + folder + ".xlsx"
          #print("\n" + "Working on file: " + folder)
          try:
               wb = openpyxl.load_workbook(excel_file, data_only=True)
          except Exception as err:
               print("Error opening Excel file: " + excel_file + "\n", err)
               exit(1)

          #Get the sheet from the current workbook
          sheet = wb.worksheets[0]
          current_process = sheetData()
          if sheet.max_row > 6:
               #Process the spreadsheet code goes here
               index = 0
               current_process = sheetData()
               current_process.faculty_name = sheet['A3'].value
               current_process.dept_name = sheet['B3'].value
               current_process.dept_id = dept_dict[sheet['B3'].value]
               for row in sheet.iter_rows(min_row=7, max_col=1, max_row=50):
                    for cell in row:
                         if cell.value != None:
                              name = cell.value
                              current_process.papers.append(paper_data(name))
                              target = cell.offset(row=0,column=2).value
                              coauth1, coauth2, coauth3, coauth4 = cell.offset(row=0,column=4).value, cell.offset(row=0,column=5).value, cell.offset(row=0,column=6).value, cell.offset(row=0,column=7).value
                              role = cell


                              

               
               


          wb.close()





#Checking command line arguments
if len(sys.argv) != 3:
     print("etl.py '<database_file_path>' '<spreadsheet_root_path>'")
     sys.exit()
     
#assigning DB and spreadsheet file paths
db_file = sys.argv[1]
spreadsheet_root = sys.argv[2]

#connecting to DB
try:
     connection = sqlite3.connect(db_file)
except Exception as err:
     print("CONNECTION ERROR: " + str(err))
cursor =  connection.cursor()


#main basically
clear_transaction_tables()
dept_fill()
target_types_fill()
activity_types_fill()
roles_fill()
faculty_fill()
dict_fill()
processsheets()


cursor.close()
connection.close()
sys.exit()