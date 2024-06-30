from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import os

# Code block to check CLs folder exists to save new CLs:
existing_folder = 'CLs'
if not os.path.exists(existing_folder):
    os.makedirs(existing_folder)
    
    
# Variables to hold template replacable items and add to dictionary:
doc = DocxTemplate("cover-letter-template.docx")
my_name = "Full Name"
my_phone = "(987) 354 - 321"
my_email = "my@email.com"
my_address = "Street Number, City, State 12345"
my_linkedin = "linkedin.com/in/alma-camarillo"
today_date = datetime.today().strftime("%b %d, %Y")

my_context = {'my_name' : my_name,
              'my_phone' : my_phone,
              'my_email' : my_email,
              'my_address' : my_address,
              'today_date' : today_date,
              'my_linkedin' : my_linkedin}


# This accesses the spreadsheet with data on each company
df = pd.read_csv('fake_data.csv')
for index, row in df.iterrows():
    """
    To test info from correct file is printing properly, 
    uncomment two print statements below, otherwise,
    delete them or uncomment them whenever you deem best.
    """
    # print(index)
    # print(row)
    
    context = {'hiring_manager_name' : row['name'],
                    'address' : row['address'],
                    'phone_number' : row['phone_number'],
                    'email' : row['email'],
                    'job_position' : row['job'],
                    'company_name' : row['company']}
    
    context.update(my_context)
    label = context['job_position']

# Saving each new doc for each new company:
doc.render(context)
file_path = os.path.join(existing_folder, f"CL_{label}.docx")
doc.save(file_path)