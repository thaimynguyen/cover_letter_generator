from mailmerge import MailMerge
import pandas as pd
import os

# Import sample cover letter Word file
template = "sample_cv_with_MergeField.docx"

# Import the CSV field input file 
job_list = pd.read_csv("cv_field_input.csv")
# Convert the dataframe into list of dictionaries, each of which contains all required text fields for one cover letter
job_list = job_list.to_dict(orient='records')

#fill in the required values and generate new Word files
for job in job_list:
    document = MailMerge(template)
    company_name = job['company_name']
    document.merge(
        date = job['date'],
        company_name = job['company_name'],
        job_title = job['job_title'],
        job_title_uppercase = job['job_title'].upper(),
        job_number = str(job['job_number']),
        company_comment = job['company_comment']
        )
    document.write(f'Completed_CVs/Cover_Letter_{company_name}.docx')
document.close()

# convert Word files to PDF files
convert("Completed_CVs")