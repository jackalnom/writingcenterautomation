from O365 import Message, MSGraphProtocol, Account
import pandas as pd
import os
from dotenv import load_dotenv
import numpy as np

load_dotenv()

staff_name = os.getenv('staff_name')
microsoft_client_id = os.getenv('microsoft_client_id')
microsoft_secret_value = os.getenv('microsoft_secret_value')
microsoft_tenant_id = os.getenv('microsoft_tenant_id')
email_override = os.getenv('email_override')

orphan_df = pd.read_csv("orphans.csv")
tutor_emails_df = pd.read_csv("tutor_emails.csv")

# join and make sure every every tutor has a corresponding email in tutor_emails.csv
orphan_with_emails_df = orphan_df.merge(tutor_emails_df, on="tutor", how="left")
assert orphan_with_emails_df.shape[0] == orphan_df.shape[0]

has_errors = False
for index, row in orphan_with_emails_df.iterrows():
   if pd.isnull(row['email']):
      print(f"Tutor {row['tutor']} is missing from tutor_emails.csv. Please add them as a row and rerun.")
      has_errors = True

if has_errors:
   print("Stopping script until all tutors have emails. No emails have been sent.")
   quit()

grouped_by_tutor = orphan_with_emails_df.groupby('tutor')

protocol = MSGraphProtocol()
credentials = (microsoft_client_id,microsoft_secret_value)

account = Account(credentials, tenant_id=microsoft_tenant_id)
if account.authenticate(scopes=['basic', 'message_all', 'message_send', 'users']):
   print('Authenticated!')

for group_name, df_group in grouped_by_tutor:
   print(f"Tutor: {group_name}")

   missed_appointments = []
   for row_index, row in df_group.iterrows():
      print(f"{row['student']}")
      missed_appointments.append(f"""
      <p>
         {row['student']}<br>
         <strong> {row['time']} on {row['date']} in {row['room']}</strong>
      </p>
      """)

   m = account.new_message()
   if email_override is not None:
      m.to.add(email_override)
   else:
      m.to.add(row['email'])
   m.subject = f"orphan app - {row['tutor']}"
   body = f"""
       <html>
           <body>
           Dear {row['tutor']},

           <p>
           You are missing a client report form for the following appointments:
           </p>
           {" ".join(missed_appointments)}
           <p>
           If these students did not attend this session, please mark them as "missed."
            Remember in the future that students should be marked "Missed at ten minutes past
             the appointment start time.
           </p>
           <p>
           As a reminder client report forms are required for every session. Please 
           complete these at the end of each appointment, do not wait until the end of the day.
           </p>
           <p>
           Regards,
           </p>
           <p>
           {staff_name}
           </p>
           </body>
       </html>
       """
   m.body = body
   print(body)

   m.send()