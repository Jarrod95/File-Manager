# FILE MANAGEMENT:
# This could run once a day just after my computer is turned on ??
# 1. Loop through all files in Contractor Documents that are not from Deleted contractors or in an 'Old' folder
# 2. Find Expired documents and move them to the 'Old' folder
# 3. Send an email to notify us what files have been moved
import os, pprint, time, glob, re
from fnmatch import fnmatch 
import pandas as pd
from datetime import datetime, date
from pathlib import Path
import datefinder
import shutil
import secrets

#Import therapists data from Zoho
therapists = pd.read_csv(secrets.contractor_csv, skiprows=1)[:-3]
# Using a copy of all therapist documents to test 
start = time.perf_counter()
root = secrets.root_path
pattern ='*'

## Function currently gets all file paths into a list
## 1. Need to split by state and organise
## 2. Need to split by name

docs = []
for path, subdirs, files in os.walk(root):
    for name in files:
        if fnmatch(name, pattern):
            docs.append(os.path.join(path,name))
            ## print(os.path.join(path,name))
# Remove therapists in unecessary folders - should be a better way to do this?
docs = [x for x in docs if 'Z - DELETED CONTRACTORS' not in x]
docs = [x for x in docs if 'Old Documents' not in x]
docs = [x for x in docs if 'Y - REFERENCE THERAPISTS' not in x]
docs = [x for x in docs if 'X - OTHER' not in x]
docs = [x for x in docs if 'Sophie Conidi' not in x]

#Break down files into states -- not currently being used for anything excecpt stats
VIC = []
for item in glob.glob(root + '\\VIC\\*'):
    VIC.append(item)
    
NSW = []
for item in glob.glob(root + '\\NSW\*'):
    NSW.append(item)

QLD = []
for item in glob.glob(root + '\\QLD\\*'):
    QLD.append(item)

WA = []
for item in glob.glob(root + '\\WA\\*'):
    WA.append(item)

TAS = []
for item in glob.glob(root + '\\TAS\\*'):
    TAS.append(item)

NT = []
for item in glob.glob(root + '\\NT\\*'):
    NT.append(item)

ACT = []
for item in glob.glob(root + '\\ACT\\*'):
    ACT.append(item)
#Break down folders into contractor names -- this works
#Get this into a for-loop inside the above loops ??????
vic_contractors = [i.split('\\', -1)[-1] for i in VIC]
nsw_contractors = [i.split('\\', -1)[-1] for i in NSW]
qld_contractors = [i.split('\\', -1)[-1] for i in QLD]
wa_contractors = [i.split('\\', -1)[-1] for i in WA]
tas_contractors = [i.split('\\', -1)[-1] for i in TAS]
nt_contractors = [i.split('\\', -1)[-1] for i in NT]
act_contractors = [i.split('\\', -1)[-1] for i in ACT]

#This isn't necessary? But could be for reporting/comparison, etc.???
total_contractors = len(wa_contractors) + len(vic_contractors) + len(qld_contractors) +\
      len(nsw_contractors) + len(tas_contractors) + len(nt_contractors) + \
      len(act_contractors) - 7

print(f'We have {len(wa_contractors)-1} contractors in WA \n\
We have {len(vic_contractors)} contractors in Victoria \n\
We have {len(qld_contractors)} contractors in QLD \n\
We have {len(nsw_contractors)} contractors in NSW \n\
We have {len(tas_contractors)} contractors in TAS \n\
We have {len(nt_contractors)} contractors in NT \n\
We have {total_contractors} contractors in total\n\n')

#Search for all police checks
police_checks = []
for check in glob.glob(root + '\\*\\*\\*Police Check*'):
	police_checks.append(check)
police_checks = [i.split('\\', -1)[-1] for i in police_checks]
print(f'We have {len(police_checks)} Police Checks for contractors in our file system')

#Search for all contractor forms
contractor_forms = []
for check in glob.glob(root + '\\*\\*\\*Contractor Form*'):
	contractor_forms.append(check)
contractor_forms = [i.split('\\', -1)[-1] for i in contractor_forms]
print(f'We have {len(contractor_forms)} Contractor Forms for contractors in our file system')

#Search for all certs of currencies
cert_of_currency = []
for check in glob.glob(root + '\\*\\*\\*Cert of Currency*'):
    cert_of_currency.append(check)
cert_of_currency = [i.split('\\', -1)[-1] for i in cert_of_currency]
print(f'We have {len(cert_of_currency)} Certificates of Currency for contractors in our file system')

## captures expired files and moves them to 'Old Documents' dir in file path
cols=['Name', 'State', 'Document', 'Expiry', 'Other 1', 'Other 2', 'Other 3']
expired = []
test3 = []
df = pd.DataFrame(columns=['Name', 'State', 'Document', 'Expiry'])
for x in docs:
    z = Path(x).stem.split('Exp')
    for k in z:
        test2 = []
        test2.append(re.split(r'[(|)]|Exp|\.', k))
        for y in test2:
            test3.append(dict(zip(cols,y)))
        zz = datefinder.find_dates(k.strip())
        for i in zz:
            if i < datetime.now():
##                expired.append(x)
                print(i < datetime.now(), ': ', i, ' - ', x)
                expired.append(i)
                i.strftime('%d/%m/%Y') # to human readable date
                l = x.rsplit('\\',1)[1]
                dest = x.rsplit('\\',1)[0] + '\\Old Documents\\' + x.rsplit('\\',1)[1]
                #Create df to be emailed
                test = re.split(r'[(|)]|Exp|\.', l)[:-1]
                #Ensure list is the right length
##                if len(test) < 4:
##                    break
                if len(test) > 4:
                    x = ' '.join(test[:-3])
                    del test[:-3]
                    test.insert(0, x)
                df.loc[len(df), :] = test
                #Move expired document to 'Old Documents' folder
##                if not os.path.exists(dest):
##                    os.makedirs(dest)
##                shutil.move(x, dest)
df = df.sort_values(by=['Name']).reset_index(drop=True)
#This almost works to combine dicts on 'Name' column
from collections import defaultdict
from functools import reduce
from itertools import chain
def merge(d1, d2, key='Name'):
	r = defaultdict(list)

	for k, v in chain(d1.items(), d2.items()):
		if k != key:
			r[k].extend(v if isinstance(v, list) else [v])

	return {**r, key: d1[key]}
common = defaultdict(list)
for d in test3:
	common[d['Name']].append(d)
result = [reduce(merge, value) for value in common.values()]
#Search for all certs of currencies
# Will this be necessary ???
first_aid = []
for check in glob.glob(root + '\\*\\*\\*First Aid*'):
	first_aid.append(check)
first_aid = [i.split('\\', -1)[-1] for i in first_aid]
print(f'We have {len(first_aid)} First Aid Certificates for contractors in our file system')

#Search for all aged care agreements
ac_agreement = []
for check in glob.glob(root + '\\*\\*\\*Aged Care Agreement*'):
	ac_agreement.append(check)
ac_agreement = [i.split('\\', -1)[-1] for i in ac_agreement]
print(f'We have {len(ac_agreement)} Aged Care Agreements for contractors in our file system')

##Create DataFrame
##cols = ['Police Check', 'Contractor Forms', 'Cert of Currency', 'First Aid', 'AC Agreement']
##items = [police_checks, contractor_forms, cert_of_currency, first_aid, ac_agreement]
##df = pd.DataFrame(list(zip(police_checks, contractor_forms, cert_of_currency, first_aid, ac_agreement)), 
##               columns =cols) 
##data = {'Police Check':police_checks, 'Contractor Form':contractor_forms, 'First Aid':first_aid}
##df = pd.DataFrame(data)

##df.to_excel('C:\\Users\\info\\OneDrive\\Desktop\\Contractor Docs - '+ datetime.now().strftime("%d-%m-%Y")+'.xlsx')

# Send an email with the dataframe of expired docs
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd
import matplotlib

def send_mail():
    sender = secrets.from_address
    receiver = [secrets.to_address]
    msg = MIMEMultipart('related')

    msg['Subject'] = "Test Dataframe"
    msg['From'] = sender
    msg['To'] = ", ".join(receiver)
    msg['Date'] = email.header.Header(email.utils.formatdate())
    html = """\
    <html>
      <head></head>
      <body>
        <p>Hi!<br><br>
           The following documents have been archived:<br><br><br>
            {0}
            <br><br><br><br>
           Kind Regards, <br><br>
           Your Friendly File Management Bot
        </p>
      </body>
    </html>
    """.format(df.to_html(index=False, border=None))

    partHTML = MIMEText(html, 'html')
    msg.attach(partHTML)
    ser = smtplib.SMTP(server, port)
    ser.login(secrets.from_address, secrets.password)
    ser.sendmail(sender, receiver, msg.as_string())

finish = time.perf_counter()
print(f'\n\n\n Program finished in {round(finish,2)} second(s)')
