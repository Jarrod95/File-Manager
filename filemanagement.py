# RETRIEVE ALL FILE NAMES IN DIRECTORY AND SUBDIRECTORIES
import os, pprint, time, glob, re
from fnmatch import fnmatch 
import pandas as pd
from datetime import datetime, date
from pathlib import Path
import datefinder
import shutil

# Using a copy of all therapist documents to test 
start = time.perf_counter()
root = r'C:\Users\info\OneDrive\1. M2M Administration\PROPOSALS\Jarrods Folder\Test Therapists'
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
# Remove therapists in deleted folder
docs = [x for x in docs if 'Z - DELETED CONTRACTORS' not in x]

#Break down files into states 
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
cert_of_currency = [x for x in cert_of_currency if 'Z - DELETED CONTRACTORS' not in x]
##cert_of_currency = [i.split('\\', -1)[-1] for i in cert_of_currency]
print(f'We have {len(cert_of_currency)} Certificates of Currency for contractors in our file system')

## captures expired files and moves them to 'Old Documents' dir in file path
expired = []
for x in cert_of_currency:
    z = Path(x).stem.split('Exp')
    for k in z:
        zz = datefinder.find_dates(k.strip())
        for i in zz:
            if i < datetime.now():
##                expired.append(x)
                print(i < datetime.now(), ': ', i)
                dest = x.rsplit('\\',1)[0] + '\\Old Documents\\' + x.rsplit('\\',1)[1]
                if not os.path.exists(dest):
                    os.makedirs(dest)
                shutil.move(x, dest)

## Splits on last folder
##for x in expired:
##	l =x.rsplit('\\',1)[0]
##	print(l)
			
#Search for all certs of currencies
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
cols = ['Police Check', 'Contractor Forms', 'Cert of Currency', 'First Aid', 'AC Agreement']
##items = [police_checks, contractor_forms, cert_of_currency, first_aid, ac_agreement]
df = pd.DataFrame(list(zip(police_checks, contractor_forms, cert_of_currency, first_aid, ac_agreement)), 
               columns =cols) 
##data = {'Police Check':police_checks, 'Contractor Form':contractor_forms, 'First Aid':first_aid}
##df = pd.DataFrame(data)

df.to_excel('C:\\Users\\info\\OneDrive\\Desktop\\Contractor Docs - '+ datetime.now().strftime("%d-%m-%Y")+'.xlsx')
finish = time.perf_counter()
print(f'\n\n\n Program finished in {round(finish,2)} second(s)')
