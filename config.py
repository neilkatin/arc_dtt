
COOKIE_FILE = 'dtt.cookies' # name of file to store web session cookies in
REQUESTS_TIMEOUT = 30       # seconds

# siteid for the NHQDCSDLC site
#SITE_ID = 'americanredcross.sharepoint.com,38988760-70fd-4850-90e4-61f59a1e3bbf,4e1787c4-bf1b-4828-876a-6d7b1613ddec'

# info needed to access the AVIS report in National HQ's DCS Disaster Logistics Center sharepoint
NHQDCSDLC_DRIVEID = 'b!YIeYOP1wUEiQ5GH1mh47v8SHF04bvyhIh2ptexYT3ewviNAPjQJ8SJM6MEC7Zdmh'
FY21_ITEM_PATH = '/Gray Sky/Avis Reports/FY22 Avis Report'

# info needed to access the Transportation file store area
TRANS_REPORTS_DRIVEID = 'b!o_zImoyHfUGqAtgz-lV_puzIkc7DM09NtuhToAIWGRvIebczIDDAT7ZLeWVofoHk'
TRANS_REPORTS_FOLDER_PATH = '/Reports and Data Analytics workgroup/Auto-generated Reports'

# info about the current DR we are reporting on
DR_NUM = '155'       # without any leading zero
DR_YEAR = '22'      # two digit year; program will add optional leading '20'; yes this breaks next century...

#DR_ID = '516'       # internal DTT id for the DR
#DR_NAME = '155-2022 - GC Command 7/21 FOR'

MAIL_BCC = 'generic@askneil.com'
SEND_EMAIL = 'DR155-22Log-Tra2@redcross.org'
REPLY_EMAIL = SEND_EMAIL

TOKEN_FILENAME_MAIL= 'o365_token-mail.txt'
TOKEN_FILENAME_AVIS= 'o365_token-avis.txt'
