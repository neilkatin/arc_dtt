
#COOKIE_FILE = 'dtt.cookies' # name of file to store web session cookies in
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
#DR_NUM = '155'       # without any leading zero
#DR_YEAR = '22'      # two digit year; program will add optional leading '20'; yes this breaks next century...

#DR_ID = '516'       # internal DTT id for the DR
#DR_NAME = '155-2022 - GC Command 7/21 FOR'

EMAIL_BCC = 'generic@askneil.com'
#SEND_EMAIL = 'DR155-22Log-Tra2@redcross.org'
#EMAIL_TARGET_LIST = 'dr155-22-staffing-reports@AmericanRedCross.onmicrosoft.com'
#EMAIL_TARGET_LIST = 'neil.katin@redcross.org'
#REPLY_EMAIL = SEND_EMAIL

#TOKEN_FILENAME_MAIL = 'o365_token-mail.txt'
TOKEN_FILENAME_AVIS = 'o365_token-avis.txt'

PROGRAM_EMAIL = 'DR-Report-Automation@redcross.org'

DR_CONFIGURATIONS = {}

class DRConfig:
    def __init__(self, dr_num, dr_year, send_email, dtt_user, target_list, reply_email=None):
        self._dr_num = dr_num.rjust(3, '0')
        self._dr_year = dr_year
        self._send_email = send_email
        self._dtt_user = dtt_user
        self._target_list = target_list

        self._email_bcc = EMAIL_BCC
        self._program_email = PROGRAM_EMAIL

        self._reply_email = reply_email if reply_email != None else send_email

        DR_CONFIGURATIONS[self.dr_id] = self

    @property
    def dr_num(self):
        return self._dr_num

    @property
    def dr_year(self):
        return self._dr_year

    @property
    def send_email(self):
        return self._send_email

    @property
    def reply_email(self):
        return self._reply_email

    @property
    def program_email(self):
        return self._program_email

    @property
    def dtt_user(self):
        return self._dtt_user

    @property
    def target_list(self):
        return self._target_list

    @property
    def email_bcc(self):
        return self._email_bcc

    @property
    def dr_id(self):
        return f"{ self._dr_num }-{ self._dr_year }"

    @property
    def token_filename(self):
        return f"o365_token-{ self.dr_id }.txt"

    @property
    def cookie_filename(self):
        return f"dtt_cookies-{ self.dr_id }.txt"


DRConfig('155', '22', 'DR155-22Log-Tra2@redcross.org', 'DR155-22Log-Tra2@redcross.org', 'dr155-22-staffing-reports@AmericanRedCross.onmicrosoft.com')
DRConfig('204', '22', 'DR204-22Log-Tra2@redcross.org', 'DR204-22Log-Tra2@redcross.org', 'DR204-22Log-Tra1@redcross.org', reply_email='DR204-22Log-Tra1@redcross.org')
DRConfig('225', '22', 'DR225-22Log-Tra2@redcross.org', 'DR225-22Log-Tra2@redcross.org', 'DR225-22Log-Tra1@redcross.org', reply_email='DR225-22Log-Tra1@redcross.org')
DRConfig('234', '22', 'DR234-22Log-Tra2@redcross.org', 'DR234-22Log-Tra2@redcross.org', 'DR234-22Log-Tra1@redcross.org', reply_email='DR234-22Log-Tra1@redcross.org')

