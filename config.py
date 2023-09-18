
#COOKIE_FILE = 'dtt.cookies' # name of file to store web session cookies in
REQUESTS_TIMEOUT = 30       # seconds

# siteid for the NHQDCSDLC site
#SITE_ID = 'americanredcross.sharepoint.com,38988760-70fd-4850-90e4-61f59a1e3bbf,4e1787c4-bf1b-4828-876a-6d7b1613ddec'

# info needed to access the AVIS report in National HQ's DCS Disaster Logistics Center sharepoint
NHQDCSDLC_DRIVEID = 'b!YIeYOP1wUEiQ5GH1mh47v8SHF04bvyhIh2ptexYT3ewviNAPjQJ8SJM6MEC7Zdmh'
FYxx_ITEM_PATH = '/Gray Sky/Avis Reports/FY24 Avis Report'

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
    def __init__(self, dr_num, dr_year, send_email, dtt_user, target_list, reply_email=None, extra_drs=None, suppress_erv_mail=True, avis_list=None):
        """ extra_drs is an array of (dr_num, dr_year) tuples.  Probably rarely needed, but DR155 changed to DR285 """
        self._dr_num = dr_num.rjust(3, '0')
        self._dr_year = dr_year
        self._send_email = send_email
        self._dtt_user = dtt_user
        self._target_list = target_list
        self._avis_list = avis_list

        self._email_bcc = EMAIL_BCC
        self._program_email = PROGRAM_EMAIL

        self._reply_email = (reply_email if reply_email != None else send_email)

        self._id = None
        self._name = None

        self._extra_drs = extra_drs
        self._suppress_erv_mail = suppress_erv_mail

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
    def id(self):
        return self._id

    @id.setter
    def id(self, value):
        self._id = value

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def token_filename(self):
        return f"o365_token-{ self.dr_id }.txt"

    @property
    def cookie_filename(self):
        return f"dtt_cookies-{ self.dr_id }.txt"

    @property
    def extra_drs(self):
        return self._extra_drs

    @property
    def suppress_erv_mail(self):
        return self._suppress_erv_mail

    def get_dr_list(self):
        retval = [ (self.dr_num, self.dr_year) ]
        if self.extra_drs != None:
            retval.extend(self.extra_drs)
        return retval

    @property
    def avis_list(self):
        if self._avis_list is not None:
            return self._avis_list
        return self._reply_email


#       DRNum   DRYear  Send Email                        DTT User                           Target List for group-vehicle
DRConfig('155', '22', 'DR155-22Log-Tra2@redcross.org', 'DR155-22Log-Tra2@redcross.org', 'dr155-22-tra-reports@americanredcross.onmicrosoft.com')
DRConfig('285', '22', 'DR285-22Transport@redcross.org', 'DR285-22Log-Tra2@redcross.org', 'dr155-22-tra-reports@americanredcross.onmicrosoft.com',
        extra_drs=[ ('155', '22') ]
        )

DRConfig('204', '22', 'DR204-22Log-Tra2@redcross.org', 'DR204-22Log-Tra2@redcross.org', 'DR204-22Log-Tra1@redcross.org', reply_email='DR204-22Log-Tra1@redcross.org')
DRConfig('225', '22', 'DR225-22Log-Tra2@redcross.org', 'DR225-22Log-Tra2@redcross.org', 'DR225-22Log-Tra1@redcross.org', reply_email='DR225-22Log-Tra1@redcross.org')
DRConfig('234', '22', 'DR234-22Log-Tra2@redcross.org', 'DR234-22Log-Tra2@redcross.org', 'DR234-22Log-Tra1@redcross.org', reply_email='DR234-22Log-Tra1@redcross.org')
DRConfig('255', '22', 'DR255-22Log-Tra9@redcross.org', 'DR255-22Log-Tra9@redcross.org', 'DR255-22Log-Tra9@redcross.org', reply_email='DR255-22Log-Tra9@redcross.org')
DRConfig('337', '22', 'DR337-22Log-Tra2@redcross.org', 'DR337-22Log-Tra2@redcross.org', 'dr337-22groupvehiclereport@AmericanRedCross.onmicrosoft.com', reply_email='DR337-22Log-Tra2@redcross.org')
DRConfig('466', '22', 'DR466-22Log-Tra9@redcross.org', 'DR466-22Log-Tra9@redcross.org', 'DR466-22Log-Tra1@redcross.org')
DRConfig('606', '22', 'DR606-22Log-Tra9@redcross.org', 'DR606-22Log-Tra9@redcross.org', 'harry.feirman@redcross.org', reply_email='harry.feirman2@redcross.org')
DRConfig('637', '22', 'DR637-22Log-Tra9@redcross.org', 'DR637-22Log-Tra9@redcross.org', 'thomas.altavilla@redcross.org', reply_email='thomas.altavilla@redcross.org')

DRConfig('739', '23', 'DR739-23Log-Tra3@redcross.org', 'DR739-23Log-Tra3@redcross.org', 'dr739-23dailytransportationreport@AmericanRedCross.onmicrosoft.com', reply_email='DR739-23Log-Tra3@redcross.org')
DRConfig('765', '23', 'DR765-23Log-Tra9@redcross.org', 'DR765-23Log-Tra9@redcross.org', 'DR765-23Log-Tra1@redcross.org', reply_email='harry.feirman2@redcross.org')
DRConfig('766', '23', 'DR766-23Log-Tra9@redcross.org', 'DR766-23Log-Tra9@redcross.org', 'dr766-23dailytransportationreport@AmericanRedCross.onmicrosoft.com', reply_email='DR766-23Log-Tra1@redcross.org')
DRConfig('836', '23', 'DR836-23Log-Tra2@redcross.org', 'DR836-23Log-Tra1@redcross.org', 'dr836-23-daily-transportation-report@AmericanRedCross.onmicrosoft.com', reply_email='DR836-23Log-Tra2@redcross.org', suppress_erv_mail=True)

DRConfig('032', '23', 'DR032-23Log-Tra2@redcross.org', 'DR032-23Log-Tra1@redcross.org', 'dr032-23-daily-transportation-report@AmericanRedCross.onmicrosoft.com', reply_email='DR032-23Log-Tra2@redcross.org')
DRConfig('053', '23', 'DR053-23Log-Tra5@redcross.org', 'DR053-23Log-Tra1@redcross.org', 'dr053-23@AmericanRedCross.onmicrosoft.com', reply_email='DR053-23Log-Tra5@redcross.org', extra_drs=[ ('055', '23') ], avis_list='dr053-23-avis-reports@AmericanRedCross.onmicrosoft.com')
DRConfig('064', '23', 'DR064-23Log-Tra1@redcross.org', 'DR064-23Log-Tra1@redcross.org', 'dr064-23-reports@AmericanRedCross.onmicrosoft.com', reply_email='DR064-23Log-Tra1@redcross.org' )
DRConfig('080', '23', 'DR080-23Log-Tra2@redcross.org', 'DR080-23Log-Tra2@redcross.org', 'DR080-23Log-Tra2@redcross.org', reply_email='DR080-23Log-Tra2@redcross.org' )
DRConfig('176', '23', 'DR176-23Log-Tra2@redcross.org', 'DR176-23Log-Tra2@redcross.org', 'dr176-23-transportation-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('241', '23', 'DR241-23Log-Tra5@redcross.org', 'DR241-23Log-Tra5@redcross.org', 'dr241-23-transportation-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('243', '23', 'DR243-23Log-Tra5@redcross.org', 'DR243-23Log-Tra5@redcross.org', 'dr243-23-transportation-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('252', '23', 'DR252-23Log-Tra5@redcross.org', 'DR252-23Log-Tra5@redcross.org', 'dr252-23Log-Tra5@redcross.org' )
DRConfig('295', '23', 'DR295-23Log-Tra5@redcross.org', 'DR295-23Log-Tra5@redcross.org', 'dr295-23Log-Tra5@redcross.org' )
DRConfig('335', '23', 'DR335-23Log-Tra5@redcross.org', 'DR335-23Log-Tra5@redcross.org', 'dr335-23-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR335-23-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('420', '24', 'DR420-24Log-Tra5@redcross.org', 'DR420-24Log-Tra5@redcross.org', 'dr420-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR420-24-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('462', '24', 'DR462-24Log-Tra5@redcross.org', 'DR462-24Log-Tra5@redcross.org', 'dr462-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR462-24-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('486', '24', 'DR486-24Log-Tra5@redcross.org', 'DR486-24Log-Tra5@redcross.org', 'dr486-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR486-24-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('495', '24', 'DR495-24Log-Tra5@redcross.org', 'DR495-24Log-Tra5@redcross.org', 'dr495-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR495-24-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('510', '24', 'DR510-24Log-Tra5@redcross.org', 'DR510-24Log-Tra5@redcross.org', 'dr510-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR510-24-avis-reports@AmericanRedCross.onmicrosoft.com' )

