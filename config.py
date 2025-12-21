
#COOKIE_FILE = 'dtt.cookies' # name of file to store web session cookies in
REQUESTS_TIMEOUT = 30       # seconds

# siteid for the NHQDCSDLC site
#SITE_ID = 'americanredcross.sharepoint.com,38988760-70fd-4850-90e4-61f59a1e3bbf,4e1787c4-bf1b-4828-876a-6d7b1613ddec'

# info needed to access the AVIS report in National HQ's DCS Disaster Logistics Center sharepoint
NHQDCSDLC_DRIVEID = 'b!YIeYOP1wUEiQ5GH1mh47v8SHF04bvyhIh2ptexYT3ewviNAPjQJ8SJM6MEC7Zdmh'
FYxx_ITEM_PATH = '/Gray Sky/Avis Reports/Current Avis Reports'

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
    def __init__(self, dr_num, dr_year, send_email, dtt_user, target_list, reply_email=None, extra_drs=None, suppress_erv_mail=True, avis_list=None, staffing_subject=None):
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
        self._staffing_subject = staffing_subject

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
    def staffing_subject(self):
        return self._staffing_subject

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

    # turn an email into a list name to add remove instructions
    def get_list_name(self, email):
        index =  email.find('@')
        if index == -1:
            return None

        # strip off the domain; turn to lower case
        list_name = email[0:index]
        list_name = list_name.lower()

        # see if it is not one of the two 'standard' lists
        if not list_name.endswith('group-vehicle-reports') and not list_name.endswith('avis-reports'):
            return None

        return list_name



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
DRConfig('525', '24', 'DR525-24Log-Tra5@redcross.org', 'DR525-24Log-Tra5@redcross.org', 'dr525-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR525-24-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('525', '24', 'DR525-24Log-Tra5@redcross.org', 'DR525-24Log-Tra5@redcross.org', 'dr525-24-group-vehicle-reports@AmericanRedCross.onmicrosoft.com', avis_list='DR525-24-avis-reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('577', '24', 'DR577-24Log-Tra5@redcross.org', 'DR577-24Log-Tra5@redcross.org', 'DR577-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR577-24-Avis-Reports@AmericanRedCross.onmicrosoft.com' )
DRConfig('633', '24', 'DR633-24Log-Tra5@redcross.org', 'DR633-24Log-Tra5@redcross.org', 'DR633-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR633-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('663', '24', 'DR663-24Log-Tra5@redcross.org', 'DR663-24Log-Tra5@redcross.org', 'DR663-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR663-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('725', '24', 'DR725-24Log-Tra5@redcross.org', 'DR725-24Log-Tra5@redcross.org', 'DR725-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR725-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('754', '24', 'DR754-24Log-Tra5@redcross.org', 'DR754-24Log-Tra5@redcross.org', 'DR754-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR754-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('755', '24', 'DR755-24Log-Tra5@redcross.org', 'DR755-24Log-Tra5@redcross.org', 'DR755-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR755-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('770', '24', 'DR770-24Log-Tra5@redcross.org', 'DR770-24Log-Tra5@redcross.org', 'DR770-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR770-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('776', '24', 'DR776-24Log-Tra5@redcross.org', 'DR776-24Log-Tra5@redcross.org', 'DR776-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR776-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('788', '24', 'DR788-24Log-Tra5@redcross.org', 'DR788-24Log-Tra5@redcross.org', 'DR788-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR788-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('801', '24', 'DR801-24Log-Tra5@redcross.org', 'DR801-24Log-Tra5@redcross.org', 'DR801-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR801-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('811', '24', 'DR811-24Log-Tra5@redcross.org', 'DR811-24Log-Tra5@redcross.org', 'DR811-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR811-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('809', '24', 'DR809-24Log-Tra5@redcross.org', 'DR809-24Log-Tra5@redcross.org', 'DR809-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR809-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('829', '24', 'DR829-24Log-Tra5@redcross.org', 'DR829-24Log-Tra5@redcross.org', 'DR829-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR829-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('838', '24', 'DR838-24Log-Tra5@redcross.org', 'DR838-24Log-Tra5@redcross.org', 'DR838-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR838-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('843', '24', 'DR843-24Log-Tra5@redcross.org', 'DR843-24Log-Tra5@redcross.org', 'DR843-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR843-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('844', '24', 'DR844-24Log-Tra5@redcross.org', 'DR844-24Log-Tra5@redcross.org', 'DR844-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR844-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('845', '24', 'DR845-24Log-Tra5@redcross.org', 'DR845-24Log-Tra5@redcross.org', 'DR845-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR845-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('846', '24', 'DR846-24Log-Tra5@redcross.org', 'DR846-24Log-Tra5@redcross.org', 'DR846-24-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR846-24-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('857', '25', 'DR857-25Log-TRA5@redcross.org', 'DR857-25Log-TRA5@redcross.org', 'DR857-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR857-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('139', '25', 'DR139-25Log-TRA5@redcross.org', 'DR139-25Log-TRA5@redcross.org', 'DR139-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR139-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('159', '25', 'DR159-25Log-TRA5@redcross.org', 'DR159-25Log-TRA5@redcross.org', 'DR159-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR159-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('207', '25', 'DR207-25Log-TRA5@redcross.org', 'DR207-25Log-TRA5@redcross.org', 'DR207-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR207-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('211', '25', 'DR211-25Log-TRA5@redcross.org', 'DR211-25Log-TRA5@redcross.org', 'DR211-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR211-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('224', '25', 'DR224-25Log-TRA5@redcross.org', 'DR224-25Log-TRA5@redcross.org', 'DR224-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR224-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('220', '25', 'DR220-25Log-TRA5@redcross.org', 'DR220-25Log-TRA5@redcross.org', 'DR220-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR220-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('193', '25', 'DR193-25Log-TRA5@redcross.org', 'DR193-25Log-TRA5@redcross.org', 'DR193-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR193-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('215', '25', 'DR215-25Log-TRA5@redcross.org', 'DR215-25Log-TRA5@redcross.org', 'DR215-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR215-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('231', '25', 'DR231-25Log-TRA5@redcross.org', 'DR231-25Log-TRA5@redcross.org', 'DR231-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR231-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('268', '25', 'DR268-25Log-TRA5@redcross.org', 'DR268-25Log-TRA5@redcross.org', 'DR268-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR268-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('344', '25', 'DR344-25Log-TRA5@redcross.org', 'DR344-25Log-TRA5@redcross.org', 'DR344-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR344-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('363', '25', 'DR363-25Log-TRA5@redcross.org', 'DR363-25Log-TRA5@redcross.org', 'DR363-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR363-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('370', '25', 'DR370-25Log-TRA5@redcross.org', 'DR370-25Log-TRA5@redcross.org', 'DR370-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR370-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('500', '25', 'DR500-25Log-TRA5@redcross.org', 'DR500-25Log-TRA5@redcross.org', 'DR500-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR500-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('484', '25', 'DR484-25Log-TRA5@redcross.org', 'DR484-25Log-TRA5@redcross.org', 'DR484-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR484-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('503', '25', 'DR503-25Log-TRA5@redcross.org', 'DR503-25Log-TRA5@redcross.org', 'DR503-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR503-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False )
DRConfig('539', '25', 'DR539-25Log-TRA5@redcross.org', 'DR539-25Log-TRA5@redcross.org', 'DR539-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR539-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('535', '25', 'DR535-25Log-TRA5@redcross.org', 'DR535-25Log-TRA5@redcross.org', 'DR535-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR535-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('515', '25', 'DR515-25Log-TRA5@redcross.org', 'DR515-25Log-TRA5@redcross.org', 'DR515-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR539-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False, staffing_subject='DR515-25 Automated Staffing Reports')
DRConfig('549', '25', 'DR549-25Log-TRA5@redcross.org', 'DR549-25Log-TRA5@redcross.org', 'DR549-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR549-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('550', '25', 'DR550-25Log-TRA5@redcross.org', 'DR550-25Log-TRA5@redcross.org', 'DR550-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR550-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('548', '25', 'DR548-25Log-TRA5@redcross.org', 'DR548-25Log-TRA5@redcross.org', 'DR548-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR548-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('586', '25', 'DR586-25Log-TRA5@redcross.org', 'DR586-25Log-TRA5@redcross.org', 'DR586-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR586-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('594', '25', 'DR594-25Log-TRA5@redcross.org', 'DR594-25Log-TRA5@redcross.org', 'DR594-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR594-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('740', '26', 'DR740-26Log-TRA5@redcross.org', 'DR740-26Log-TRA5@redcross.org', 'DR740-26-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR740-26-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('558', '26', 'DR558-26Log-TRA5@redcross.org', 'DR558-26Log-TRA5@redcross.org', 'DR558-26-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR558-26-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('691', '26', 'DR691-26Log-TRA5@redcross.org', 'DR691-26Log-TRA5@redcross.org', 'DR691-26-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR691-26-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('661', '25', 'DR661-25Log-TRA5@redcross.org', 'DR661-25Log-TRA5@redcross.org', 'DR661-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR661-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('676', '25', 'DR676-25Log-TRA5@redcross.org', 'DR676-25Log-TRA5@redcross.org', 'DR676-25-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR676-25-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('793', '26', 'DR793-26Log-TRA5@redcross.org', 'DR793-26Log-TRA5@redcross.org', 'DR793-26-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR793-26-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)
DRConfig('033', '26', 'DR033-26Log-TRA5@redcross.org', 'DR033-26Log-TRA5@redcross.org', 'DR033-26-Group-Vehicle-Reports@AmericanRedCross.onmicrosoft.com', avis_list='DR033-26-Avis-Reports@AmericanRedCross.onmicrosoft.com', suppress_erv_mail=False)




