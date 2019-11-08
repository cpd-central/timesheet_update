#email as key, tuple of timesheet path and phone number as value
def get_dictionary(year):
    email_phone_timesheet_dict = {
        "speichel@ceg-engineers.com": (f"H://CEG Timesheets//{year}//PeichelS.xls", "+16127430136"),
        "jmarsnik@ceg-engineers.com": (f"H://CEG Timesheets//{year}//MarsnikJ.xls", "+17158299927"),
        "rduncan@ceg-engineers.com": (f"H://CEG Timesheets//{year}//DuncanR.xls", None),
        "cdolan@ceg.mn": (f"H://CEG Timesheets//{year}//DolanC.xls", "+15076763319"),
        "kburk@ceg-engineers.com": (f"H://CEG Timesheets//{year}//BurkK.xls", None),
        "mkaas@ceg-engineers.com": (f"H://CEG Timesheets//{year}//KaasM.xls", None),
        "bahlsten@ceg-engineers.com": (f"H://CEG Timesheets//{year}//AhlstenB.xls", None),
        "mbartholomay@ceg-engineers.com": (f"H://CEG Timesheets//{year}//BartholomayM.xls", None),
        "dborkovic@ceg-engineers.com": (f"H://CEG Timesheets//{year}//BorkovicD.xls", None),
        "ebryden@ceg-engineers.com": (f"H://CEG Timesheets//{year}//BrydenE.xls", None),
        "rbuckingham@ceg-engineers.com": (f"H://CEG Timesheets//{year}//BuckinghamR.xls", None),
        "jcasanova@ceg-engineers.com": (f"H://CEG Timesheets//{year}//CasanovaJ.xls", None),
        "schowdhary@ceg-engineers.com": (f"H://CEG Timesheets//{year}//ChowdharyS.xls", None),
        "vince@ceg.mn": (f"H://CEG Timesheets//{year}//GranquistV.xls", None),
        "nguddeti@ceg-engineers.com": (f"H://CEG Timesheets//{year}//GuddetiN.xls", None),
        "siqbal@ceg-engineers.com": (f"H://CEG Timesheets//{year}//IqbalS.xls", None),
        "ajama@ceg-engineers.com": (f"H://CEG Timesheets//{year}//JamaA.xls", None),
        "skatz@ceg-engineers.com": (f"H://CEG Timesheets//{year}//KatzS.xls", None),
        "pmalamen@ceg-engineers.com": (f"H://CEG Timesheets//{year}//MalamenP.xls", None),
        "jmitchell@ceg-engineers.com": (f"H://CEG Timesheets//{year}//MitchellJ.xls", None),
        "ntmoe@ceg.mn": (f"H://CEG Timesheets//{year}//MoeN.xls", None),
        "jromero@ceg.mn": (f"H://CEG Timesheets//{year}//RomeroJ.xls", None),
        "dsindelar@ceg-engineers.com": (f"H://CEG Timesheets//{year}//SindelarD.xls", None),
        "turban@ceg-engineers.com": (f"H://CEG Timesheets//{year}//UrbanT.xls", None),
        "yzhang@ceg-engineers.com": (f"H://CEG Timesheets//{year}//ZhangY.xls", None),
        "mtuma@ceg-engineers.com": (f"H://CEG Timesheets//{year}//TumaM.xls", None)
    }
    return email_phone_timesheet_dict

def get_dictionary_develop(year):
    email_phone_timesheet_dict_develop = {
        "speichel@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//PeichelS.xls", "+16127430136"),
        "jmarsnik@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//MarsnikJ.xls", "+17158299927"),
        "rduncan@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//DuncanR.xls", None),
        "cdolan@ceg.mn": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//DolanC.xls", "+15076763319"),
        "kburk@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//BurkK.xls", None),
        "mkaas@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//KaasM.xls", None),
        "bahlsten@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//AhlstenB.xls", None),
        "mbartholomay@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//BartholomayM.xls", None),
        "dborkovic@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//BorkovicD.xls", None),
        "ebryden@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//BrydenE.xls", None),
        "rbuckingham@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//BuckinghamR.xls", None),
        "jcasanova@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//CasanovaJ.xls", None),
        "schowdhary@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//ChowdharyS.xls", None),
        "vince@ceg.mn": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//GranquistV.xls", None),
        "nguddeti@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//GuddetiN.xls", None),
        "siqbal@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//IqbalS.xls", None),
        "ajama@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//JamaA.xls", None),
        "skatz@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//KatzS.xls", None),
        "pmalamen@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//MalamenP.xls", None),
        "jmitchell@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//MitchellJ.xls", None),
        "ntmoe@ceg.mn": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//MoeN.xls", None),
        "jromero@ceg.mn": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//RomeroJ.xls", None),
        "dsindelar@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//SindelarD.xls", None),
        "turban@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//UrbanT.xls", None),
        "yzhang@ceg-engineers.com": ("C://Users/jmarsnik//Desktop//timesheet_test_folder//ZhangY.xls", None)
    }
    return email_phone_timesheet_dict_develop
