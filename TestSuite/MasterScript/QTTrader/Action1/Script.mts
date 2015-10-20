'*************************************************************************************************************************************************
'                                          Master Driver 
'*************************************************************************************************************************************************
'DESCRIPTION:
'    This script is the Master Script, which will call the appropriate Driver Script based on the Input Data provided in the MasterDriver.xls

'ASSUMPTIONS:
'   None

'DEPENDENCIES:
'    MasterDriver.xls should be updated with the necessary test suite to be executed.
'    Create an environment variable named "EnvTestAutomationRootPath" ,which should have the root directory name of the Test Suite Path

'PARAMETERS:
'    None

'RETURN
'   None

'ERROR
'   None

'REVISION HISTORY
'Author:
'Create On Date:
'Change History:
'Date                    Name                  Modification Details
'***************************************************************************************************************************************************
Public g_TestSuitePath
Dim TestAutomationRootPath

TestAutomationRootPath = Environment.Value("EnvTestAutomationRootPath")

pos = InStr(Environment.Value("TestDir"),TestAutomationRootPath)
g_TestSuitePath = Mid(Environment.Value("TestDir"),1,pos+Len(TestAutomationRootPath))
ExecuteFile(g_TestSuitePath&"Resource\Init.vbs")
InitResult = Init()
if(InitResult = 1) then
  Reporter.ReportEvent micFail, "VBS Initialized Failed", "Failed to initialize VBS Scripts"
end if
g_strdatafile = g_TestSuitePath&"TestSuite\MasterScript\QTTrader_MasterDriver.xls"
datatable.AddSheet("Configuration")
datatable.ImportSheet g_strdatafile,"Configuration","Configuration"
datatable.ImportSheet g_strdatafile,"MasterDriver","Action1"

DriverRows = datatable.GetSheet("Action1").GetRowCount
if DriverRows = 0 then 
  Reporter.ReportEvent micFail, "Master Driver Return Zero Records", "Master Driver - There is no records in Master Driver Table"
  ExitRun(0)
end if 
datatable.GetSheet("Configuration").SetCurrentRow(1)
g_Timeout = datatable.Value("Timeout","Configuration")

For Counter = 1 to DriverRows
   datatable.GetSheet("Action1").SetCurrentRow(Counter)
   if(trim(Ucase(datatable.Value("TobeExecuted","Action1"))) = "YES") then
      g_TestScenario = trim(datatable.Value("TestScenario","Action1"))
      g_Script = trim(datatable.Value("Script","Action1"))
      g_Batch = trim(datatable.Value("Batch","Action1"))
      g_TestDataFile = trim(datatable.Value("TestDataFile","Action1"))
      g_strApplicationName = trim(datatable.Value("Application","Action1"))
      DataFile = g_TestSuitePath&"DataFile\"&g_TestDataFile
      datatable.ImportSheet DataFile,"Global","Global"
      ExecuteFile(g_TestSuitePath&"TestSuite\DriverScripts\"&g_Script&".vbs")
      script = g_Script&"()"
      g_Result = eval(script)
    end if
 Next

 Exittest
