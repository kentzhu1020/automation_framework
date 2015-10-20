'*************************************************************************************************************************************************
'Driver Script Name  : Init.vbs
'DESCRIPTION:
'    This VBS file will initial all the vbs and global variables.

'ASSUMPTIONS:
'   None

'REVISION HISTORY
'Author:          Kent Zhu 
'Create On Date:
'Change History:
'Date                    Name                  Modification Details
'***************************************************************************************************************************************************
Public Function Init()

  init_result = 0

  Err.Clear
  On Error Resume Next 
  ExecuteFile(g_TestSuitePath&"Resource\GlobalVariables.vbs")
  If Err.number <> 0 Then 
     init_result = init_result + 1
  End If


  Err.Clear
  On Error Resume Next 
  ExecuteFile(g_TestSuitePath&"Resource\CommFunctions.vbs")
  If Err.number <> 0 Then 
     init_result = init_result + 1
  End If

  Err.Clear
  On Error Resume Next 
  ExecuteFile(g_TestSuitePath&"Resource\AppFunctions.vbs")
 
  If Err.number <> 0 Then 
     init_result = init_result + 1
  End If

  If init_result >=1 Then
     init_result = 1
  End If

    Init = init_result

  End Function 


