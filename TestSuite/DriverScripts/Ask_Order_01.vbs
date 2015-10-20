'*************************************************************************************************************************************************
'Driver Script Name  :  Ask_Order_01.vbs
'DESCRIPTION:
'    This VBS file will be invoked by the Master Script to execute the test case containing flow to bid orders with valid data and verify the error.

'ASSUMPTIONS:
'   None

'REVISION HISTORY
'Author:
'Create On Date:
'Change History:
'Date                    Name                  Modification Details
'***************************************************************************************************************************************************
Public function Ask_Order_01()

  FunctionName = "Ask_Order_01"
  Reporter.ReportEvent micDone, FunctionName, "Started"
  g_Row = FW_GetRowNumber()

  if g_Row < 0 then
     Reporter.ReportEvent micFail, "Input Data Check", "No Available Records in InputData"
     Exit function
  else	
     datatable.SetCurrentRow(g_Row)
  End if

  Do while ((Trim(datatable.Value("TestScenario")) = Trim(g_TestScenario)))
     g_Result = 0
     datatable.SetCurrentRow(g_row) 
     g_intStartTime = FW_GetFormatTime(Now)
     g_Seq_No = datatable.Value("Seq_NO")
     g_Sub_Seq_No = datatable.Value("Sub_Seq_NO")
     g_TestDesc = datatable.Value("TestDesc")
     Reporter.ReportEvent micDone, g_TestDesc & "** Script **"&FunctionName, "Started"
	 g_startTime = Now
     Set InputData = FW_CreateDataKey()
     g_Result = g_Result + PlaceAnOrder(InputData)
     g_IO = "Expected Data In Ask BTC Order"
     FW_LogFinalResult g_Result,InputData,"exp"
     
     if g_Result = 0 then 
	     g_Result = g_Result + VerifyConfirmationBox(InputData)
    	 g_Result = g_Result + VerifyOrderDetails(InputData)
     else
        g_Result =1
     end if
     g_endTime = Now
	
     if(g_Result >=1) then
         FW_LogFinalResult 1,g_OutputData,"FinalResult"
     else
        FW_LogFinalResult 0,g_OutputData,"FinalResult"
     end If 

     g_Row = g_Row + 1
     datatable.SetCurrentRow(g_Row)
     Set InputData = Nothing
     Reporter.ReportEvent micDone, g_TestDesc & "** Script **"&FunctionName, "End"
     ClearVariables()
   Loop
   FormatResults()
   Reporter.ReportEvent micDone, FunctionName, "End"
End function





 

























