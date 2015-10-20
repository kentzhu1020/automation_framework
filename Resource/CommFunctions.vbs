'*********************************************************************************************************
'
'
'FunctionName		:		FW_GetRowNumber()
'Author Name		:		Kent Zhu
'Description		:		Set the row in running time.
'
'
'
'*********************************************************************************************************
Function FW_GetRowNumber()
  FunctionName = "FW_GetRowNumber"
  Reporter.ReportEvent micDone, FunctionName, "Started"
  Dim index 
  Rows = datatable.GetRowCount
  If Rows < 1 Then
    index = -1
  End If
  For i = 1 to Rows
    datatable.SetCurrentRow(i)
    If Trim(datatable.Value("TestScenario")) = Trim(g_TestScenario) Then
      index = i
      Exit For 
    End If
  Next
  FW_GetRowNumber = index
  Reporter.ReportEvent micDone, FunctionName, "Ended"
End Function 
'*********************************************************************************************************
'
'
'FunctionName		:		FW_GetFormatTime()
'Author Name		:		Kent Zhu
'Description		:		Format the time as MonthName_Day_Year and name the result file.
'
'
'
'*********************************************************************************************************
Function FW_GetFormatTime(MyDate)
  Dim YY
  Dim MM
  Dim DD
  Dim Nowstr
  Currentdatetime = CDate(MyDate)
  YY = Year(Currentdatetime)
  MM = Month(Currentdatetime)
  MM = MonthName(MM,True)
  DD = Day(Currentdatetime)
   If DD<10 Then
	 DD="0"&DD
   End If
  Nowstr=MM&"_"&DD&"_"&YY
  FW_GetFormatTime = Nowstr
End Function
'*********************************************************************************************************
'
'
'FunctionName		:		FW_CreateDataKey()
'Author Name		:		Kent Zhu
'Description		:		Create an dictionary object to save the input data.
'
'
'
'*********************************************************************************************************

Function FW_CreateDataKey()
	FunctionName = "FW_CreateDataKey"
	Reporter.ReportEvent micDone, FunctionName, "Started"

	Dim Data
	Set Data = CreateObject("Scripting.Dictionary")
	datatable.SetCurrentRow(g_Row)
	cols = DataTable.GetSheet("Global").GetParameterCount
	For i = 1 To cols
	Set FiledKeys= datatable.GetSheet("Global").GetParameter(i)
	FieldName = FiledKeys.Name
	FieldValue = datatable.Value(FieldName)
	If(IsNull(FieldValue)) Then
	   Data.Add FieldName,""
	Else
	   Data.Add FieldName,FieldValue
	End If
	Next

	Set FW_CreateDataKey = Data
	Reporter.ReportEvent micDone, FunctionName, "Ended"
End Function

'*********************************************************************************************************
'
'
'FunctionName		:		FW_LogFinalResult()
'Author Name		:		Kent Zhu
'Description		:		Create an dictionary object to save the input data.
'
'
'
'*********************************************************************************************************
Function FW_LogFinalResult(strResult,Data,str)
 FunctionName = "FW_LogFinalResult"
 Reporter.ReportEvent micDone, FunctionName,"Started"
 Dim currentRow 

 Set objFso = CreateObject("Scripting.FileSystemObject")
 g_strTemplateFilePath = g_TestSuitePath&"Results\Template\Template.xls"
 g_strFilePath = g_TestSuitePath&"Results\"&g_strApplicationName&"_"&g_Batch&"_"&g_intStartTime&".xls"
 datatable.AddSheet("IO")
 datatable.AddSheet("Results")
  
 If Not objFso.FileExists(g_strFilePath) Then
     currentRow  = 1
	 datatable.ImportSheet g_strTemplateFilePath, "IO","IO"
	 datatable.ImportSheet g_strTemplateFilePath,"Results","Results"
	 
	 If str= "FinalResult" Then
	 	datatable.GetSheet("Results").SetCurrentRow(currentRow)
	 	datatable.Value("TC_NO","Results") = g_TestDesc
	 	If(strResult = 0) Then
	    	datatable.Value("RESULTS","Results") = "Passed"
	    Else
	        datatable.Value("RESULTS","Results") = "Failed"
	 	End If
	    datatable.Value("REASON","Results") = g_Reason 'Need to get from verification
	    datatable.Value("SEQ_NO","Results") = g_Seq_No
	    datatable.Value("SUB_SEQ_NO","Results") = g_Sub_Seq_No
	    datatable.Value("STARTTIME","Results") = g_startTime
        datatable.Value("ENDTIME","Results") = g_endTime 		
	 Else
	   datatable.GetSheet("IO").SetCurrentRow(currentRow)
	   datatable.Value("TestCase","IO") = g_TestDesc
	   datatable.Value("Seq_NO","IO") = Data.Item("Seq_NO")
	   datatable.Value("Sub_Seq_NO","IO") = Data.Item("Sub_Seq_NO") 
	   datatable.Value("IO","IO") = g_IO
	   datatable.Value("Order_ID","IO")=Data.Item("Order_ID")
	   datatable.Value("Price","IO") = Data.Item("Price")
	   datatable.Value("Amount","IO") = Data.Item("Amount")
	   datatable.Value("Status","IO") = Data.Item("Status")
	   datatable.Value("Type","IO") = Data.Item("Type")
	   datatable.Value("Total","IO") = Data.Item("Total")
	   datatable.Value("Order_ID","IO") = Data.Item("Order_ID")
	 End If
	  
 Else
    datatable.ImportSheet g_strFilePath,"IO","IO"
    datatable.ImportSheet g_strFilePath,"Results","Results"
    
     If str= "FinalResult" Then
        row = datatable.GetSheet("Results").GetRowCount
        rowIndex = row + 1
       
	 	datatable.GetSheet("Results").SetCurrentRow(rowIndex)
	 	datatable.Value("TC_NO","Results") = g_TestDesc
	 	If(strResult = 0) Then
	    	datatable.Value("RESULTS","Results") = "Passed"
	    Else
	        datatable.Value("RESULTS","Results") = "Failed"
	 	End If
		datatable.Value("REASON","Results") = g_Reason 'Need to get from verifucation
	    datatable.Value("SEQ_NO","Results") = g_Seq_No
	    datatable.Value("SUB_SEQ_NO","Results") = g_Sub_Seq_No
	    datatable.Value("STARTTIME","Results") = g_startTime 
		datatable.Value("ENDTIME","Results") = g_endTime 
	    
	 Else
	   row = datatable.GetSheet("IO").GetRowCount
	   rowIndex = row + 1
       datatable.GetSheet("IO").SetCurrentRow(rowIndex)
       datatable.Value("TestCase","IO") = g_TestDesc
	   datatable.Value("Seq_NO","IO") = Data.Item("Seq_NO")
	   datatable.Value("Sub_Seq_NO","IO") = Data.Item("Sub_Seq_NO") 
	   datatable.Value("IO","IO") = g_IO
	   datatable.Value("Order_ID","IO")=Data.Item("Order_ID")
	   datatable.Value("Price","IO") = Data.Item("Price")
	   datatable.Value("Amount","IO") = Data.Item("Amount")
	   datatable.Value("Status","IO") = Data.Item("Status")
	   datatable.Value("Type","IO") = Data.Item("Type")
	   datatable.Value("Total","IO") = Data.Item("Total")  
	 End If
    
 End If
 
  datatable.ExportSheet g_strFilePath,"IO"
  datatable.ExportSheet g_strFilePath,"Results"
  
 Reporter.ReportEvent micDone, FunctionName,"Ended"
End Function

'*********************************************************************************************************
'
'
'FunctionName		:		RegExpReference()
'Author Name		:		Kent Zhu
'Description		:		Using regular expression to fetch out the digital numbers
'
'
'
'*********************************************************************************************************
Function RegExpReference(patrn,strng)
  Dim regEx, Match, Matches
  Set regEx = New RegExp
  regEx.Pattern=patrn   'Set the Regular Expression for the function. 
  regEx.IgnoreCase= True
  regEx.Global= True
  Set Matches =regEx.Execute(strng)  'Search in the string.
  For Each Match in Matches
   RetStr=RetStr & Match.Value & "&"
  Next
  RegExpReference=RetStr 
End Function

'*********************************************************************************************************
'
'
'FunctionName		:		ClearVariables()
'Author Name		:		Kent Zhu
'Description		:		Clearing all the global variables which is no use.
'
'
'
'*********************************************************************************************************
Function ClearVariables()
  FunctionName = "ClearVariables"
  Reporter.ReportEvent micDone, FunctionName, "Started" 
  g_ConfirmationLabel = " "
  g_Reason = " "
  g_IO = " "
  g_startTime = " "
  Set g_OutputData = Nothing
  Reporter.ReportEvent micDone, FunctionName, "Ended"
End Function

'*********************************************************************************************************
'
'
'FunctionName		:		FormatResults()
'Author Name		:		Kent Zhu
'Description		:		For Format The Results Sheet  --Green is represent Pass, Red is for Fail.
'
'
'
'*********************************************************************************************************
Function FormatResults()
    Dim index
    'g_strFilePath = "E:\Auto_Testing\QTP\test01.xls"
	Set excelApp = CreateObject("Excel.Application")
	Set excelBook = excelApp.Workbooks.Open(g_strFilePath)
	Set excelSheet = excelBook.Worksheets("Results")
	
	intRowsCount = excelBook.ActiveSheet.UsedRange.Rows.Count
	intColsCount = excelBook.ActiveSheet.UsedRange.Columns.Count
    Err.Clear
    On Error Resume Next 
	For col = 1 To intColsCount
	  If Ucase(excelSheet.Cells(1,col).Value) = "RESULTS" Then
	  	index = col
	  	Exit For
	  End If	  	
    Next
    
   For Row = 1 To intRowsCount
      If Ucase(excelSheet.Cells(Row,index).Value) = "PASSED" Then
         excelSheet.Cells(Row,index).Font.Color = vbGreen
      ElseIf Ucase(excelSheet.Cells(Row,index).Value) = "FAILED" Then
         excelSheet.Cells(Row,index).Font.Color = vbRed
      End If
   Next 
   excelApp.DisplayAlerts = False
   excelApp.ActiveWorkbook.Save
   excelApp.Quit
   Set excelSheet = Nothing
   Set excelBook = Nothing
   Set excelApp = Nothing
  

End Function

'*********************************************************************************************************
'
'
'FunctionName		:		CreateDebugFile()
'Author Name		:		Kent Zhu
'Description		:		For Creating the text file.
'
'
'
'*********************************************************************************************************

Function CreateDebugFile()
	gstrLocalAutomationPath = FN_RELATIVE_PATH(FN_RELATIVE_PATH(Environment.Value("ENV_LOGFILE")))
	gstrDebugFilePath  = gstrLocalAutomationPath

	'currenttest should be the test case name

	currenttest = Environment.Value("TEST_CASE_NAME")

	gstrDebugLogFile = gstrDebugFilePath &  "Log_"& AutomationDate() & "_" & currenttest & ".log"


	if (gobjFileSystem.FileExists(gstrDebugFilePath &  "Log_"& AutomationDate() & "_" & currenttest & ".log" )) Then

	'	Set gobjDebugLogFile = gobjFileSystem.CreateTextFile(gstrDebugFilePath & "Log_"& AutomationDate() & "_" & currenttest & ".log" ,2,false) ' 2 - For Writing
	'	gobjDebugLogFile.close
		Set gobjDebugLogFile = gobjFileSystem.OpenTextFile(gstrDebugFilePath & "Log_"& AutomationDate() & "_" & currenttest & ".log",8, false) ' 8 - Appending constant
		Debug "Info","\n"
	Else

		 Set gobjDebugLogFile = gobjFileSystem.CreateTextFile(gstrDebugFilePath & "Log_"& AutomationDate() & "_" & currenttest & ".log",2,false) ' 2 - For Writing
		 
	End If

	Debug "Info","Log file created!"
End Function
'----------------------------------------------------------------------------------------------------------
'
'
'FunctionName		:		CreateDebugFile()
'Author Name		:		Kent Zhu
'Description		:		This is a simple debug method which will create debug file with todays date
'
'
'
'----------------------------------------------------------------------------------------------------------

'fstart:Debug
Sub Debug (strType,strText)
   If Trim(strText) = "\n"  Then
   gobjDebugLogFile.WriteLine(vbNewLine)
	ElseIf gDebugOn and not Trim(strText) = "" Then
	   gobjDebugLogFile.WriteLine(Now & " | " & strType  & " | " & Trim(strText))
	End If
End Sub
'fend:Debug
