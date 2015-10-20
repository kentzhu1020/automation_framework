Function PlaceAnOrder(ByVal InputData)
Dim returnValue,username,APIlag,Product,orderType
returnValue = 0
username = ""
APIlag = ""
Product = Ucase(trim(InputData.Item("Product")))
orderType = Ucase(trim(InputData.Item("Type")))
FunctionName = "PlaceAnOrder"
Reporter.ReportEvent micDone, FunctionName, "Started"

 Window("BTCChinaTrader").Activate
 Window("BTCChinaTrader").Maximize
 If Window("BTCChinaTrader").WinObject("APILag").Exist Then
 	returnValue = 1
    PlaceAnOrder = returnValue
    g_Reason = g_Reason&"API CONNECTION IS DOWN..."
	Reporter.ReportEvent micDone, FunctionName, "Ended"
	Exit Function
 Else
  username = Window("BTCChinaTrader").WinObject("accountLoginLabel").GetVisibleText
  APIlag = Window("BTCChinaTrader").WinObject("API").GetVisibleText
 End If

 symbol = trim(Window("BTCChinaTrader").WinObject("marketSymbol").GetVisibleText)
 symbol = Split(symbol," ")
 currentMarket = symbol(0)
 
 If trim(Lcase(username)) = "offline" OR trim(Lcase(APIlag)) = "api is down" Then
    returnValue = 1
    PlaceAnOrder = returnValue
    g_Reason = g_Reason&"API CONNECTION IS DOWN..."
	Reporter.ReportEvent micDone, FunctionName, "Ended"
	Exit Function
 Else  
 
      If currentMarket = "BTC/CHY" Then
         If Product = "LTC" Then
          Window("BTCChinaTrader").WinObject("marketSymbol").Click
          Call SELECT_LTC_MARKET_FROM_BTC() 
         ElseIf Product = "BTCLTC" Then
         Window("BTCChinaTrader").WinObject("marketSymbol").Click
         Call SELECT_BTCLTC_MARKET_FROM_BTC()
         End If
      ElseIf currentMarket = "LTC/CHY" Then
          If Product = "BTC" Then
          Window("BTCChinaTrader").WinObject("marketSymbol").Click
          Call SELECT_BTC_MARKET_FROM_LTC() 
         ElseIf Product = "BTCLTC" Then
         Window("BTCChinaTrader").WinObject("marketSymbol").Click
         Call SELECT_BTCLTC_MARKET_FROM_LTC()
         End If
      ElseIf currentMarket = "LTC/BTC" Then
          If Product = "BTC" Then
          Window("BTCChinaTrader").WinObject("marketSymbol").Click
          Call SELECT_BTC_MARKET_FROM_BTCLTC() 
         ElseIf Product = "LTC" Then
         Window("BTCChinaTrader").WinObject("marketSymbol").Click
         Call SELECT_LTC_MARKET_FROM_BTCLTC()
         End If
      End If
      
	 Err.Clear
	 On Error Resume Next
	 If Window("BTCChinaTrader").WinObject("orderIdentity").Exist Then
	    If Window("BTCChinaTrader").WinObject("tableOpenOrders").Exist then
		   If Err.Number <> 0 Then 
		        returnValue = 1
				PlaceAnOrder = returnValue
				g_Reason = g_Reason&Err.Description
				Reporter.ReportEvent micDone, FunctionName, "Ended"
				Exit Function
				
		    Else 
			  orders= Window("BTCChinaTrader").WinObject("tableOpenOrders").GetVisibleText
			  Msgbox orders
			  orders = Split(trim(orders),Chr(13)&Chr(10),-1,1)
			  orderNo = Ubound(orders)
		
				If orderNo > 7 Then
					Window("BTCChinaTrader").WinObject("cancel_All_Orders").Click
					Set objOrders= Window("BTCChinaTrader").WinObject("tableOpenOrders")
					returnValue = returnValue + waitUntilPanelCleared(objOrders,g_Timeout)
				End If
			End If 
	    End If
	 End If
	 
	 
	 If returnValue = 0 Then
	   If orderType = "BID" Then 
   		Window("BTCChinaTrader").WinObject("buyBTC").Click
        Call typeTextBox(InputData.Item("Amount"))
        wait(1)
        Window("BTCChinaTrader").WinObject("buyPrice").Click
        Call typeTextBox(InputData.Item("Price"))
        wait(1)
        Window("BTCChinaTrader").WinObject("btnBuy").Click

       Set Obj =  Window("BTCChinaTrader").Window("Confirm_Buy")
       returnValue = returnValue + waitObject(Obj,g_Timeout) 
	   If returnValue = 0 then
		   g_ConfirmationLabel = Window("BTCChinaTrader").Window("Confirm_Buy").WinObject("msgboxLabel").GetVisibleText 
		   Window("BTCChinaTrader").Window("Confirm_Buy").WinObject("btnYes").Click
		   Set openOrdersObj = Window("BTCChinaTrader").WinObject("tableOpenOrders")
		   returnValue = returnValue+ waitObject(openOrdersObj,g_Timeout)
		   If returnValue = 0 then 
		     g_OrderDetails = Window("BTCChinaTrader").WinObject("tableOpenOrders").GetVisibleText 
			 Msgbox g_OrderDetails
		   Else
			returnValue = 1
            PlaceAnOrder = returnValue
			'g_Reason = g_Reason & "Until "&Chr(32)&g_Timeout&Chr(32)&",Could Not Place An Order!..."
	        Reporter.ReportEvent micDone, FunctionName, "Ended"
	        Exit Function
		   End If 
		Else
		returnValue = 1
        PlaceAnOrder = returnValue
		'g_Reason = g_Reason & "Until "&Chr(32)&g_Timeout&Chr(32)&",Could Not Place An Order!..."
	    Reporter.ReportEvent micDone, FunctionName, "Ended"
	    Exit Function
	  End If
     ElseIf orderType = "ASK" Then
  
      Window("BTCChinaTrader").WinObject("sellBTC").Click
      Call typeTextBox(InputData.Item("Amount"))
      wait(1)
      Window("BTCChinaTrader").WinObject("sellPRICE").Click
      Call typeTextBox(InputData.Item("Price"))
      wait(2)
      Window("BTCChinaTrader").WinObject("btnSELL").Click
      Set Obj =  Window("BTCChinaTrader").Window("Confirm_Sell")
      returnValue = returnValue + waitObject(Obj,g_Timeout)
      If returnValue = 0 then
		   g_ConfirmationLabel = Window("BTCChinaTrader").Window("Confirm_Sell").WinObject("msgboxLabel").GetVisibleText 
		   Window("BTCChinaTrader").Window("Confirm_Sell").WinObject("btnYes").Click
		   Set openOrdersObj = Window("BTCChinaTrader").WinObject("tableOpenOrders")
		   returnValue = returnValue+ waitObject(openOrdersObj,g_Timeout)
		   If returnValue = 0 then 
		     g_OrderDetails = Window("BTCChinaTrader").WinObject("tableOpenOrders").GetVisibleText 
			 Msgbox g_OrderDetails
		   Else
			returnValue = 1
            PlaceAnOrder = returnValue
	        Reporter.ReportEvent micDone, FunctionName, "Ended"
	        Exit Function
		   End If 
		Else
		returnValue = 1
        PlaceAnOrder = returnValue
		'g_Reason = g_Reason & "Until "&Chr(32)&g_Timeout&Chr(32)&",Could Not Place An Order!..."
	    Reporter.ReportEvent micDone, FunctionName, "Ended"
	    Exit Function
	  End If
    End If
  End if
End If
    
PlaceAnOrder = returnValue
Reporter.ReportEvent micDone, FunctionName, "Ended" 
End Function

Function FW_GetConfirmationBoxValue()

  FunctionName = "FW_GetConfirmationBoxValue"
  Reporter.ReportEvent micDone, FunctionName, "Started"
  Dim ODic
  Set ODic = CreateObject("Scripting.Dictionary")
  values = RegExpReference("\d*\.\d*",g_ConfirmationLabel)
  values =Split(values,"&")
  ODic.Add "Amount",values(0)
  ODic.Add "Price",values(1)
  Set FW_GetConfirmationBoxValue = ODic
  Reporter.ReportEvent micDone, FunctionName, "Ended" 
End Function

Function FW_GetOrderDetails()
  FunctionName = "FW_GetOrderDetails"
  Reporter.ReportEvent micDone, FunctionName, "Started"
  Dim ODic
  Set ODic = CreateObject("Scripting.Dictionary")

  g_OrderDetails =Split(trim(g_OrderDetails),Chr(13)&Chr(10),-1,1)
  If IsArray(g_OrderDetails) Then
     middleCount = UBound(g_OrderDetails)/2
     For i = 0 To middleCount-1
      If g_OrderDetails(i) = "#" Then
         g_OrderDetails(i) = "Order_ID"
      ElseIf Instr(1,g_OrderDetails(middleCount+i),"гд",1) > 0 Then
         g_OrderDetails(middleCount+i) = Right(g_OrderDetails(middleCount+i),Len(g_OrderDetails(middleCount+i))-1)
	  ElseIf (InStr(1,g_OrderDetails(middleCount+i),"B",1)>0 And InStr(1,g_OrderDetails(middleCount+i),".",1)>0)Then
	     g_OrderDetails(middleCount+i) = Right(g_OrderDetails(middleCount+i),Len(g_OrderDetails(middleCount+i))-1)
      End If 
      ODic.Add g_OrderDetails(i),g_OrderDetails(middleCount+i)     
    Next
  End If
  Set FW_GetOrderDetails = ODic
  g_OrderDetails = " "
  Reporter.ReportEvent micDone, FunctionName, "Ended" 
End Function
Function VerifyConfirmationBox(InputData)
  Dim returnValue,ComfirmationBoxValue
  Set ConfirmationBoxValue = FW_GetConfirmationBoxValue
  returnValue = 0
  Set g_OutputData = CreateObject("Scripting.Dictionary")
  FunctionName = "VerifyConfirmationBox"
  Reporter.ReportEvent micDone, FunctionName, "Started"
  
  returnValue =  strComp(CStr(CDbl(ConfirmationBoxValue.Item("Price"))*1.0000),CStr(CDbl(InputData.Item("Price"))*1.0000),1)
  If returnValue <> 0  Then
    returnValue = 2
  	g_Reason = g_Reason&"Price_In Confirmation Box"&"&"
  	Reporter.ReportEvent micFail,"Price In Confirmation Box Is Incorrect","Expected Price is:"&InputData.Item("Price")&",Actual It's"&ConfirmationBoxValue.Item("Price")
  Else
    Reporter.ReportEvent micPass,"Price In Confirmation Box","Passed"
  End If
  returnValue = returnValue + strComp(CStr(CDbl(ConfirmationBoxValue.Item("Amount"))*1.0000),CStr(CDbl(InputData.Item("Amount"))*1.0000),1)
   If returnValue <> 0 Then
  	g_Reason = g_Reason&"Amount_In Confirmation Box"&"&"
  	Reporter.ReportEvent micFail,"Amount in Confirmation Box Is Incorrect","Expected Amount is:"&InputData.Item("Amount")&",Actual It's"&ConfirmationBoxValue.Item("Amount")
   Else
    Reporter.ReportEvent micPass,"Amount In Confirmation Box","Passed"
   End If
   If returnValue <> 0 Then
  	returnValue = 1
  End If
  VerifyConfirmationBox = returnValue 
  Reporter.ReportEvent micDone, FunctionName, "Ended" 
	
End Function

Function VerifyOrderDetails(InputData)
  Dim returnValue,ComfirmationBoxValue
  Set VerifyedData = FW_GetOrderDetails
  returnValue = 0
  FunctionName = "verifyOrderDetails"
  Reporter.ReportEvent micDone, FunctionName, "Started"
  g_IO = "Actual Data In Order Panel"
  FW_LogFinalResult returnValue, VerifyedData,"Act"
  
  returnValue = returnValue + strCompare(VerifyedData,InputData,"Type")
  returnValue = returnValue + numCompare(VerifyedData,InputData,"Amount")
  returnValue = returnValue + numCompare(VerifyedData,InputData,"Price")
  
  If returnValue <> 0 Then
  	returnValue =1
  End If
  verifyOrderDetails = returnValue
  Reporter.ReportEvent micDone, FunctionName, "Ended" 
  
End Function

Function strCompare(VerifiedData,InputData,FieldName)
  Dim returnValue
  returnValue = 0
  FunctionName = "strCompare"
  
  Reporter.ReportEvent micDone, FunctionName, "Started"
  returnValue = strComp(VerifiedData.Item(FieldName),InputData.Item(FieldName),1)
  g_OutputData.Add FieldName,VerifiedData.Item(FieldName)
  If returnValue <> 0 Then
     returnValue =2
     g_Reason = g_Reason&FieldName&"&"
     Reporter.ReportEvent micFail,FieldName&Chr(32)&"Is Incorrect","Expected Data For"&Chr(32)&FieldName&Chr(32)&"Is:"&Chr(32)&InputData.Item(FieldName)&Chr(32)&", Actual It's"&Chr(32)&VerifiedData.Item(FieldName)
  Else
      Reporter.ReportEvent micPass,FieldName,"Passed"
  End If
   If returnValue <> 0 Then
     returnValue = 1
  End If
  strCompare = returnValue
  Reporter.ReportEvent micDone, FunctionName, "Ended"
End Function
Function  numCompare(VerifiedData,InputData,FieldName) '9/16
  Dim returnValue
  returnValue = 0
  FunctionName = "numCompare"
  
  Reporter.ReportEvent micDone, FunctionName, "Started"
  returnValue = strComp(CStr(CDbl(VerifiedData.Item(FieldName))*1.0000),CStr(CDbl(InputData.Item(FieldName))*1.0000),1)
  g_OutputData.Add FieldName,VerifiedData.Item(FieldName)
  If returnValue <> 0 Then
     returnValue =2
     g_Reason = g_Reason&FieldName&"&"
     Reporter.ReportEvent micFail,FieldName&Chr(32)&"Is Incorrect","Expected Data For"&Chr(32)&FieldName&Chr(32)&"Is:"&Chr(32)&InputData.Item(FieldName)&Chr(32)&", Actual It's"&Chr(32)&VerifiedData.Item(FieldName)
  Else
      Reporter.ReportEvent micPass,FieldName,"Passed"
  End If
   If returnValue <> 0 Then
     returnValue = 1
  End If
  numCompare = returnValue
  Reporter.ReportEvent micDone, FunctionName, "Ended"
End Function
Function SELECT_LTC_MARKET_FROM_BTC()
 Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{DOWN}" 
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing
End Function

Function SELECT_BTCLTC_MARKET_FROM_BTC()
 Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{DOWN}" 
 WshShell.SendKeys "{DOWN}" 
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing

End Function

Function SELECT_BTC_MARKET_FROM_LTC
Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{UP}" 
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing
End Function

Function SELECT_BTCLTC_MARKET_FROM_LTC
 Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{DOWN}" 
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing
	
End Function

Function SELECT_BTC_MARKET_FROM_BTCLTC
 Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{UP}"
 WshShell.SendKeys "{UP}" 
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing
End Function

Function SELECT_LTC_MARKET_FROM_BTCLTC
 Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{UP}"
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing
End Function

Function typeTextBox(text)
 Dim WshShell  
 Set WshShell = CreateObject("WScript.Shell")  
 WshShell.SendKeys "{UP}"
 WshShell.SendKeys text
 WshShell.SendKeys "{ENTER}" 
 Set WshShell = Nothing
End Function

Function waitUntilPanelCleared(Obj,timeout)
Dim retrunValue,order
returnValue = 1
FunctionName = "waitUntilPanelCleared"
Reporter.ReportEvent micDone, FunctionName, "Started" 
orderNo = Obj.GetVisibleText
orderNo = Split(trim(orderNo),Chr(13)&Chr(10),-1,1)
order = Ubound(orderNo)
 For  i= 1 To timeout
	If order < 7 Then
	  returnValue = 0  
	  waitUntilPanelCleared = returnValue
	  Reporter.ReportEvent micDone, FunctionName, "Ended"   
	  Exit Function
	Else
      wait(1)
      If Not Obj.Exist Then
        order = 0
      Else
       orderNo = Obj.GetVisibleText
       orderNo = Split(trim(orderNo),Chr(13)&Chr(10),-1,1)
       order = Ubound(orderNo)    
      End If      
	End If
 Next

 g_Reason = g_Reason & "Until "&Chr(32)&timeout&Chr(32)&",It could not cancel orders!..." 
 waitUntilPanelCleared =returnValue
 Reporter.ReportEvent micDone, FunctionName, "Ended" 
End Function

Function waitObject(Obj,timeout)
Dim retrunValue
returnValue = 1
FunctionName = "waitObject"
Reporter.ReportEvent micDone, FunctionName, "Started" 
	For  i= 1 To timeout
		If Obj.Exist Then
		  returnValue =0  
		  waitObject =returnValue
		  Reporter.ReportEvent micDone, FunctionName, "Ended"  
		  Exit Function
		Else
          wait(1)          
		End If
    Next
 g_Reason = g_Reason & "Until "&Chr(32)&timeout&Chr(32)&",Could Not Place An Order!..."
 waitObject =returnValue
 Reporter.ReportEvent micDone, FunctionName, "Ended" 
End Function

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