'###########################################################################################################################
'#
'#	FW_Data_Lib:-		Contains Data Control functions Used by Generic Automation Framework
'#__________________________________________________________________________________________________________________________
'#		KEYWORDS			PARAMETERS
'#__________________________________________________________________________________________________________________________
'#
'#		1.	COMPARE			actValue, expValue
'#		2.	WAIT_			waitTime
'#		3.	IF_				condition, trueCount, falseCount
'#		4.	EXECUTE_		statement
'#		5.	VALIDATE_DB		DBConnection_String, Input_SQL, expValue
'#		6.	EXCELOUTPUT		filePath, fileName, sheetname, colNames
'#		7.	SENDKEY			winDesc, In_Key
'#		8.	EXIT_RUN		-
'#__________________________________________________________________________________________________________________________
'#		COMMON FUNCTIONS:
'#__________________________________________________________________________________________________________________________
'#		1.	MailAlert			-
'#		2.	DBConnect			byRef curSession ,DBConnection_String
'#		3.	createDescription	objectDesc
'#		4.	captureScreen		-
'#		5.	ErrorHandler		methodName
'#		6.	clearCache			-
'#		7.	openExcel			filePath, fileName
'#		8.	closeExcel			filePath, fileName
'#		9.	PopUp_Recovery		-
'#__________________________________________________________________________________________________________________________
'#		HTML REPORT FUNCTIONS:
'#__________________________________________________________________________________________________________________________
'#		1.	HTMLResultSummary()
'#		2.	HTMLStepResults()
'#		3.	HTMLErrorLog()
'###########################################################################################################################

'Option Explicit		'	-	Declare all the variables used

'___________________________________________________________________________________________________________________________
'# Function Name	: COMPARE()
'# Purpose	        : To compare the Actual & Expected Values
'# Parameters 		: actValue		-> Gets the Actual Value
'#                        expValue		-> Gets the Expected Value
'# Return	    	: 0  : Success
'#         	     	 -1  : Failure
'___________________________________________________________________________________________________________________________
Public Function COMPARE(actValue, expValue)
	On Error Resume Next
	Dim methodName, Act_Val, Exp_Val, temp
	methodName = "COMPARE" : COMPARE = 0
	If Step_Description = "" AND Exp_Result = "" Then
		Step_Description = "Compare Actual Value(s) with Expected Values"
		Exp_Result = "Should compare Actual Value(s) -> " & actValue & " with Expected Value(s) -> " & expValue
	End If
	If Exec_Flag = "Y" Then
		Act_Val = trim(actValue) : Exp_Val = trim(expValue)
		temp = StrComp(Act_Val, Exp_Val, 1)
		If temp = 0 Then
			Actual_Res = "Actual & Expected Values are Matching !!" & Chr(13) & "Actual Value :-  " & actValue & Chr(13) & "Expected Value :-  " & expValue
			Reporter.reportevent micInfo,StepName, Actual_Res
		Else
			Actual_Res = "Actual & Expected Values are NOT Matching  !! " & Chr(13) & "Actual Value :-   " & actValue & Chr(13) & "Expected Value :-   " & expValue
			Reporter.ReportEvent micInfo, StepName, Actual_Res
			COMPARE = 1
		End If
	End If
	' Handling Error
	methodName = "COMPARE" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function COMPARE() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: WAIT_()
'# Purpose			: Instructs QTP to wait for fixed time
'# Parameters 		: waitTime		-> Gets the time to be waited
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function WAIT_(waitTime)
	On Error Resume Next
	Dim methodName
	methodName = "WAIT_" : WAIT_ = 0

	'Script to generate Test Step Description and Expected Result
	If Step_Description = "" AND Exp_Result = "" Then
		Step_Description = "Wait for a window to open or an object to appear."
		Exp_Result = "QTP should Wait for " & waitTime & " sec for a window to open or an object to appear."
	End If

	If Exec_Flag = "Y" Then 
		Wait(waitTime)
		Actual_Res = "QTP Waited for " & waitTime & " sec." 
		Reporter.ReportEvent micInfo, StepName, Actual_Res
	End If
'	handle error
	methodName = "WAIT_" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function WAIT_() ----------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: IF_()
'# Purpose			: To check a condition. If the condition is TRUE, continue to Next Step. If the condition is FALSE, skip 'skipCount' steps. 
'# Parameters 		: condition		-> Gets the condition to be checked
'#					  skipCount		-> Gets the no of steps to be skipped for FALSE condition.
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public EXIT_COUNT
Public Function IF_(condition, trueCount, falseCount, usrMsg)
	On Error Resume Next
	Dim methodName, rc
	methodName = "IF_" : IF_ = 0

	'Script to generate Test Step Description and Expected Result
	Step_Description = "If the condition is TRUE, skip -> " & trueCount & " steps and continue execution. Else skip -> " & falseCount & " steps and continue execution."
	Exp_Result = "If the condition is TRUE, should skip -> " & trueCount & " steps and continue execution. Else should skip -> " & falseCount & " steps and continue execution."

	EXIT_COUNT = ""
	If Exec_Flag = "Y" Then		
		rc = Eval(condition)
		If rc = "True" or rc = 0 Then
			If trueCount > 0 Then
				Actual_Res = "Condition is TRUE, Skipping -> " & trueCount & " steps and Continuing execution." 
				Reporter.reportevent micInfo,StepName, Actual_Res
				EXIT_COUNT = trueCount
			ElseIf trueCount = 0 Then
				Actual_Res = "Condition is TRUE, Continuing execution with next step." 
				Reporter.reportevent micInfo,StepName, Actual_Res
			ElseIf trueCount = -1 Then
				If usrMsg = "" Then
					Actual_Res = "Condition is TRUE, Stopping the execution." 
				Else
					Actual_Res = usrMsg
				End If
				Reporter.reportevent micFail,StepName, Actual_Res
				IF_ = -1
			End If
		Else
			If falseCount > 0 Then
				Actual_Res = "Condition is FALSE, Skipping -> " & falseCount & " steps and Continuing execution." 
				Reporter.reportevent micInfo,StepName, Actual_Res
				EXIT_COUNT = falseCount
			ElseIf falseCount = 0 Then
				Actual_Res = "Condition is FALSE, Continuing execution with next step." 
				Reporter.reportevent micInfo,StepName, Actual_Res
			ElseIf falseCount = -1 Then
				If usrMsg = "" Then
					Actual_Res = "Condition is FALSE, Stopping the execution." 
				Else
					Actual_Res = usrMsg
				End If 
				Reporter.reportevent micFail,StepName, Actual_Res
				IF_ = -1
			End If
		End If

'		If UCase(Exit_Flag ) = "Y" Then
'			If rc = "True" or rc = 0 Then
'				Actual_Res = "Condition is TRUE, Skipping -> " & skipCount & " steps and Continuing execution." 
'				Reporter.reportevent micInfo,StepName, Actual_Res
'				EXIT_COUNT = skipCount
'			Else
'				Actual_Res = "Condition is FALSE, Continuing to next Step."
'				Reporter.reportevent micInfo,StepName, Actual_Res
'			End If
'		Else
'			If rc = "True" or rc = 0 Then
'				Actual_Res = "Condition is TRUE, Continuing to next Step."
'				Reporter.reportevent micInfo,StepName, Actual_Res
'			Else
'				Actual_Res = "Condition is FALSE, Skipping -> " & skipCount & " steps and Continuing execution. " & Actual_Res
'				Reporter.reportevent micInfo,StepName, Actual_Res
'				EXIT_COUNT = skipCount
'			End If
'		End If		
	End If
	' handle error
	methodName = "IF_" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function IF_() ------------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: VALIDATE_DB()
'# Purpose			: To get the Actual Value of object(s) in the application
'# Parameters 		: DBConnection_String	-> Gets the DB Connection parameters
'#					  Input_SQL				-> Gets the SQL Query to be executed to get the Actual Value
'#					  expValue				-> Gets the Expected Value(s) to be compared with Actual DB values
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function VALIDATE_DB(DBConnection_String, Input_SQL, expValue)
	On Error Resume Next
	Dim methodName, ExpVal_Array, ExpVal_Count, RC_DBCon, i, rc, ActVal_Count
	methodName = "VALIDATE_DB" : VALIDATE_DB = ""
	
	ExpVal_Array = split(expValue,", ",-1,1) : ExpVal_Count = UBOUND(ExpVal_Array)
	RC_DBCon = DBConnect(curSession, DBConnection_String)
	If RC_DBCon = 0 Then
		Set ActVal_Array = curSession.Execute(Input_SQL)
		Reporter.ReportEvent micInfo, methodName, "Successfully executed the SQL Query :--->  " & Input_SQL
		ActVal_Count = UBOUND(ActVal_Array)
		If ActVal_Count = ExpVal_Count Then
			i = 0
			For i = 0 to ActVal_Count 
				rc = COMPARE(ActVal_Array(i), ExpVal_Array(i)) 
			Next
		Else
			Results_Msg = "Actual Value Count :-  " & ActVal_Count & "    AND Expected Value Count :-  " & ExpVal_Count & "    Doesn't match ???"
			Reporter.reportevent micFail,StepName, Results_Msg
			VALIDATE_DB = -1
		End If
	End If
	' Handling Error
	methodName = "VALIDATE_DB" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function VALIDATE_DB() ----------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: EXCELOUTPUT()
'# Purpose			: Output the values to an Excel sheet.
'# Parameters 		: filePath, fileName, sheetname, colNames, cellValue, flag
'#					  If there are more than one colNames and cellValue, then put values separated by "; "
'#					  colNames = "ColumnName1; ColumnName2" : cellValue = "Value1; Value2"
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function EXCELOUTPUT(filePath, fileName, sheetname, colNames, cellValue, flag)
	On Error Resume Next
	Dim methodName, objExcel, objWrkBook, objWrkSheet, range, rowCount, colCount, col, i, colName, Col_Name, Result_Msg, testcasename, row, exec_indicator, j
'	methodName = "EXCELOUTPUT" : EXCELOUTPUT = ""

	'Script to generate Test Step Description and Expected Result
	Step_Description = "Output the values of -> " & colNames & " in Excel sheet -> " & filePath & fileName	
	Exp_Result = "Should output the values of -> " & colNames & " in Excel sheet -> " & filePath & fileName	

	If Exec_Flag = "Y" Then	
		Reporter.ReportEvent micInfo, methodName,"Opening Excel file -> " & filePath & fileName	
		Set objExcel = CreateObject("Excel.Application")				' create the excel object
		
		Set objWrkBook = objExcel.Workbooks.Open(filePath & fileName)	' open workbook
			Set objWrkSheet = objWrkBook.WorkSheets(sheetname)			' get data worksheet	
		Set range = objWrkSheet.UsedRange								' get data used range
		rowCount = range.Rows.Count : colCount = range.Columns.Count	' get row/column count 
		If rowCount > 0 OR colCount > 0 Then							' check for valid no of rows.
			Col_Name = split(colNames,"; ",-1,1)
			Cell_Value = split(cellValue,"; ",-1,1)
			i = 0
			If flag = 1 Then
				For row = 2 to rowCount
					testcasename = trim(range.cells(row,1)) : exec_indicator = Trim(range.Cells(row,2))
					If UCASE(testcasename) = TESTCASE_NAME and exec_indicator = "Y" Then
						rowcount = row
						Exit For
					End If
				Next
			Else
				rowcount= rowcount+1
			End If
			For col = 1 to colCount
				If i <= UBound(Col_Name) Then
					colName = Trim(range.Cells(1, col))
					If colName = Col_Name(i) Then
						range.Cells(rowCount, col) = Cell_Value(i)	'Trim(eval(colName))
						i = i + 1
					End If
				End If
				'If i <= UBound(Col_Name) Then
					'colName = Trim(range.Cells(1, col))
					'If colName = Col_Name(i) Then
					'	range.Cells(rowCount, col) = Trim(eval(colName))
					'	i = i + 1
					'End If
				'End If
			Next
			Actual_Res = "Outputting the value/s -> " & cellValue & " to column/s -> " & colNames & " into Excel sheet -> " & filePath & fileName 
			'"Outputting the values of -> " & colNames & " into Excel sheet -> " & filePath & fileName 
			Reporter.reportevent micInfo,StepName, Actual_Res
			Else
			Result_Msg = "Excel has invalid no. of rows/columns. Row Count :- " & rowCount & " Column Count :- " & colCount
			Reporter.ReportEvent micFail, methodName, Result_Msg
			EXCELOUTPUT = -1
		End If
		Reporter.ReportEvent micInfo, methodName,"Closing Excel file -> " & filePath & fileName	
		objWrkBook.Save
		objExcel.Workbooks.Close		' close the workbooks
		objExcel.Application.Quit		' clean the objects
		Set objExcel = Nothing : Set objWrkBook = Nothing : Set objWrkSheet = Nothing : Set range = Nothing
	End If
	' Handling Error
	methodName = "EXCELOUTPUT" :	rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function EXCELOUTPUT() ----------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: SendKey(PageDesc, In_Key)
'# Purpose			: Send a key to the active window.
'# Parameters 		:	
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function SENDKEY(PageDesc, In_Key)
	On Error Resume Next
	Dim methodName, Page_Name, pageName, WshShell
	methodName = "SENDKEY" : SENDKEY = 0

	'Script to generate Test Step Description and Expected Result
	If Step_Description = "" AND Exp_Result = "" Then
		Page_Name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
		Step_Description = "Send key -> " & In_Key & " to the active window -> " & pageName
		Exp_Result = "Should send key -> " & In_Key & " to the window -> " & pageName
	End If

	If Exec_Flag = "Y" Then 
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.AppActivate PageDesc, 5 ' Activate the browser window
		WshShell.SendKeys In_Key
		Set WshShell = Nothing
		Actual_Res = "Sending key -> " & In_Key & " to the active window -> " & pageName
		Reporter.ReportEvent micInfo, StepName, Actual_Res
	End If
'  	handle error
	methodName = "SENDKEY" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function SendKey() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: EXIT_RUN()
'# Purpose			: Exit running test case
'# Parameters 		:	
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function EXIT_RUN()
	On Error Resume Next
	Dim methodName
	methodName = "EXIT_RUN" : EXIT_RUN = 0

	'Script to generate Test Step Description and Expected Result
	If Step_Description = "" AND Exp_Result = "" Then
		Step_Description = "Exit running Test Steps."
		Exp_Result = "Should exit running Test Steps."
	End If

	If Exec_Flag = "Y" Then 
		EXIT_FLAG = "Y"
		Actual_Res = "Stopping Test Step execution."
		Reporter.ReportEvent micInfo, StepName, Actual_Res
	End If
'	handle error
	methodName = "EXIT_RUN" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function EXIT_RUN() -------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: HTMLResultSummary
'# Purpose			: Generates Summary results in HTML format
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public htmlSummaryResult
Public Function HTMLResultSummary()

	'If TEST_DATA_INPUT_TYPE = 'Excel' then	

	On Error Resume Next
	Dim methodName, Title, Header_Array, Header_Count, fso, MyFile, objTempFile, objFolder, sFileText, iPos, i, Temp_Value
	methodName = "HTMLResultSummary"

	Title = APPLICATION_NAME & "  -  Automated Test Execution Summary" 
	Header_Array = split(HTML_HEADER,", ",-1,1) : Header_Count = UBOUND(Header_Array)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FolderExists(RESULTS_PATH)Then Set objFolder=fso.createFolder(RESULTS_PATH) End If
	If not fso.FileExists(HTML_RESULT_SUMMARY) Then
		'Create Header of HTML file, if it does not exist already
		Set MyFile = fso.CreateTextFile(HTML_RESULT_SUMMARY, True)
		MyFile.write("<html>")
		MyFile.write("<style>")
		MyFile.write(".subheading { BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;BACKGROUND-COLOR: #CCE77B}")
		MyFile.write(".subheading1{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 10px;}")
		MyFile.Write(".subheading2{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 2px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 2px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 10px;}")
		MyFile.Write(".tdborder_1{BORDER-RIGHT: #bdd0e4 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #bdd0e4 1px solid;PADDING-LEFT: 4px;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #bdd0e4 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #bdd0e4 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Pass{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #41A317;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Fail{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #ff0000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Done{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP:#014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #CC9900;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Skipped{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #B8860B;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".heading {FONT-WEIGHT: bold; FONT-SIZE: 20px; COLOR: #348017;FONT-FAMILY: Arial, Verdana, Tahoma, Arial;}")
		MyFile.Write(".style1 { border: 1px solid #014E07;padding: 0px 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
		MyFile.Write(".style3 { border: 1px solid #014E07;padding: 0px 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 2px;}")
		MyFile.Write("</style>") : MyFile.Write("<head>")
		MyFile.Write("<title>" & Title & "</title>")
		MyFile.Write("</head>") : MyFile.Write("<body>")
		MyFile.Write("<table cellSpacing='0' cellPadding='0' width='96%' border='0'  align='center' style='height: 40px'>")
		MyFile.Write("<tr>")
		MyFile.Write("<td   align=center><span class='heading'>"  & Title  & "</span>")
		MyFile.Write("<br />") : MyFile.Write("<br />")
		MyFile.Write("<FONT SIZE='4' FACE='courier' COLOR=#C11B17><MARQUEE WIDTH=100% BEHAVIOR=SCROLL DIRECTION=LEFT BGColor=white> ... Click on the Test Case Name to get the Step Level Results ... </MARQUEE></FONT> ")
		MyFile.Write("<br />")
		MyFile.Write("</td></tr>")
		MyFile.Write("</table>")
		MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' style='width:96%'; margin-left:'20px;';margin-right:20px;>")
		MyFile.Write("<tr>")
		For i = 0 to Header_Count
			MyFile.Write("<td class='subheading'>" & UCase(Header_Array(i)) & "</td>")
		Next
		MyFile.Write("</tr>")
	Else
'		Setting the position of the file if the file is already created
		Set MyFile = fso.OpenTextFile(HTML_RESULT_SUMMARY,1)
		sFileText = MyFile.readall
		iPos = instr(1,sFileText,"<!--LOGDETAILS-->",vbTextCompare)
		If iPos > 0 Then
			sFileText = mid(sFileText,1,iPos-1)
			MyFile.close
		End If
		Set MyFile = fso.OpenTextFile(HTML_RESULT_SUMMARY, 2)
		MyFile.write sFileText
	End If	
	i = 0
	For i = 0 to Header_Count
		Temp_Value = Trim(Eval(Header_Array(i)))
		If Temp_Value = "" Then Temp_Value = "-" End If
		Select Case Header_Array(i)
			Case "TESTCASE_NAME"
				MyFile.Write("<td class =style1 >")
				MyFile.Write("<a href='" & "file:///" & htmlStepResult &"' style='color: #0000FF' target='blank'>" & TESTCASE_NAME & "</a>")
			Case "STATUS"
				If Temp_Value = "PASS" Then MyFile.Write("<td  class ='tdborder_1_Pass' >" & Temp_Value & "</td>") End If
				If Temp_Value = "FAIL" Then MyFile.Write("<td  class ='tdborder_1_Fail'>" & Temp_Value & "</td>") End If
				If Temp_Value = "ERROR" Then MyFile.Write("<td  class ='tdborder_1_Fail'>" & Temp_Value & "</td>") End If
			Case "STEPS_PASSED"
				MyFile.Write("<td  class ='tdborder_1_Pass' >" & Temp_Value & "</td>")
			Case "STEPS_FAILED"
				MyFile.Write("<td  class ='tdborder_1_Fail' >" & Temp_Value & "</td>")
			Case "STEPS_EXECUTED"
				MyFile.Write("<td  class ='tdborder_1_Skipped' >" & Temp_Value & "</td>")
			Case "ERRORS"
				MyFile.Write("<td  class ='tdborder_1_Fail' >" & Temp_Value & "</td>")
			Case Else
				MyFile.Write("<td class =style1>"& Temp_Value &"</td>")
		End Select	
	Next
	MyFile.Write("</tr>")
	Reporter.ReportEvent micInfo, "HTML_Result", "Created HTML Results Summary in   ===>   " & HTML_RESULT_SUMMARY
	' handle error
	methodName = "HTMLResultSummary" : rc = ErrorHandler(methodName)
'End If

End Function
' --------------------------- End of Function HTMLResultSummary() ----------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: HTMLStepResults
'# Purpose			: Generates step wise results in HTML format
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public htmlStepResult
Public Function HTMLStepResults()
	On Error Resume Next
	Dim methodName, Head_Title, fso, MyFile, objFolder, i, TestStep_Key, Steps_Count, StepDesc_Key, ExpRes_Key, ActRes_Key, ErrorSnap_Key, StepTime_Key
	Dim StepTime_Count, StepFlag_Key, StepFlag_Count, StepName, StepValue, StepDesc, ExpRes, ActRes, StepFlag, Step_Time, StepTime, ErrorSnap, temp
	methodName = "HTMLStepResults"

	htmlStepResult = HTML_STEP_RESULT_PATH & TESTCASE_NAME & "_" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & " " & Hour(Now) & " " & Minute(Now) & " " & Second(Now) & ".htm"
	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FolderExists(HTML_STEP_RESULT_PATH)Then Set objFolder = fso.createFolder(HTML_STEP_RESULT_PATH) End If
    Set MyFile = fso.CreateTextFile(htmlStepResult, True)
    MyFile.write("<html>")
    MyFile.write("<style>")
	MyFile.write(".subheading { BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;BACKGROUND-COLOR: #CCE77B}")
    MyFile.write(".subheading1{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 50px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 10px;}")
	MyFile.Write(".subheading2{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 2px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 2px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 10px;}")
	MyFile.Write(".tdborder_1{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
	MyFile.Write(".tdborder_1_Pass{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #41A317;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
	MyFile.Write(".tdborder_1_Fail{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #ff0000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
	MyFile.Write(".tdborder_1_Skipped{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #736F6E;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
    MyFile.Write(".heading {FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR: #008000;FONT-FAMILY: Arial, Verdana, Tahoma, Arial;}")
	MyFile.Write(".style1 { border: 1px solid #014E07;padding: 0px 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;COLOR: #696969;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
	MyFile.Write(".style2 { border: 1px solid #014E07;padding: 0px 4px;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
	MyFile.Write(".style3 { border: 1px solid #014E07;padding: 0px 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 2px;}")
	MyFile.Write(".style4 { border: 1px solid #014E07;padding: 0px 4px;FONT-SIZE: 9pt;COLOR: #0000CC;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
	MyFile.Write("</style>")
	MyFile.Write("<head>")
	MyFile.Write("<title>" & APPLICATION_NAME & "_Test Step Results</title>")
	MyFile.Write("</head>")
	MyFile.Write("<body>")
	MyFile.Write("<table cellSpacing='0' cellPadding='0' width='96%' border='0'  align='center' style='height: 40px'>")
	MyFile.Write("<tr>")
	Head_Title = " Step Results for Test Case :-  " & TESTCASE_NAME
	MyFile.Write("<td   align=center><span class='heading'><A NAME=Heading>" & Head_Title & "</A> </span>")
	MyFile.Write("<br />")
	MyFile.Write("</td></tr>")
	MyFile.Write("</table>")
	MyFile.Write("</table>")
	If HTML_RESULT_SUMMARY <> "" Then 
		MyFile.Write("<a href='" & "file:///" & HTML_RESULT_SUMMARY &"' style='color: #FF9966' target='blank'>"& "Summary" &"</a>")
		MyFile.Write("<br />")
	End If
	If errLogPath <> "" Then
		MyFile.Write("<a href='" & "file:///" & errLogPath &"' style='color: #FF9966' target='blank'>"& "Error Log" &"</a>")
		MyFile.Write("<br />")
	End If
	MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' style='width:96%; margin-left:20px;'>")
	MyFile.Write("<TR>")
	MyFile.Write("<td class='subheading1' colspan=6 align=center>")
	MyFile.Write("<TR>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Iteration</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >TestCase Status</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Tester</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Total Steps</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Steps - Executed</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Steps - No Run</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Steps - Passed</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Steps - Failed</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Errors</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Exe - Date</TD>")
	MyFile.Write("<TD class=subheading vAlign=center align=middle >Duration</TD>")
	MyFile.Write("<TR>")
	MyFile.Write("<TD class=bg_darkblue height=1></TD>")
	MyFile.Write("<TD class=bg_darkblue  height=1></TD></TR>")
	MyFile.Write("<TR>")
	MyFile.Write("<TD class=bg_gray_eee  height=1></TD>")
	MyFile.Write("<TD class=bg_gray_eee height=1></TD></TR>")
	MyFile.Write("<TR>")
	MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle><b>" & ITERATION & "</b></TD>")
	If STATUS = "PASS" Then
		MyFile.Write("<td  class ='tdborder_1_Pass' vAlign=center align=middle><b>" & STATUS & "</b></td>")
	ElseIf STATUS = "FAIL" then
		MyFile.Write("<td  class ='tdborder_1_Fail' vAlign=center align=middle><b>" & STATUS & "</b></td>")
	ElseIf STATUS = "ERROR" then
		MyFile.Write("<td  class ='tdborder_1_Fail' vAlign=center align=middle><b>" & STATUS & "</b></td>")
	End If
	MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle ><b>" & TESTER & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle><b>" & TOTAL_STEPS & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle ><b>" & STEPS_EXECUTED & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1_Skipped'  vAlign=center align=middle><b>" & TOTAL_STEPS-STEPS_EXECUTED & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1_Pass'  vAlign=center align=middle ><b>" & STEPS_PASSED & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1_Fail'  vAlign=center align=middle ><b>" & STEPS_FAILED & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1_Fail'  vAlign=center align=middle ><b>" & ERRORS & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle ><b>" & Now & "</b></TD>")
	MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle ><b>" & DURATION & "</b></TD>")
	MyFile.Write("</TD></TR>")
  	MyFile.Write("</table>")
	MyFile.Write("<br />")
	MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' margin-left:20px;'>")
    MyFile.Write("<tr>")
    MyFile.Write("<td class='subheading'>Step #</td>")
	MyFile.Write("<td class='subheading'>TestStep_Script</td>")
	MyFile.Write("<td class='subheading'>TestStep_Description</td>")
	MyFile.Write("<td class='subheading'>Expected Result</td>")
    MyFile.Write("<td class='subheading'>Actual Result</td>")
	MyFile.Write("<td class='subheading'>Step_Status</td>")
	MyFile.Write("<td class='subheading'>Step_Duration</td>")
	MyFile.Write("</tr>")

	TestStep_Key = TEST_STEP_DICT.Keys : Steps_Count = UBound(TestStep_Key)
	StepDesc_Key = StepDesc_Dict.Keys : ExpRes_Key = ExpRes_Dict.Keys
	ActRes_Key = ActRes_Dict.Keys : ErrorSnap_Key = StepErSnap_Dict.Keys
	StepTime_Key = StepTime_Dict.Keys : StepTime_Count = UBound(StepTime_Key)
	StepFlag_Key = StepFlag_Dict.Keys : StepFlag_Count = UBound(StepFlag_Key)
	
	For i = 0 to Steps_Count
		MyFile.Write("<tr> ")
		StepName = TestStep_Key(i) : StepValue = TEST_STEP_DICT.Item(StepName) : StepDesc = StepDesc_Dict.Item(StepName)
		ExpRes = ExpRes_Dict.Item(StepName) : ActRes = ActRes_Dict.Item(StepName)
		StepFlag = StepFlag_Dict.Item(StepName) : StepTime = StepTime_Dict.Item(StepName) : ErrorSnap = StepErSnap_Dict.Item(StepName)
		MyFile.Write("<td class =style2>" & StepName & "</td>")
		MyFile.Write("<td class =style2>" & StepValue & "</td>")
		MyFile.Write("<td class =style2>" & StepDesc & "</td>")
		MyFile.Write("<td class =style4>" & ExpRes & "</td>")
		If ActRes = "" Then ActRes = "-" End If
		If StepFlag_Count + 1 > i Then
			If StepFlag = "PASS" Then
				MyFile.Write("<td class =style4>" & ActRes & "</td>")
				MyFile.Write("<td  class ='tdborder_1_Pass' >" & StepFlag & "</td>")
			ElseIf StepFlag = "FAIL" Then
				If ErrorSnap = "" Then
					MyFile.Write("<td class = 'tdborder_1_Fail'>" & ActRes & "</td>")
				Else
					temp = "<a href='" & "file:///" & ErrorSnap &"' style='color: #0000FF' target='blank'>"&".   Click here to see Error Screen"&"</a>"
					MyFile.Write("<td class = 'tdborder_1_Fail'>" & ActRes & temp & "</td>") 
				End If
				MyFile.Write("<td  class ='tdborder_1_Fail'>" & StepFlag & "</td>")
			ElseIf StepFlag = "Skipped" Then
				MyFile.Write("<td class =style2>" & ActRes & "</td>")
				MyFile.Write("<td class ='tdborder_1_Skipped' >" & StepFlag & "</td>")
			ElseIf StepFlag = "ERROR" Then
				temp = "<a href='" & "file:///" & errLogPath &"' style='color: #C11B17' target='blank'>"&".   Click here to see Error Log"&"</a>"
				MyFile.Write("<td class =style4>" & ActRes & temp & "</td>")
				MyFile.Write("<td class ='tdborder_1_Fail' >" & StepFlag & "</td>")
			End If
		Else
			MyFile.Write("<td class =style2>" & ActRes & "</td>")
			MyFile.Write("<td class ='tdborder_1_Skipped' >" & "No Run" & "</td>")
		End If
		If StepTime_Count + 1 > i Then
			Step_Time = StepTime & " Sec"
			If StepTime < 60 Then
				MyFile.Write("<td class =style1>" & Step_Time & "</td>")
			Else
				MyFile.Write("<td  class ='tdborder_1_Fail' >" & Step_Time & "</td>")
			End If
		Else
			Step_Time =  "0 Sec"
			MyFile.Write("<td class =tdborder_1_Skipped >" & Step_Time & "</td>")
		End If
		MyFile.Write("</tr>")
	Next  
	MyFile.Write("</table>") : MyFile.Write("<br />")
	MyFile.Write("</body>") : MyFile.Write("</html>")
	MyFile.Close
	Reporter.ReportEvent micInfo, "HTML_Result", "Created HTML Step Results in   ===>   " & htmlStepResult
	Set fso = Nothing : Set MyFile = Nothing : Set StepDesc_Dict = Nothing : Set ExpRes_Dict = Nothing : Set ActRes_Dict = Nothing
	Set StepTime_Dict = Nothing : Set StepFlag_Dict = Nothing : Set StepErSnap_Dict = Nothing
	' handle error
	methodName = "HTMLStepResults" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function HTMLStepResults() ------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: HTMLStepResults
'# Purpose			: Generates step wise results in HTML format
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public errLogPath, ERRORS
Public Function HTMLErrorLog()
	On Error Resume Next
	Dim methodName, Head_Title, FWErr, FWErr_Array, FWErr_Count, RunErr, RunErr_Array, RunErr_Count, fso, MyFile, objFolder, i, j, S_No
	methodName = "HTMLErrorLog" : ERRORS = 0

	If Run_Error <> "" OR FW_Error <> "" Then
		If TESTCASE_NAME = "" Then TESTCASE_NAME = "NA" End If
		errLogPath	= LOG_PATH & TESTCASE_NAME & "_Log_" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & " " & Hour(Now) & " " & Minute(Now) & " " & Second(Now) & ".htm"
		FWErr_Array = split(FW_Error," / ",-1,1) : FWErr_Count = UBOUND(FWErr_Array) : If FWErr_Count < 0 Then FWErr_Count = 0 End If
		RunErr_Array = split(Run_Error," / ",-1,1) : RunErr_Count = UBOUND(RunErr_Array) : If RunErr_Count < 0 Then RunErr_Count = 0 End If
		Set fso = CreateObject("Scripting.FileSystemObject")
		If not fso.FolderExists(LOG_PATH)Then Set objFolder=fso.createFolder(LOG_PATH) End If
		Set MyFile = fso.CreateTextFile(errLogPath, True)
		MyFile.write("<html>")
		MyFile.write("<style>")
		MyFile.write(".subheading { BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;BACKGROUND-COLOR: #CCE77B}")
		MyFile.write(".subheading1{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 50px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 10px;}")
		MyFile.Write(".subheading2{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 2px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 2px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 10px;}")
		MyFile.Write(".tdborder_1{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #000000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Pass{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #41A317;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Fail{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #ff0000;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".tdborder_1_Skipped{BORDER-RIGHT: #014E07 1px solid;PADDING-RIGHT: 4px;BORDER-TOP: #014E07 1px solid;PADDING-LEFT: 4px;FONT-SIZE: 9pt;PADDING-BOTTOM: 0px;BORDER-LEFT: #014E07 1px solid;COLOR: #736F6E;PADDING-TOP: 0px;BORDER-BOTTOM: #014E07 1px solid;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px}")
		MyFile.Write(".heading {FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR: #008000;FONT-FAMILY: Arial, Verdana, Tahoma, Arial;}")
		MyFile.Write(".style1 { border: 1px solid #014E07;padding: 0px 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
		MyFile.Write(".style2 { border: 1px solid #014E07;padding: 0px 4px;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
		MyFile.Write(".style3 { border: 1px solid #014E07;padding: 0px 4px;FONT-WEIGHT: bold;FONT-SIZE: 9pt;COLOR: #000000;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 2px;}")
		MyFile.Write(".style4 { border: 1px solid #014E07;padding: 0px 4px;FONT-SIZE: 9pt;COLOR: #0000CC;FONT-FAMILY: Arial, helvetica, sans-serif;HEIGHT: 20px;width: 180px;}")
		MyFile.Write("</style>")
		MyFile.Write("<head>")
		MyFile.Write("<title>" & APPLICATION_NAME & "_Error Log</title>")
		MyFile.Write("</head>")
		MyFile.Write("<body>")
		MyFile.Write("<table cellSpacing='0' cellPadding='0' width='96%' border='0'  align='center' style='height: 40px'>")
		MyFile.Write("<tr>")
		Head_Title =" Error Log for Test Case :-  " & TESTCASE_NAME
		MyFile.Write("<td   align=center><span class='heading'><A NAME=Heading>" & Head_Title & "</A> </span>")
		MyFile.Write("<br />")
		MyFile.Write("</td></tr>")
		MyFile.Write("</table>")
		MyFile.Write("</table>")
		MyFile.Write("<td class =style1 >")
		If HTML_RESULT_SUMMARY <> "" Then 
			MyFile.Write("<a href='" & "file:///" & HTML_RESULT_SUMMARY &"' style='color: #FF9966' target='blank'>"& "Summary" &"</a>")
			MyFile.Write("<br />")
		End If
		If htmlStepResult <> "" Then
			MyFile.Write("<a href='" & "file:///" & htmlStepResult &"' style='color: #FF9966' target='blank'>"& "Step Results" &"</a>")
			MyFile.Write("<br />")
		End If
		MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' margin-left:20px;'>")
		MyFile.Write("<TR>")
		MyFile.Write("<TD class=subheading vAlign=center align=middle >Framework_Error</TD>")
		MyFile.Write("<TD class=subheading vAlign=center align=middle >RunTime_Error</TD>")
		MyFile.Write("<TD class=subheading vAlign=center align=middle >Execution_Date</TD>")
		MyFile.Write("<TR>")
		MyFile.Write("<TD class=bg_darkblue height=1></TD>")
		MyFile.Write("<TD class=bg_darkblue  height=1></TD></TR>")
		MyFile.Write("<TR>")
		MyFile.Write("<TD class=bg_gray_eee  height=1></TD>")
		MyFile.Write("<TD class=bg_gray_eee height=1></TD></TR>")
		MyFile.Write("<TR>")
		MyFile.Write("<td  class ='tdborder_1_Fail' vAlign=center align=middle><b>" & FWErr_Count & "</b></td>")
		MyFile.Write("<td  class ='tdborder_1_Fail' vAlign=center align=middle><b>" & RunErr_Count & "</b></td>")
		MyFile.Write("<TD class='tdborder_1'  vAlign=center align=middle ><b>"& Now &"</b></TD>")
		MyFile.Write("</TD></TR>")
  		MyFile.Write("</table>")
		MyFile.Write("<br />")
		MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' margin-left:20px;'>")
		MyFile.Write("<tr>")
		MyFile.Write("<td class='subheading' >S No</td>")
		MyFile.Write("<td class='subheading'>Function_Name</td>")
		MyFile.Write("<td class='subheading'>Error_Type</td>")
		MyFile.Write("<td class='subheading'>Error_Description</td>")
		MyFile.Write("</tr>")
		If FWErr_Count > 0 Then
			For i=1 to 	FWErr_Count
				FWErr = split(FWErr_Array(i),"=>",-1,1)
				S_No = i
				MyFile.Write("<tr> ")
				MyFile.Write("<td class =style2>" & S_No & "</td>")
				MyFile.Write("<td class =style2>" & FWErr(0) & "()" & "</td>")
				MyFile.Write("<td class =style2>" & "Framework Error" & "</td>")
				MyFile.Write("<td class = 'tdborder_1_Fail'>" & FWErr(1) & "</td>")
				MyFile.Write("</tr>")
			Next
			ERRORS = S_No
		End if
		If RunErr_Count > 0 Then
			For j=1 to RunErr_Count
				RunErr = split(RunErr_Array(j),"=>",-1,1)
				S_No = FWErr_Count + j
				MyFile.Write("<tr> ")
				MyFile.Write("<td class =style2>" & S_No & "</td>")
				MyFile.Write("<td class =style2>" & RunErr(0) & "()" & "</td>")
				MyFile.Write("<td class =style2>" & "RunTime Error" & "</td>")
				MyFile.Write("<td class = 'tdborder_1_Fail'>" & RunErr(1) & "</td>")
				MyFile.Write("</tr>")
			Next 
			ERRORS = S_No
		End If
		MyFile.Write("</table>") : MyFile.Write("<br />")
		MyFile.Write("</body>") : MyFile.Write("</html>")
		MyFile.Close
		Reporter.ReportEvent micInfo, "HTML_Result", "Created HTML Error Log in   ===>   " & errLogPath
		Set fso = Nothing : Set MyFile = Nothing
	End If
	' handle error
	methodName = "HTMLErrorLog" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function HTMLErrorLog() ---------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: DBConnect()
'# Purpose			: To establish Database connection
'# Parameters 		: DBConnection_String		-> Gets the Database Connection String
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function DBConnect( byRef curSession ,DBConnection_String)
	On Error Resume Next
	Dim methodName, DBConnection, rc
	methodName="DB Connection" : DBConnect = 0

	' Opening connection
	Set DBConnection = CreateObject("ADODB.Connection")
	DBConnection.Open DBConnection_String
	Set curSession = DBConnection
	Reporter.ReportEvent micInfo, methodName, "Successfully connected to the Database"
	
	' Handling Error
	rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function DBConnect() ------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: createDescription()
'# Purpose			: To Create Object Description with Dynamic inputs
'# Parameters 		: objectDesc		-> Input has the value of Object properties ex: "name:=object1" & "+" & "index:=1"
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Object_Desc
Public Function createDescription(objectDesc)
	On Error Resume Next
	Dim methodName, ObjSplit_Array, ObjSplit_Count, Results_Msg
	methodName = "createDescription" : createDescription=0 : Object_Desc = ""

	ObjSplit_Array = split(objectDesc," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
	Select Case (ObjSplit_Count)
		Case "0"
			Object_Desc = objectDesc
		Case "1"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
		Case "2"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
		Case "3"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
		Case "4"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
		Case Else
			Results_Msg = "Object :- " & UCase(objectDesc) & " has MORE than 5 properties defined, Currently function :- " & UCASE(methodName) & "() won't handle object has more than 5 properties. Please check Object Description or Enhance the function"
			Reporter.reportevent micFail,StepName, Results_Msg
			FW_Error = FW_Error & " / " & methodName & " --> " & Results_Msg
			createDescription = -1
	End Select
	' Handling Error
	methodName = "createDescription" : rc = ErrorHandler(methodName)
End Function	
' --------------------------- End of Function createDescription() ----------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: captureScreen()
'# Purpose			: To capture Screen-shot
'# Parameters		: None	
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public ERROR_SCREEN_FILE
Public Function captureScreen()
	On Error Resume Next
	Dim methodName, fso, objFolder
	methodName = "captureScreen" : captureScreen = 0

	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FolderExists(ERROR_SCREEN_PATH)Then Set objFolder=fso.createFolder(ERROR_SCREEN_PATH) End If
	ERROR_SCREEN_FILE = "" : ERROR_SCREEN_FILE = ERROR_SCREEN_PATH & TESTCASE_NAME & "_" & StepName & ".png"
	Desktop.CaptureBitmap(ERROR_SCREEN_FILE), True
	' handle error
	methodName = "captureScreen" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function captureScreen() --------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: clearCache()
'# Purpose			: To delete cache memory
'# Parameters		: None	
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function clearCache()
	Dim methodName : methodName = "clearCache" : clearCache = 0
	Dim WshShell
'	Below code will delete all the files : History, Cookies, Temporary Internet Files, Form Data, Stored Passwords

	Set WshShell = CreateObject("WScript.Shell")
	WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255"
	Set WshShell = Nothing
'On Error Resume Next  
''On Error Resume Next  
'          Dim methodName, fso, WshNetwork, WshShell, strVer, OSver, strWinFolder, strTempFolder, GetOS, strProfile, WindowsFolder, TemporaryFolder, NewFolder
'          Dim objUserEnv, userProfile, strTempPath, objFile, objFolder, suffix, fname
'          Dim Flag1, Flag2, Flag3
'          methodName = "clearCache" : clearCache = 0 'methodName, 
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'          Set WshNetwork = CreateObject("Wscript.Network") 
'          Set WshShell = createobject("wscript.shell")
'          Set strVer = WshShell.exec("cmd /c ver")
'          Set objUserEnv=WshShell.Environment("User")
'          OSver = strVer.stdout.readall 
'
'          If InStr(OSver, "XP") Then GetOS = "WXP"
'          If InStr(OSver, "2000") Then GetOS = "W2K"
'          If InStr(OSver, "NT") Then GetOS = "NT4"
'          If InStr(OSver, "98") Then GetOS = "W98" 
'          If InStr(OSver, "Millennium") Then GetOS = "W98"
'
''         Set strWinFolder =  fso.GetSpecialFolder(WindowsFolder)
''         strTempFolder =  fso.GetSpecialFolder(TemporaryFolder)
'
'    If GetOS = "WXP" OR GetOS = "W2K" Then
'                   Flag1 = True
'        strProfile = "C:\Documents and Settings\"
'    Else
'                   Flag1 = False
'    End If
'                
'    If GetOS = "NT4" Then
'                   Flag2 = True
'                   strProfile = "C:\winnt\profiles\"
'    Else
'                   Flag2 = False
'    End If
'                
'    If Flag1 = False and Flag2 = False Then
'                   GetOS = "W7"
'                   Flag3 = True
'                   strProfile = "C:\Users\"
'    Else
'                   Flag3 = False
'    End If
'
'          userProfile = WshShell.ExpandEnvironmentStrings("%userprofile%")
'          strTempPath = userProfile & "\Cookies\"
'
'          Set objFolder = fso.GetFolder(strTempPath)'.GetFolder(strTempPath)'strProfile & "Cookies\"
''         Delete all the cookies
'          For Each objFile In objFolder.Files
'                   fname = objFile.Name
'                   suffix = LCase( Right( fname, 4 ) )
'                   If Not suffix = ".dat" Then
'                             objFile.delete True
'                   End If
'          Next
'
'          Set strWinFolder =  fso.GetSpecialFolder(WindowsFolder)
'                   If Flag1 = True Or Flag2 = True Then
''                            Delete the Temp files from C:\Windows\Temp
'                             If fso.FolderExists(strWinFolder & "\Temp\") Then 
''                                      fso.DeleteFile (strWinFolder & "\Temp\*.*"), True : fso.DeleteFolder (strWinFolder & "\Temp\*.*"), True 
'                             End If 
'            If  Not fso.FolderExists(strWinFolder & "\Temp\") Then
'                                      NewFolder = fso.CreateFolder (strWinFolder & "\Temp\")
'                             End If
'                 
''                            Delete the recently viewed document links in the C:\Documents and Settings\"USERNAME"\Recent folder. 
'                             If fso.FolderExists(strProfile & WshNetwork.username & "\Recent\") Then
'                                      fso.DeleteFile  (strProfile & WshNetwork.username & "\Recent\*.*"), True
'                             End If
'
''                            Delete the Temp files in the C:\Documents and Settings\"USERNAME"\Local Settings\Temp folder. 
'                             If fso.FolderExists(strProfile & WshNetwork.username & "\Local Settings\Temp\") Then 
''                                      fso.DeleteFile (strProfile & WshNetwork.username & "\Local Settings\Temp\*.*"), True 
''                                      fso.DeleteFolder (strProfile & WshNetwork.username & "\Local Settings\Temp\*.*"), True 
'            End If 
'
''           Delete the Temporary Internet files and folders in the C:\Documents and Settings\"USERNAME"\Local Settings\Temporary Internet Files folder. 
''           wshShell.run "cmd /c del " & strprofile & WshNetwork.Username & "\Tempor~1\*.* /q", 1, True 
''           fso.DeleteFile (strProfile & WshNetwork.username & "\Tempor~1\*.*"), True
'            If fso.FolderExists(strProfile & WshNetwork.username & "\Local Settings\Temporary Internet Files\") Then 
'                                      fso.DeleteFile (strProfile & WshNetwork.username & "\Local Settings\Temporary Internet Files\*.*"), True 
''                                      fso.DeleteFile (strProfile & WshNetwork.username & "\Local Settings\Temporary Internet Files\Content.IE5\*.*"), True 
'            End If
'                                
''                            Deletes the saved passwords of the internet explorer from C:\Documents and Settings\"USERNAME"\Application Data\Microsoft\Credentials
'            If fso.FolderExists(strProfile & WshNetwork.username & "\Application Data\Microsoft\Credentials\") Then 
'                                      fso.DeleteFile (strProfile & WshNetwork.username & "\Application Data\Microsoft\Credentials\*.*"), True                                                '
'            End If 
'        End If
'
'        If Flag3 = True Then
''         Delete the Recent files on WIndows7    C:\Users\AA24800\AppData\Roaming\Microsoft\Windows\Recent
'                             If fso.FolderExists(strProfile & WshNetwork.username & "\AppData\Roaming\Microsoft\Windows\Recent") Then
'                                      fso.DeleteFile (strProfile & WshNetwork.username & "\AppData\Roaming\Microsoft\Windows\Recent\*.*"), True 
'                             End If
'
''         Delete Temporary Internet Files on WIndows7
'                             If fso.FolderExists(strProfile & WshNetwork.username & "\AppData\Local\Microsoft\Windows\Temporary Internet Files") Then
'                                      fso.DeleteFile (strProfile & WshNetwork.username & "\AppData\Local\Microsoft\Windows\Temporary Internet Files\*.*"), True
'                             End If
'        End If
'
'          set fso = Nothing : set wshshell = Nothing : set WshNetwork = Nothing : set NewFolder = Nothing : Set objFolder = Nothing : Set objUserEnv = Nothing
'
'	handle error
    methodName = "clearCache" : rc = ErrorHandler(methodName)
End Function

' --------------------------- End of Function clearCache() -----------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function			:	ErrorHandler(method_Name)
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Run_Error
Public Function ErrorHandler(method_Name)
	Dim ERR_MSG
	ErrorHandler = 0
	If Err.Number <> 0 Then
		ERR_MSG = "Error # " & CStr(Err.Number) & " - " & Err.Description & " - " & Err.Source & " in Function -> " & method_Name & "()"
		Reporter.ReportEvent micFail, "ErrorHandler", ERR_MSG
		Run_Error = Run_Error & " / " & method_Name & "=>" & ERR_MSG
		ErrorHandler = 1 : RError_Flag = "Y" : ErrorFlag = "e" : ERR_MSG = ""
		Err.Clear	
	End If	
End Function
' --------------------------- End of Function ErrorHandler() ---------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name    : openExcel()
'# Purpose          : To open an excel.
'# Usage			: <rc = openExcel>
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public excelObj, wrkBookObj
Public Function openExcel(filePath, fileName)
	Dim methodName
	methodName = "openExcel" : openExcel = 0
	Reporter.ReportEvent micInfo, methodName,"Opening DATA file -> " & filePath & fileName	
	Set excelObj = CreateObject("Excel.Application")				' create the excel object
	Set wrkBookObj = excelObj.Workbooks.Open(filePath & fileName, False, True)		' open workbook
	' handle error
	methodName = "openExcel" : openExcel = ErrorHandler(methodName)
End Function
' --------------------------- End of Function openExcel() ------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name    : closeExcel()
'# Purpose          : To close an open excel.
'# Usage			: <rc = closeExcel>
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function closeExcel(filePath, fileName)
	Dim methodName
	methodName = "closeExcel" : closeExcel = 0
	Reporter.ReportEvent micInfo, methodName,"Closing DATA file -> " & filePath & fileName
	excelObj.DisplayAlerts = False	' disabling the alerts
	excelObj.Workbooks.Close		' close the workbook
	excelObj.Application.Quit		' clean the object
	Set excelObj = Nothing : Set wrkBookObj = Nothing
	' handle error
	methodName = "closeExcel" : closeExcel = ErrorHandler(methodName)
End Function
' --------------------------- End of Function closeExcel() -----------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name    : MailAlert()
'# Purpose          : To send mail with Framework HTML Result.
'# Usage			: <rc = MailAlert>
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function MailAlert(In_TO, In_CC, In_BCC, In_Subject, In_Body, In_Attachment)
	Dim methodName, objOutlook, objMail
	methodName = "MailAlert" : MailAlert = 0
	Set objOutlook=CreateObject("Outlook.Application")
	Set objMail=objOutlook.CreateItem(0)
	objMail.TO= In_To : objMail.CC=In_CC : objMail.BCC=In_Bcc
	objMail.Subject=In_Subject : objMail.Body= In_Body
	If (In_Attachment <> "") Then objMail.Attachments.Add(In_Attachment) End If   
	objMail.Send
	Set objMail = Nothing : Set objOutlook = Nothing
	Reporter.ReportEvent micInfo, methodName, "Sent email with Framework Test Results."
	methodName = "MailAlert" : MailAlert = ErrorHandler(methodName)
End Function
' --------------------------- End of Function MailAlert() ------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name    : PopUp_Recovery()
'# Purpose          : To Handle PopUps.
'# Usage			: -
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Function PopUp_Recovery()
	On Error Resume Next
	Dim rc, method_Name
	If Dialog("Class Name:=Dialog").exist(0) Then
		rc = Dialog("Class Name:=Dialog").WinButton("text:=ok|OK|Yes|&Run|&Finish").exist(0)
		If rc = "True" Then
			Dialog("Class Name:=Dialog").WinButton("text:=ok|OK|Yes|&Run|&Finish").Click
		End If
	End If
	If Dialog("regexpwndtitle:=.*Internet Explorer").Exist(0) Then
		rc = Dialog("regexpwndtitle:=.*Internet Explorer").WinButton("regexpwndtitle:=ok|OK|Yes|&Run|&Finish").exist(0)
		If rc = "True" Then
			Dialog("regexpwndtitle:=.*Internet Explorer").WinButton("regexpwndtitle:=ok|OK|Yes|&Run|&Finish").Click
		End If
	End If
	If Dialog("regexpwndtitle:= Security Warning").Exist(0) Then
		rc = Dialog("regexpwndtitle:= Security Warning").WinButton("regexpwndtitle:=ok|OK|Yes|&Run|&Finish").exist(0)
		If rc = "True" Then
			Dialog("regexpwndtitle:= Security Warning").WinButton("regexpwndtitle:=ok|OK|Yes|&Run|&Finish").Click
		End If
	End If
	If Window("text:=.*Internet Explorer|regexpwndtitle:= Security Warning").Exist(1)  Then
		Window("text:=.*Internet Explorer").WinButton("text:=ok|OK|Yes|&Run|&Finish").Click
	End if 
	method_Name = "PopUp_Recovery"
	Call ErrorHandler(method_Name)
End Function 

Dim method_Name : method_Name = "FW_Common_Lib" : Call ErrorHandler(method_Name)

'#******************   End of FW_Common_Lib   ******************************************************************************




