'###########################################################################################################################
'#
'#	FW_Driver_Lib:-		Contains Driver Functions Used by Generic Automation Framework - Version 3.0
'#__________________________________________________________________________________________________________________________
'#	Functions Present:
'#		1.	TESTRUNNER()
'#		2.	EXECUTETESTCASE()
'#		3.	LOADLIBRARIES()
'#		4.	READDATA()
'#
'###########################################################################################################################

'Option Explicit		'	-	Declare all the variables used

'___________________________________________________________________________________________________________________________
'# Function Name	: TESTRUNNER()
'# Purpose			: Drive the Execution of Test Cases 
'# Usage			: <rc = TESTRUNNER()>
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public TESTCASE_NAME, ITERATION, TESTER, FW_Error, ErrorFlag

Public Function TESTRUNNER()
	On Error Resume Next
	Dim methodName, TC_NAME, TD_USER_NAME, rc, Run_Count, Result_Msg, i
	methodName = "TESTRUNNER" : TESTRUNNER = 0	

	Reporter.ReportEvent micInfo, methodName, "Running Automation Framework - Version " & FRAMEWORK_VERSION
	' Get the Test Case Name	
	'#6/27 SS- Commenting the QCUtil code from 32-40
	If QCutil.IsConnected Then
		Reporter.ReportEvent micInfo, methodName, "Reading Test Case Name from Quality Center."
		TD_USER_NAME = Trim(QCUtil.TDConnection.UserName)
		TESTER = "QC-" & TD_USER_NAME
		TC_NAME = Trim(QCUtil.CurrentTest.Name)
		Reporter.Reportevent micInfo, methodName, "Test Case Name Read from QC is :- " & TC_NAME


	 ElseIf  Ucase(Environment.value("TestDataInput_Type")) = "FEATUREFILE" Then
		
	print "Test path is %s" &TEST_DATA_PATH
	print Environment.value("TestDataInput_Type")
	'If  Ucase(Environment.value("TestDataInput_Type")) = "FEATUREFILE" Then	
		
		Call EXT_method_exec()

		TC_TEMP_ARRAY = split(TC_TEMP,"+++",-1,1)


'for each x in TC_TEMP_ARRAY
 'msgbox x
'next



		TESTER = "QTP-User" : TC_NAME = ""
		Reporter.ReportEvent micInfo, methodName,"Reading Test Case Name from Feature File -> " & TEST_DATA_PATH & TEST_DATA_FILE_NAME


	Else
		TESTER = "QTP-User" : TC_NAME = ""
		Reporter.ReportEvent micInfo, methodName,"Reading Test Case Name from DATA file -> " & TEST_DATA_PATH & TEST_DATA_FILE_NAME	
	End If

	'Open the Data Excel sheet
	Print "Open Excel opertation performed:::"
	Print "TEST Data File Name:::" & TEST_DATA_FILE_NAME
	call openExcel(TEST_DATA_PATH, TEST_DATA_FILE_NAME)
	Print "Close excel operation:::"
	'TC_NAME= "FFWF_BCS_03_FULFILLMENT_ENJ_MASTER_PROCESS_REDESIGN"
	' Get the No. of Iterations
	rc = READDATA(TEST_DATA_SHEET_NAME, TC_NAME, "Iterations")					'Read Test Cases with EXCE_IND=Y	
	Print "Read Test Case data from the Excel:::"
	Print rc
	
	If rc < 0 Then
		TESTRUNNER = -1
	Else
		Iteration_Count = NO_OF_RECORDS
	End If

	' Run Test Case for No. of Iterations
	If Iteration_Count >= 1 Then
		For Run_Count = 1 To Iteration_Count
			ITERATION = Run_Count : Run_Error = "" : FW_Error = "" : ErrorFlag = ""
			If TC_Name <> "" Then
				TESTCASE_NAME = Trim(TC_Name)
				Reporter.ReportEvent micInfo, "Iteration " & ITERATION, "Running Test Case ->  "& UCase(TESTCASE_NAME) & " from QC for  " & UCase(APPLICATION_NAME) & " application." 
			
			
			ElseIf TestCase_Array(Run_Count) <> "" and  TC_TEMP_ARRAY(Run_Count) = "" Then
				TESTCASE_NAME = Trim(TestCase_Array(Run_Count))
				
				Reporter.ReportEvent micInfo, "Iteration " & ITERATION, "Running Test Case ->  "& UCase(TESTCASE_NAME) & " from QTP/UFT for  " & UCase(APPLICATION_NAME) & " application." 
			

			ElseIf TC_TEMP_ARRAY(Run_Count) <> "" Then
				TESTCASE_NAME = Trim(TC_TEMP_ARRAY(Run_Count))

'msgbox TESTCASE_NAME
				Reporter.ReportEvent micInfo, "Iteration " & ITERATION, "Running Test Case ->  "& UCase(TESTCASE_NAME) & " from FEATURE FILE for  " & UCase(APPLICATION_NAME) & " application."

			



			Else
				Result_Msg = "Test Case Name NOT available for Execution" 
				Reporter.ReportEvent micFail, methodName, Result_Msg
				FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
				TESTRUNNER = -1 : ErrorFlag = "e"
				Exit Function
			End If	
			If Run_Count > 1 Then
				' Open the Data Excel sheet
				call openExcel(TEST_DATA_PATH, TEST_DATA_FILE_NAME)
			End If
			rc = READDATA(TEST_DATA_SHEET_NAME, TESTCASE_NAME, "TestData")		' Reading Test Data
			If rc < 0 Then
				Result_Msg = "Failed to read the Test Data for the Test Case :-  "  & UCase(TESTCASE_NAME)
				Reporter.Reportevent micFail, methodName, Result_Msg
				FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
				TESTRUNNER = -1 : ErrorFlag = "e"
				Exit Function
			Else
				Reporter.Reportevent micInfo, methodName, "Successfully read Test Data for the Test Case :-  "  & UCase(TESTCASE_NAME)
			End If				
			rc = READDATA(TEST_STEP_SHEET_NAME, TESTCASE_NAME, "TestStep")		' Reading Test Steps
			If rc < 0 Then
				Result_Msg = "Failed to read the Test Step for the Test Case :-  "  & UCase(TESTCASE_NAME)
				Reporter.Reportevent micFail, methodName, Result_Msg
				FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
				TESTRUNNER = -1 : ErrorFlag = "e"
				Exit Function
			Else
				Reporter.ReportEvent micInfo, methodName, "Successfully read Test Steps for Test Case :-    " & UCase(TESTCASE_NAME)
			End If
			If Run_Count = 1 Then
				rc = READDATA(OBJ_DESC_SHEET_NAME, "", "ObjectDesc")			' Reading Test Steps
				' Close the Data Excel sheet
				Call closeExcel(TEST_DATA_PATH, TEST_DATA_FILE_NAME)
			Else
				' Close the Data Excel sheet
				Call closeExcel(TEST_DATA_PATH, TEST_DATA_FILE_NAME)
			End If
			rc = LOADLIBRARIES													' Load Application libraries
			If rc = 0 Then
				Reporter.ReportEvent micInfo, methodName, "Imported Libraries for Application :- " & APPLICATION_NAME
			Else
				Result_Msg = "Failed to Import Libraries for Application :- " & APPLICATION_NAME
				Reporter.ReportEvent micFail, methodName, Result_Msg
				FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
				TESTRUNNER = -1 : ErrorFlag = "e"
				Exit Function
			End If
			If ITERATIONS = "" Then ITERATIONS = 1 End If
			i = 0
			For i = 1 to ITERATIONS
				ITERATION = i
				rc = EXECUTETESTCASE			' Execute Test Steps
				If rc = 0 Then
					Reporter.ReportEvent micInfo, methodName, "Successfully Executed the Test Steps"
				Else
					Result_Msg = "Failed to Execute Test Steps"
					Reporter.ReportEvent micFail, methodName, Result_Msg
					FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
					TESTRUNNER = -1 : ErrorFlag = "e"
				End If
				' HTML Results
				If UCase(HTML_RESULT_FLAG) <> "OFF" Then
					Call HTMLErrorLog			'Call HTML Error Log
					Call HTMLStepResults		'Call HTML Step Results
					Call HTMLResultSummary		'Call HTML Result Summary
				End If
			Next
		Next
		If UCase(MAIL_ALERT_FLAG) = "ON" Then
			Call MailAlert(TO_MAIL_LIST, CC_MAIL_LIST, BCC_MAIL_LIST, MAIL_SUBJECT, MAIL_BODY, MAIL_ATTACHMENT)
		End If
	Else
		Result_Msg = "NO Test Cases are with EXEC_IND = Y in Data file for Application -> " & UCase(APPLICATION_NAME)
		Reporter.Reportevent micFail, methodName, Result_Msg
		FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
		TESTRUNNER = -1
	End If
	
	' handle error
	methodName = "TESTRUNNER" : TESTRUNNER = ErrorHandler(methodName)
End Function
' --------------------------- End of Function TESTRUNNER() -----------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name    : EXECUTETESTCASE()
'# Purpose          : General Application Run function
'# Usage	    	: <rc = EXECUTETESTCASE()>
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public StepName, EXEC_DATE, START_DATE, START_TIME, END_TIME, DURATION, StepDesc_Dict, ExpRes_Dict, ActRes_Dict, StepTime_Dict, StepFlag_Dict, Step_Description
Public Exp_Result, Actual_Res, StepErSnap_Dict, Exec_Flag, STATUS, TOTAL_STEPS, STEPS_EXECUTED, STEPS_PASSED, STEPS_FAILED, STEPS_ERRORED, RError_Flag
Public Function EXECUTETESTCASE
	On Error Resume Next
	Dim methodName, orderFlowKeysArr, orderFlowKeysCount, orderFlowKey, rc, Temp_Time, ResultFlag, i, Result_Msg, StepValue, Temp1, Temp2, Temp_Arr
	methodName = "EXECUTETESTCASE" : EXECUTETESTCASE = 0
	STEPS_FAILED = 0 : STEPS_PASSED = 0 : STEPS_EXECUTED = 0 : STEPS_ERRORED = 0 : ResultFlag = 0 : Exec_Flag = "Y" : RError_Flag = "" : START_TIME = Time

	orderFlowKeysArr = TEST_STEP_DICT.Keys
	orderFlowKeysCount = UBound(orderFlowKeysArr)
	TOTAL_STEPS = orderFlowKeysCount + 1
	Set StepDesc_Dict = CreateObject("Scripting.Dictionary") : Set ExpRes_Dict = CreateObject("Scripting.Dictionary")
	Set ActRes_Dict = CreateObject("Scripting.Dictionary") : Set StepTime_Dict = CreateObject("Scripting.Dictionary")
	Set StepFlag_Dict = CreateObject("Scripting.Dictionary") : Set StepErSnap_Dict = CreateObject("Scripting.Dictionary")
	MercuryTimers("Total_ExecTime").Start
	For orderFlowKey = 0 To orderFlowKeysCount
		StepName = orderFlowKeysArr(orderFlowKey) : StepValue = TEST_STEP_DICT.Item(StepName)				
		If StepValue <> "" Then
			Step_Description = "" : Exp_Result = "" : Actual_Res = "" : ERROR_SCREEN_FILE = ""
			If Exec_Flag = "Y" Then
				Reporter.ReportEvent micInfo, StepName, "Running Function   ==>   " & StepValue
'			End If
			MercuryTimers(StepName).Start : rc = eval(StepValue) : MercuryTimers(StepName).Stop
			StepDesc_Dict.Add StepName, Step_Description : ExpRes_Dict.Add StepName, Exp_Result
'			If Exec_Flag = "Y" Then
				Temp_Time = MercuryTimers(StepName).Elapsedtime/1000
				StepTime_Dict.Add StepName, Temp_Time : ActRes_Dict.Add StepName, Actual_Res
				StepErSnap_Dict.Add StepName, ERROR_SCREEN_FILE
				STEPS_EXECUTED = STEPS_EXECUTED + 1
			End If
			' rc = 0 --> Step executed successfully and Passed
			'If IsEmpty(rs) Then
			'	rc = -1
			'	StepValue = "NOT EXECUTED THIS STEP AND HENCE SCRIPT EXECUTION STOPPED"		
			'End if
			
			If rc = 0 AND Exec_Flag = "Y" Then
				If RError_Flag = "Y" Then
					StepFlag_Dict.Add StepName, "ERROR" : ErrorFlag = "e" : RError_Flag = ""
				Else
					StepFlag_Dict.Add StepName, "PASS"
				End If
				Result_Msg = "Executed Function ==>  " & StepValue
				Reporter.ReportEvent micpass, StepName, Result_Msg
				STEPS_PASSED = STEPS_PASSED + 1
				If EXIT_FLAG = "Y" Then
					Exec_Flag = "N"
				End If
				If EXIT_COUNT <> "" Then
					Exec_Flag = "N"
					If orderFlowKeysCount + 1 >= orderFlowKey + EXIT_COUNT Then
					For i = orderFlowKey + 1 to orderFlowKey + EXIT_COUNT
						StepName = orderFlowKeysArr(i)
						StepValue = TEST_STEP_DICT.Item(StepName)
						rc = eval(StepValue)
						StepDesc_Dict.Add StepName, Step_Description
						ExpRes_Dict.Add StepName, Exp_Result
						StepFlag_Dict.Add StepName, "Skipped"
						StepTime_Dict.Add StepName, 0
					Next
					Else
						Result_Msg = "Invalid SKIP COUNT - " & EXIT_COUNT &". Not having required Test Steps to Skip."
						Reporter.Reportevent micFail, methodName, Result_Msg
						FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
					End If
					Exec_Flag = "Y"
					orderFlowKey = orderFlowKey + EXIT_COUNT
					EXIT_COUNT = ""
				End If
			' rc < 0 --> Step executed with error or Failed, and won't continue execution
			ElseIf rc < 0 AND Exec_Flag = "Y" Then
				Exec_Flag = "N"
				If RError_Flag = "Y" Then
					StepFlag_Dict.Add StepName, "ERROR" : ErrorFlag = "e" : RError_Flag = ""
				Else
					StepFlag_Dict.Add StepName, "FAIL"
				End If
				Result_Msg = "Failed on Executing Function   ==>   " & StepValue
				Reporter.ReportEvent micfail, StepName, Result_Msg
				STEPS_FAILED = STEPS_FAILED + 1
				ResultFlag = -1
				MercuryTimers("Total_ExecTime").Stop
				Temp1 = MercuryTimers("Total_ExecTime").Elapsedtime/(1000*60)
				Temp_Arr=split(Temp1,".",-1,1) : Temp2 = "."& Temp_Arr(1)
				DURATION =  Temp_Arr(0) & " Min, " & Round(Temp2 * 60) & " Sec"
			' rc = 1 --> Step executed successfully and Failed, and will continue execution
			ElseIf rc = 1 AND Exec_Flag = "Y" Then
				If RError_Flag = "Y" Then
					StepFlag_Dict.Add StepName, "ERROR" : ErrorFlag = "e" : RError_Flag = ""
				Else
					StepFlag_Dict.Add StepName, "FAIL"
				End If
				Result_Msg = "Failed on Executing Function   ==>   " & StepValue
				Reporter.ReportEvent micfail, StepName, Result_Msg
				STEPS_FAILED = STEPS_FAILED + 1
				ResultFlag = -1
			End If
		End If
	Next
	MercuryTimers("Total_ExecTime").Stop : Temp1 = MercuryTimers("Total_ExecTime").Elapsedtime/(1000*60) : Temp_Arr = split(Temp1,".",-1,1) : Temp2 = "."& Temp_Arr(1)
	DURATION =  Temp_Arr(0) & " Min, " & Round(Temp2 * 60) & " Sec"
	EXEC_DATE = Date : END_TIME = Time
	If (ResultFlag = 0 AND ErrorFlag = "") Then STATUS = "PASS" End If
	If (ResultFlag = 0 AND ErrorFlag = "e") Then STATUS = "ERROR" End If
	If ResultFlag = -1 Then STATUS = "FAIL" End If
	If (ResultFlag = "" AND ErrorFlag = "e") Then STATUS = "ERROR" End If
	' handle error
	methodName = "EXECUTETESTCASE" : EXECUTETESTCASE = ErrorHandler(methodName)
End Function
' --------------------------- End of Function EXECUTETESTCASE() ------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: LOADLIBRARIES
'# Purpose			: Application library files are loaded into QTP
'# Usage			: <rc = LOADLIBRARIES()>
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
public qtVer
Public Function LOADLIBRARIES
On Error Resume Next
	Dim methodName
	methodName = "LOADLIBRARIES" : LOADLIBRARIES = 0

	Reporter.ReportEvent micInfo, methodName, "Importing Library for Application  --->   " & APPLICATION_NAME
	
	qtVer = Environment.value("ProductVer")
	
	If (APP_FUNCTION_LIB <> "") And (APP_CONSTANT_LIB <> "") Then
		If qtVer >= 11 Then
			LoadFunctionLibrary APP_FUNCTION_LIB
			LoadFunctionLibrary APP_CONSTANT_LIB
		Else
			'ExecuteFile APP_FUNCTION_LIB : ExecuteFile APP_CONSTANT_LIB
			LoadFunctionLibrary APP_FUNCTION_LIB
			LoadFunctionLibrary APP_CONSTANT_LIB
		End If
	End If
	Reporter.ReportEvent micInfo, methodName, "Successfully Imported the Libraries for Application - " & APPLICATION_NAME & " :- " & Chr(13) & APP_FUNCTION_LIB & Chr(13) & APP_CONSTANT_LIB
	
	' handle error
	methodName = "LOADLIBRARIES" : LOADLIBRARIES = ErrorHandler(methodName)
End Function
' --------------------------- End of Function LOADLIBRARIES() --------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: READDATA
'# Purpose			: To read the DATA sheet to get - Iterations, Test Case Name, Test Step and Test Data
'# Usage			: <rc = READDATA(testDataSheetName, dataSetName, InOpt)>
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public NO_OF_RECORDS, TEST_STEP_DICT, TestCase_Array, Iteration_Count
Public Function READDATA(testDataSheetName, dataSetName, InOpt)
	On Error Resume Next
	Dim methodName, wrkSheetObj, range, rowCount, colCount, row, testCaseName, exec_indicator, dictObj, FoundTestCase, keywordValue, parameterValue
	Dim Temp_Array, FoundTempTest, wrkSheetObj1, range1, rowCount1, colCount1, colt, rowt, tempTestCase, tempName, tempkeywordValue, stepName, stepValue
	Dim Result_Msg, currow, col, dataItemName, dataItemValue, VarName, objClass, objValue, Temp, i
	READDATA = 0 : methodName = "READDATA"

	Print "test Data sheet Name:ReadData():::" &testDataSheetName
	Print "Data set Name:ReadData():::" &dataSetName
	Print "Inopt Name:ReadData():::" &InOpt
		
	Set wrkSheetObj = wrkBookObj.WorkSheets(testDataSheetName)		' get data worksheet	
	Set range = wrkSheetObj.UsedRange								' get data used range
	rowCount = range.Rows.Count : colCount = range.Columns.Count	' get row/column count 
	If rowCount > 0 OR colCount > 0 Then							' check for valid no of rows.
		                                                                                                                                                                    ' Get No. of Iterations and TestCase Name with EXEC_IND = Y
		If InOpt = "Iterations" Then	
			NO_OF_RECORDS=0
			If dataSetName = "" Then	
				For row = 2 to rowCount
					testCaseName = Trim(range.Cells(row,1)) : exec_indicator = Trim(range.Cells(row,2))
					If testCaseName = "" Then Exit For End If
					If exec_indicator = "Y" then
						NO_OF_RECORDS = NO_OF_RECORDS + 1 : Temp = Temp & "+++" & testCaseName
					End If			
				Next
				TestCase_Array = split(Temp,"+++",-1,1)
				Reporter.ReportEvent micInfo, methodName, "No.of Test Cases with EXEC_IND = Y is :- " & NO_OF_RECORDS
			Else 	  	
				For row = 2 to rowcount
					testCaseName = Trim(range.Cells(row,1)) : exec_indicator = Trim(range.Cells(row,2))
					If testCaseName = "" Then Exit For End If
					If exec_indicator = "Y" AND (testCaseName = dataSetName) Then
						NO_OF_RECORDS = NO_OF_RECORDS + 1
					End If
				Next
				Reporter.Reportevent micInfo,methodName, "No.of Iterations for Test Case :- " & UCase(dataSetName) & "is  --> " & NO_OF_RECORDS
			End If	
		End If
		' Read Test Steps for a Test Case
		If InOpt = "TestStep" Then
			Set dictObj = CreateObject("Scripting.Dictionary")		' script dictionary object to store Test Steps
			FoundTestCase = -1 : i = 1
			For col = 2 To colCount
				testCaseName = Trim(range.Cells(1, col)) : dataSetName = Trim(dataSetName)
				If UCase(testCaseName) = UCase(dataSetName) Then
					FoundTestCase = 0
					For row = 3 To rowCount
						keywordValue = Trim(range.Cells(row, col)) : parameterValue	=	Trim(range.Cells(row, col+1))
						If keywordValue = "" Then Exit For End If
						If keywordValue <> "" Then
							Temp_Array = split(keywordValue,"_",-1,1)
							If Temp_Array(0) = "TEMP" Then			' Read Template Test Case Steps
							FoundTempTest = -1
								Set wrkSheetObj1 = wrkBookObj.WorkSheets(TEMP_TEST_SHEET_NAME)		' get data worksheet	
								Set range1 = wrkSheetObj1.UsedRange									' get data used range	
								rowCount1 = range1.Rows.Count : colCount1 = range1.Columns.Count	' get row/column count 
								If rowCount1 > 0 OR colCount1 > 0 Then								' check for valid no of rows.
									For colt = 2 To colCount1
										tempTestCase = Trim(range1.Cells(1, colt)) : tempName = Trim(keywordValue)
										If UCase(tempTestCase) = UCase(tempName) Then
											FoundTempTest = 0
											For rowt = 3 To rowCount1
												tempkeywordValue = Trim(range1.Cells(rowt, colt)) : parameterValue	=	Trim(range1.Cells(rowt, colt+1))
												If tempkeywordValue = "" Then Exit For End If
												If tempkeywordValue <> "" Then
													stepName = "Step " & i : stepValue = tempkeywordValue & "(" & parameterValue & ")"
													dictObj.Add stepName, stepValue
													i=i+1
												End If
											Next
										End If	
									Next
									If FoundTempTest = 0 Then
										Reporter.ReportEvent micInfo, methodName, "Successfully read the Template Test Case :-  "  & keywordValue
									Else
										Result_Msg = "NOT found Template Test Case :-  "  & keywordValue
										Reporter.Reportevent micFail, methodName, Result_Msg
										FW_Error = FW_Error & " / " & methodName & "=>" & Result_Msg
										READDATA = -1 : ErrorFlag = "e"
									End If
								End If
								Set wrkSheetObj1 = Nothing : Set range1 = Nothing
							Else
								stepName = "Step " & i : stepValue = keywordValue & "(" & parameterValue & ")"
								dictObj.Add stepName, stepValue
								i=i+1
							End If
						End If
					Next
				End If
			Next
			Set TEST_STEP_DICT = dictObj : Set dictObj = Nothing
			If FoundTestCase = -1 Then
				Result_Msg = "Test Case :-   " & dataSetName & "    NOT found in Test Step Sheet."
				Reporter.ReportEvent micFail, methodName, Result_Msg
				FW_Error = FW_Error & " / " & methodName & " ===> " & Result_Msg
				READDATA = -1 : ErrorFlag = "e"
			End If
		End If
  ' Read the test data for a test case
		If InOpt = "TestData" Then
			FoundTestCase = -1
			If DATA_ROW = "" Then currow = 2 Else currow = DATA_ROW + 1 End If
			For row = currow To rowCount
				testCaseName = Ucase(Trim(range.Cells(row, 1))) : exec_indicator = Ucase(Trim(range.Cells(row, 2)))
				dataSetName = Ucase(Trim(dataSetName))
				'msgbox testCaseName
				'msgbox dataSetName
				If testCaseName = "" Then Exit For End If
				If (testCaseName = dataSetName) AND (exec_indicator = "Y") Then
					FoundTestCase = 0
					DATA_ROW = row			
					For col = 3 To colCount
						dataItemName = Trim(range.Cells(1, col)) : dataItemValue = Trim(range.Cells(row, col))
						If dataItemName = "" Then Exit For End If
						If dataItemName <> "" Then Execute(dataItemName & " = dataItemValue") End If
					Next
					Exit For
				End If
			Next
			If FoundTestCase = -1 Then
				Result_Msg = "Test Case :-   " & dataSetName & "    NOT found in Test Data Sheet."
				Reporter.ReportEvent micFail, methodName, Result_Msg
				FW_Error = FW_Error & " / " & methodName & " ===> " & Result_Msg
				READDATA = -1 : ErrorFlag = "e"
			End If
		End If	
		' Read the object description
		If InOpt = "ObjectDesc" Then	
			For row = 2 To rowCount
				VarName = Trim(range.Cells(row, 1))
				If VarName = "" Then Exit For End If
				If VarName <> "" Then 
					objClass = Trim(range.Cells(row, 2)) : temp = "objDesc = " & Trim(range.Cells(row, 3)) : Execute(temp)
					objValue = objClass & "=>" & objDesc
					execute VarName & " = objValue" 
				End If
			Next
			Reporter.ReportEvent micInfo, methodName, "Successfully read Object Descriptions."
		End If
	Else
		Result_Msg = "Excel has invalid no. of rows/columns. Row Count :- " & rowCount & " Column Count :- " & colCount
		Reporter.ReportEvent micFail, methodName, Result_Msg
		FW_Error = FW_Error & " / " & methodName & " ===> " & Result_Msg
		READDATA = -1 : ErrorFlag = "e"
	End If

	Set wrkSheetObj = Nothing : Set range = Nothing
	' handle error
	methodName = "READDATA" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function READTESTSTEP() ---------------------------------------------------------------

Dim DATA_ROW : DATA_ROW = ""
