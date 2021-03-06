'#################################################################################################################
'###
'###	FUNCTION:               EXT_method_exec(pTestName) 
'###
'###	DESCRIPTION:    This function will execute the object-methods for the test case name passed in to the function.
'###
'###	PARAMETERS:     pTestName:  The name of the test
'###	Author: SUMAN SOME
'###
'#################################################################################################################
'On Error Resume Next

Public TestDataInputType

Public Function EXT_method_exec()
	Dim TestDataInputType, Temp1, Temp2, Temp_Arr
	Dim oShell
	MercuryTimers("Total_ExecTime").Start
	
	''### Load the test properties file
	'Set gTestProperties = loadPersonalProperties(Environment ("parameters_file") )
	'Dim  actionArray() : ReDim actionArray(-1)
	TestDataInputType = UCase(Environment.Value("TestDataInput_Type"))	'Excel / FeatureFile
	'Purpose = Environment.Value("Purpose")
	'Environment("File_dir") = "C:\work\jiffy\main_tests\" & Purpose & "\"	'Newly added
	'Environment("test_case_dir") = Environment("test_case_dir") & Purpose & "\"	'Newly added
	Select Case TestDataInputType
	
		Case "TEXTFILE"
			'### This is the location of the test case Text file
			actionFile = Environment("File_dir") & Environment.Value("FileName") & ".txt"
			'### Download the text file from ALM if ALM is connected
			If QCUtil.IsConnected Then
				'BACKUP TESTCASE FILE
				Call BackUpTestCaseFile(Environment("test_case_dir"))
				'### Download the text file from ALM
				QCGetResource Environment.Value("FileName") & ".txt", Environment("test_case_dir")
			End If

			'READING TEST STEPS FROM TEXT FILE
			'### Load the test case file into an array
			action_Array = FSO_parseFile(action_Array, actionFile)
			For i = 0 To UBound(action_Array)
				redim preserve actionArray(ubound(actionArray) + 1)
				actionArray(ubound(actionArray) ) = action_Array(i)  
			Next
		
		
		Case "FEATUREFILE"
			'READING TEST STEPS FROM FEATURE FILE


 	Dim oFolder,Folders,Item,x 
 	Set oFolder = CreateObject("Scripting.FileSystemObject") 

 
 	x= oFolder.GetAbsolutePathName(FEATURE_FOLDER) 
 	 
 	Set Folders = oFolder.GetFolder (x) 
		 
 	For Each Item In Folders.Files 
 		If InStr(1,Item.Name,".feature")>0 Then 
 			'Run a feature 
 			'ReadFeatureFile Item.Name 
			'TESTCASE_NAME= "FFWF_BCS_03_FULFILLMENT_ENJ_MASTER_PROCESS_REDESIGN"
			sFile = FEATURE_FOLDER & Item.Name
			'sTestName = TESTCASE_NAME
			Print "Paarsing the feature file"
     		Print	 "Test Case Name:::" & TESTCASE_NAME

			Call FSO_parseFeatureFile(sFile, TESTCASE_NAME)
			
			'For Each element In stepsFromFile
				'tempValue = stepsFromFile.Item(element)
				'redim preserve actionArray(ubound(actionArray) + 1)
				'actionArray(ubound(actionArray) ) = element & " : " & tempValue
			'Next
			
		Else
			msg = "Please set test data input type to TestDataInput_Type environment variable"
			Reporter.ReportEvent micFail, "EXT_method_exec", msg
			Exit Function

End if		 
 	Next
		
	End Select

	Reporter.ReportEvent micDone, "EXT_method_exec", "Running test case:  "& TESTCASE_NAME & "."
		
		'========================================================================================
		'Add number of methods to Dictionary object
		'ExeSteps = i	 + 1	
		'========================================================================================
	'Next
	If TestDataInputType = "TEXTFILE" Then
		iAction = action_Array	
	Else
		iAction = actionArray
	End If
	
	MercuryTimers("Total_ExecTime").Stop : Temp1 = MercuryTimers("Total_ExecTime").Elapsedtime/(1000*60) : Temp_Arr = split(Temp1,".",-1,1) : Temp2 = "."& Temp_Arr(1)
	DURATION =  Temp_Arr(0) & " Min, " & Round(Temp2 * 60) & " Sec"
	'========================================================================================
	'Call the HTML Step result function
	Call HTMLStepResults
	'========================================================================================
         	methodName = "EXT_method_exec" : rc = ErrorHandler(methodName)
End Function 

'#################################################################################################################
'###
'###  FUNCTION:				FSO_parseFeatureFile
'###
'###  DESCRIPTION:		This function will parse the feature file and read the Scenario, Given, When, Then and And statements into array
'###
'###  PARAMETERS:  		parm_file_name:	Feature file location including file name
'###					sTestName:			Test name
'###
'#################################################################################################################
Public stepsFromFile, TC_TEMP, TC_TEMP_ARRAY
Public Function FSO_parseFeatureFile(parm_file_name, sTestName)
	Set stepsFromFile	= CreateObject("Scripting.Dictionary")
	Dim line, config_keys, config_items
	Dim match_result
	Dim fso
	Dim msg, theFile
	Set fso = CreateObject("Scripting.FileSystemObject")

	Set theFile = FSO_file_open( parm_file_name,  ForReading, false )
	If  isobject(theFile) Then
		Reporter.ReportEvent micDone, "fillArray", "Reading contents of file " & parm_file_name & " into array "' & parm_Array
	Else
		Reporter.ReportEvent micFail, "Config File Not Found", "File " & parm_file_name & " could not be found."
		Exit Function
	End If
	Set tempFolder = fso.GetSpecialFolder(2)
	If fso.FileExists(tempFolder & "\tempFeatureFile.txt") Then
		fso.DeleteFile tempFolder & "\tempFeatureFile.txt", True
	End If
	Set tempFile = fso.CreateTextFile(tempFolder & "\tempFeatureFile.txt", True)
	
	Do While Not theFile.AtEndOfStream
		line = theFile.ReadLine
		tempFile.WriteLine(line)
	Loop
	tempFile.Close
	
	Set tempFile = FSO_file_open( tempFolder & "\tempFeatureFile.txt",  ForReading, false )
	startFlag = False : endFlag = False : readFlag = False
	
	Do While Not tempFile.AtEndOfStream
		
		line = tempFile.ReadLine
		line = CleanString(line)
		If InStr(line,"""") > 0 Then
			line = Replace(line, """", "'")
		End If

		If Not line = "" Then
			words = split(line, " ")
			keyword = lcase(words(0))
			'msgbox keyword
			If  keyword = Lcase("@Execute") Then
				startFlag = True
				endFlag = False
				
				ElseIf keyword = Lcase("@ExecuteNot")  Then
					startFlag = False
					endFlag = True
			
			End If
			
			If keyword = LCase("Scenario:") Then
				iLength = Len(line)
				'msgbox line
				line = Mid(line,10,iLength)
				tempTestName = CleanString(line)				'LTrim(line)
				tempTestName1 = replace(tempTestName," ","_")
				'Msgbox tempTestName1
				
				
				
				'If LCase(sTestName) = LCase(tempTestName) Then
					'stepsFromFile.Add keyword, tempTestName
					'startFlag = True
					'endFlag = False
				'Else
					'startFlag = False
					'endFlag = True
				'End If
			End If
			'after this read all the lines and create flag to indicate we read the steps for given scenario and then as soon as reach to next scenario set the ENDFLAG to false
			If (startFlag = True) And (endFlag = False) Then
				readFlag = True
				Select Case keyword
					Case LCase("Given")
						iLength = Len(line)
						line = Mid(line,6,iLength)
						line = CleanString(line)
						line1 = replace(line," ","_")
						TC_TEMP = TC_TEMP & "+++" & line1
						'TC_TEMP_ARRAY = 
						'msgbox line1
						'stepsFromFile.Add keyword, line
						
					
					Case LCase("When")
						iLength = Len(line)
						line = Mid(line,5,iLength)
						line = CleanString(line)
						line1 = replace(line," ","_")
						TC_TEMP = TC_TEMP & "+++" & line1
						'Msgbox line1
						'stepsFromFile.Add keyword, line
					
					Case LCase("Then")
						iLength = Len(line)
						line = Mid(line,5,iLength)
						line = CleanString(line)
						line1 = replace(line," ","_")
						TC_TEMP = TC_TEMP & "+++" & line1
						'msgbox line1
						'stepsFromFile.Add keyword, line
					
					Case LCase("And")
						iLength = Len(line)
						line = Mid(line,4,iLength)
						line = CleanString(line)
						line1 = replace(line," ","_")
						TC_TEMP = TC_TEMP & "+++" & line1
						'msgbox line1
						'stepsFromFile.Add keyword, line
					
					Case Else
					
				End Select
			End If

		End If
		
		'TC_TEMP_ARRAY = split(TC_TEMP,"+++",-1,1)
		'msgbox TC_TEMP_ARRAY
	Loop

'TC_TEMP_ARRAY = TC_TEMP_ARRAY & split(TC_TEMP,"+++",-1,1)

	thefile.Close
	tempFile.Close
	fso.DeleteFile tempFolder & "\tempFeatureFile.txt", True
	Set tempFile = Nothing : Set tempFolder = Nothing
	Set theFile = nothing : Set fso = nothing

methodName = "FSO_parseFeatureFile" : rc = ErrorHandler(methodName)
End Function


'#################################################################################################################
'# FUNCTION NAME	: HTMLSTEPRESULTS
'# PURPOSE			: GENERATES STEP WISE RESULTS IN HTML FORMAT
'# PARAMETERS  		: 
'# AUTHOR			: MANISH CHRISTIAN
'#################################################################################################################
Public Function HTMLStepResults1()

	'If TEST_DATA_INPUT_TYPE = 'FeatureFile' then

	On Error Resume Next
	Dim methodName, Head_Title, fso, MyFile, objFolder, i, TestStep_Key, Steps_Count
	Dim StepDesc, ExpRes, ActRes, StepFlag
	Dim tempValue, element, STATUS, TD_USER_NAME, TESTER
	Dim TOTAL_STEPS, STEPS_EXECUTED, htmlStepResult, tempFolder, iStep
	
	methodName = "HTMLStepResults"
	APPLICATION_NAME = NTM
'========================================================================================
	'"ITERATION" WILL ALWAYS BE 1 AS OF CURRENT FRAMEWORK
	ITERATION = 1
'========================================================================================
	'GETTING "STATUS" VALUE	
	For Each element In StepFlag_Dict
		tempValue = StepFlag_Dict.Item(element)
		If tempValue = "Pass" Then
			STATUS = "PASS"
		ElseIf tempValue = "Warning" Then
			STATUS = "WARNING"
		ElseIf tempValue = "Fail" Then
			STATUS = "FAIL"
		Else
			'Nothing
		End If
	Next

'========================================================================================
	'GETTING "TESTER" AND "TESTCASE_NAME"
	If QCutil.IsConnected Then
		'Reporter.ReportEvent micInfo, methodName, "Reading Test Case Name from Quality Center."
		TD_USER_NAME = Trim(QCUtil.TDConnection.UserName)
		TESTER = "QC-" & UCase(TD_USER_NAME)
		TESTCASE_NAME = UCase(Trim(QCUtil.CurrentTest.Name))
		Reporter.Reportevent micInfo, methodName, "Test Case Name Read from QC is :- " & TESTCASE_NAME
	Else
		TESTER = UCase(Environment.value("UserName")) : TESTCASE_NAME = UCase(Environment("TestName"))
		'Reporter.ReportEvent micInfo, methodName,"Reading Test Case Name from DATA file -> " & TEST_DATA_PATH & TEST_DATA_FILE_NAME	
	End If

'========================================================================================
	'GETTING "TOTAL_STEPS" AND "STEPS_EXECUTED" COUNT
	TOTAL_STEPS = UBound(iAction) + 1 : STEPS_EXECUTED = ExeSteps	'TOTAL_STEPS = iActionubound(actionArray) + 1
'========================================================================================

	RESULTS_PATH = RESULTS_PATH & "\" & BDD & "\"
	htmlStepResult = RESULTS_PATH & TESTCASE_NAME & "_" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & " " & Hour(Now) & " " & Minute(Now) & " " & Second(Now) & ".htm"
	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FolderExists(RESULTS_PATH) Then
		Set tempFolder = fso.CreateFolder(RESULTS_PATH)
	End If

	Set MyFile = fso.CreateTextFile(htmlStepResult, True)
	MyFile.write("<!DOCTYPE html>")
	MyFile.write("<html>")
	MyFile.write("<style>")

	MyFile.Write(".tdborder_1{BORDER-RIGHT: #767676 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #767676 1px solid; PADDING-LEFT: 4px; font-weight: bold; FONT-SIZE: 10pt; PADDING-BOTTOM: 0px; BORDER-LEFT: #767676 1px solid; COLOR: #000000; PADDING-TOP: 0px; BORDER-BOTTOM: #767676 1px solid; FONT-FAMILY: Tahoma; HEIGHT: 22px}")
	MyFile.Write(" .tdborder_1_Pass{BORDER-RIGHT: #767676 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #767676 1px solid; PADDING-LEFT: 4px; FONT-WEIGHT: bold; FONT-SIZE: 10pt; PADDING-BOTTOM: 0px; BORDER-LEFT: #767676 1px solid; COLOR: #41A317; PADDING-TOP: 0px; BORDER-BOTTOM: #767676 1px solid; FONT-FAMILY: Tahoma; HEIGHT: 22px}")
	MyFile.Write(" .tdborder_1_Fail{BORDER-RIGHT: #767676 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #767676 1px solid; PADDING-LEFT: 4px; FONT-WEIGHT: bold; FONT-SIZE: 10pt; PADDING-BOTTOM: 0px; BORDER-LEFT: #767676 1px solid; COLOR: #ff0000; PADDING-TOP: 0px; BORDER-BOTTOM: #767676 1px solid; FONT-FAMILY: Tahoma; HEIGHT: 22px}")
	MyFile.Write(" .tdborder_1_Done{BORDER-RIGHT: #767676 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #767676 1px solid; PADDING-LEFT: 4px; FONT-WEIGHT: bold; FONT-SIZE: 10pt; PADDING-BOTTOM: 0px; BORDER-LEFT: #767676 1px solid; COLOR: #333333; PADDING-TOP: 0px; BORDER-BOTTOM: #767676 1px solid; FONT-FAMILY: Tahoma; HEIGHT: 22px}")
	MyFile.Write(" .heading {FONT-WEIGHT: bold; FONT-SIZE: 14pt; COLOR: #2D2D2D;FONT-FAMILY: Tahoma;}")
	MyFile.Write(" .subheading { BORDER-RIGHT: #767676 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #767676 1px solid; PADDING-LEFT: 4px; FONT-WEIGHT: bold; FONT-SIZE: 10.5pt; PADDING-BOTTOM: 0px; BORDER-LEFT: #767676 1px solid; COLOR: #FFD630; PADDING-TOP: 0px; BORDER-BOTTOM: #767676 1px solid; FONT-FAMILY: Tahoma; HEIGHT: 25px; BACKGROUND-COLOR: #404040}")
	MyFile.Write(" .style1 { border: 1px solid #767676; padding: 0px 4px; FONT-WEIGHT: bold; FONT-SIZE: 10pt; COLOR: #696969; FONT-FAMILY: Tahoma; HEIGHT: 22px; }")
	MyFile.Write(" .style2 { border: 1px solid #767676; padding: 0px 4px; FONT-SIZE: 10pt; COLOR: #000000; FONT-FAMILY: Tahoma; HEIGHT: 22px; }")
	MyFile.Write(" .style3 { border: 1px solid #767676; padding: 0px 4px; FONT-WEIGHT: bold; FONT-SIZE: 10pt; COLOR: #000000; FONT-FAMILY: Tahoma; HEIGHT: 22px; }")
	MyFile.Write(" .style4 { border: 1px solid #767676; padding: 0px 4px; FONT-SIZE: 10pt; COLOR: #0000B2; FONT-FAMILY: Tahoma; HEIGHT: 22px; }")
	MyFile.Write(" html { position: relative; min-height: 100%; }")
	MyFile.Write(" body { margin: 0 0 100px; /* bottom = footer height */ padding: 5px; }")
	MyFile.Write(" footer { position: absolute; left: 44%; bottom: 0; height: 120px; width: 200px; overflow:hidden; }")

	MyFile.Write("</style>")

	MyFile.Write("<head>")
	MyFile.Write("<title>" & APPLICATION_NAME & "_Test Step Results</title>")
	MyFile.Write("</head>")

	MyFile.Write("<body>")
	
	MyFile.Write("<script type=""text/javascript"">")
	MyFile.Write("function toggle2(id, link) {")
	MyFile.Write("  var e = document.getElementById(id);")
	MyFile.Write("  if (e.style.display == '') {")
	MyFile.Write("    e.style.display = 'none';")
	MyFile.Write("	link.innerHTML = ""Expand""")
	MyFile.Write("  } else {")
	MyFile.Write("    e.style.display = '';")
	MyFile.Write("    link.innerHTML = ""Collapse"";")
	MyFile.Write("  }")
	MyFile.Write("}")
	MyFile.Write("</script>")

	MyFile.Write("<br>")

	MyFile.Write("<table cellSpacing='0' cellPadding='0' width='90%' border='0' align='center' style='height: 40px'>")
	MyFile.Write("<tr>")
	Head_Title = "<font color='#585858'> Step Results for Test Case : </font>" & TESTCASE_NAME
	MyFile.Write("<td align=center><span class=heading>" & Head_Title & "</span>")
	'MyFile.Write("<br>")
	MyFile.Write("</td></tr>")
	MyFile.Write("</table>")
	MyFile.Write("<br>")

	MyFile.Write("<br>")

	'OVERALL STATUS TABLE
	MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' style='width:85%;'>")
	MyFile.Write("<tr>")
	MyFile.Write("<td class=subheading vAlign=center align=middle>Iteration</td>")
	MyFile.Write("<td class=subheading vAlign=center align=middle>TestCase Status</td>")
	MyFile.Write("<td class=subheading vAlign=center align=middle>Tester</td>")
	MyFile.Write("<td class=subheading vAlign=center align=middle>Total Modules</td>")	'Total Modules
	MyFile.Write("<td class=subheading vAlign=center align=middle>Modules - Executed</td>")	'Modules - Executed
	MyFile.Write("<td class=subheading vAlign=center align=middle>Exe - Date</td>")
	MyFile.Write("<td class=subheading vAlign=center align=middle>Duration</td>")
	MyFile.Write("</tr>")
	MyFile.Write("<tr>")
	MyFile.Write("<td class=tdborder_1 vAlign=center align=middle>" & ITERATION & "</td>")
	If STATUS = "PASS" Then
		MyFile.Write("<td class =tdborder_1_Pass vAlign=center align=middle>" & STATUS & "</td>")
	ElseIf STATUS = "FAIL" then
		MyFile.Write("<td class =tdborder_1_Fail vAlign=center align=middle>" & STATUS & "</td>")
	ElseIf STATUS = "ERROR" then
		MyFile.Write("<td class =tdborder_1_Fail vAlign=center align=middle>" & STATUS & "</td>")
	End If
	MyFile.Write("<td class=tdborder_1 vAlign=center align=middle>" & TESTER & "</td>")
	MyFile.Write("<td class=tdborder_1 vAlign=center align=middle>" & TOTAL_STEPS & "</td>")
	MyFile.Write("<td class=tdborder_1 vAlign=center align=middle>" & STEPS_EXECUTED & "</td>")
	MyFile.Write("<td class=tdborder_1 vAlign=center align=middle>" & Now & "</td>")
	MyFile.Write("<td class=tdborder_1 vAlign=center align=middle>" & DURATION & "</td>")
	MyFile.Write("</tr>")
	MyFile.Write("</table>")

	MyFile.Write("<br>")

'========================================================================================
	'MODULE INFORMATION HEADER TABLE
'========================================================================================	
	MyFile.Write("<table cellSpacing='0' cellPadding='0' width='90%' border='0'  align='center' style='height: 40px'>")
	MyFile.Write("<tr>")
	Head_Title = " Module Information "
	MyFile.Write("<td align=center><span class=heading>" & Head_Title & "</span>")
	MyFile.Write("</td></tr>")
	MyFile.Write("</table>")
	MyFile.Write("<br>")

'========================================================================================
	'MODULE INFORMATION TABLE
'========================================================================================	
	MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' width='80%' margin-left='20px'>")
	MyFile.Write("<tr>")
	MyFile.Write("<td class=subheading width='3%'>Module #</td>")
	MyFile.Write("<td class=subheading width='77%'>Module_Description from <font size='4'>" & Environment.Value("TestDataInput_Type") & "</font></td>")
	MyFile.Write("</tr>")

	TestStep_Key = StepDesc_Dict.Keys : Steps_Count = UBound(TestStep_Key)
	
	For i = 0 to UBound(iAction)
		iStep = i + 1
		MyFile.Write("<tr>")
		MyFile.Write("<td class =style2>" & iStep & "</td>")
		MyFile.Write("<td class =style2>" & iAction(i) & "</td>")
		MyFile.Write("</tr>")
	Next
	MyFile.Write("</table>")
	MyFile.Write("<br>")

'========================================================================================
	'LOG INFORMATION HEADER TABLE
'========================================================================================
	MyFile.Write("<table cellSpacing='0' cellPadding='0' width='90%' border='0'  align='center' style='height: 40px'>")
	MyFile.Write("<tr>")
	Head_Title = " Logged Information "
	MyFile.Write("<td align=center><span class=heading>" & Head_Title & "</span>")
	MyFile.Write("</td></tr>")
	MyFile.Write("</table>")

'========================================================================================
	'EXPAND/COLLAPSE FUNCTIONALITY
'========================================================================================
	MyFile.Write("<table cellSpacing='0' cellPadding='0' width='90%' border='0'  align='center' style='height: 40px'>")
	MyFile.Write("<tr>")
	MyFile.Write("<td align=center>")
	MyFile.Write("<span style=""font-style: italic; size:5pt; color:#585858; font-weight:bold; font-family:Verdana; text-decoration: underline; cursor: pointer;"" onclick=""toggle2('content', this)"">Expand</span>")
	MyFile.Write("</td>")
	MyFile.Write("</tr>")
	MyFile.Write("</table>")

'========================================================================================
	'LOG INFORMATION TABLE
'========================================================================================
	MyFile.Write("<div style=""display: none; font-family: Verdana"" id=""content"">")
	MyFile.Write ("<table cellSpacing='0' cellPadding='0' border='0' align='center' width='80%' margin-left='20px'>")
	MyFile.Write("<tr>")
	MyFile.Write("<td class=subheading width='4%'>Step #</td>")
	MyFile.Write("<td class=subheading width='25%'>TestStep_Description</td>")
	MyFile.Write("<td class=subheading width='25%'>Expected Result</td>")
	MyFile.Write("<td class=subheading width='25%'>Actual Result</td>")
	MyFile.Write("<td class=subheading width='5%'>Step_Status</td>")
	MyFile.Write("</tr>")

	TestStep_Key = StepDesc_Dict.Keys : Steps_Count = UBound(TestStep_Key)
	
	For i = 0 to Steps_Count
		iStep = i + 1
		MyFile.Write("<tr>")
		StepDesc = StepDesc_Dict.Item(iStep)
		ExpRes = ExpRes_Dict.Item(iStep) : ActRes = ActRes_Dict.Item(iStep)
		StepFlag = StepFlag_Dict.Item(iStep)
		MyFile.Write("<td class =style2>" & iStep & "</td>")
		MyFile.Write("<td class =style2>" & StepDesc & "</td>")
		MyFile.Write("<td class =style2>" & ExpRes & "</td>")
		If ActRes = "" Then ActRes = "-" End If

		If StepFlag = "Pass" Then
			MyFile.Write("<td class =style4>" & ActRes & "</td>")
			MyFile.Write("<td class =tdborder_1_Pass>" & StepFlag & "</td>")
		ElseIf StepFlag = "Done" Then
			MyFile.Write("<td class =style4>" & ActRes & "</td>")
			MyFile.Write("<td class =tdborder_1_Done>" & StepFlag & "</td>")
		ElseIf StepFlag = "Warning" Then
			MyFile.Write("<td class =style4>" & ActRes & "</td>")
			MyFile.Write("<td class =tdborder_1_Fail>" & StepFlag & "</td>")
		ElseIf StepFlag = "Fail" Then
			MyFile.Write("<td class =tdborder_1_Fail>" & ActRes & "</td>")
			MyFile.Write("<td class =tdborder_1_Fail>" & StepFlag & "</td>")
		End If

		MyFile.Write("</tr>")
	Next  
	MyFile.Write("</table>")
	MyFile.Write("</div>")

	MyFile.Write("<br>")
	MyFile.Write("<br>")

'========================================================================================	
	'TABLE WITH JIFFY + CTL ICON
'========================================================================================	
	MyFile.Write("<footer>")
	MyFile.Write("<img src=""..\..\Images\CTL_JIFFY_Logo.png"" alt=""CTL + JIFFY Logo""></img>")
	MyFile.Write("</footer>")

	MyFile.Write("</body>") : MyFile.Write("</html>")
	MyFile.Close
	Reporter.ReportEvent micInfo, "HTML_Result", "Created HTML Step Results in   ===>   " & htmlStepResult
	Set fso = Nothing : Set MyFile = Nothing : Set StepDesc_Dict = Nothing : Set ExpRes_Dict = Nothing : Set ActRes_Dict = Nothing
	Set StepTime_Dict = Nothing : Set StepFlag_Dict = Nothing : Set StepErSnap_Dict = Nothing

'End If

End Function
' --------------------------- END OF FUNCTION HTMLSTEPRESULTS() ------------------------------------------------------------

'#################################################################################################################
'###
'###	FUNCTION:			CleanString(line)
'###
'###  	DESCRIPTION:		Removes the space and VbTab characters from string
'###
'###	PARAMETERS:		line : string that needs to be cleaned
'###
'###	AUTHOR:			Manish Christian
'###
'#################################################################################################################
Public Function CleanString(line)
	line = Trim(line)
	line = Replace(line, vbTab, " ")
	Do While InStr(1, line, "  ")
		line = Replace(line, "  ", " ")
	Loop
	CleanString = Trim(line)
End Function

Dim method_Name : method_Name = "Script_Bdd" : Call ErrorHandler(method_Name)
