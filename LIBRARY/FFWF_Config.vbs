'###########################################################################################################################
'#
'#   <APP_Name>_Config:		Configuration for <APP_Name> Application
'#
'###########################################################################################################################

	Option Explicit		'	-	Force explicit variable declaration

'------------ DECLARE VARIABLES --------------------------------------------------------------------------------------------
	Public AUTOMATION_PATH, APPLICATION_NAME, FRAMEWORK_VERSION, APPLICATION_PATH, FRAMEWORK_PATH, TEST_DATA_INPUT_TYPE
	Public TEST_DATA_SHEET_NAME, TEST_STEP_SHEET_NAME, TEMP_TEST_SHEET_NAME, OBJ_DESC_SHEET_NAME 
	Public TEST_DATA_PATH, TEST_DATA_FILE_NAME, APP_CONSTANT_LIB, APP_FUNCTION_LIB, RESULTS_PATH ,App_Specific_Lib
	Public HTML_HEADER, HTML_RESULT_SUMMARY, HTML_STEP_RESULT_PATH, ERROR_SCREEN_PATH, LOG_PATH
	Public SCRIPT_BDD, FILE_SYS_FUN, FW_DRIVER_LIB, FW_COMMON_LIB, Fw_Keywords_lib, MAIL_ALERT_FLAG, TO_MAIL_LIST
	Public FEATURE_FOLDER, CC_MAIL_LIST, BCC_MAIL_LIST, MAIL_SUBJECT, MAIL_BODY, MAIL_ATTACHMENT, WINDOWS_Keywords_Lib
'---------------------------------------------------------------------------------------------------------------------------
'===========================================================================================================================
'>> MAKE CHANGES BELOW <<===================================================================================================

	AUTOMATION_PATH			= "C:\jenkins\workspace"	' CHANGE THE APPLICATION PATH HERE
	APPLICATION_NAME		= "FFWF"			' CHANGE THE APPLICATION NAME HERE
	FRAMEWORK_VERSION		= "1.0"				' CHANGE FRAMEWORK VERSION HERE (No need to change as of now)
    TEST_DATA_INPUT_TYPE    = "FEATUREFILE"                                                 ' IT CAN BE FeatureFile OR Excel

	Environment.value("TestDataInput_Type")	= TEST_DATA_INPUT_TYPE
'>> DO NOT MAKE CHANGES AFTER THIS LINE <<==================================================================================
'===========================================================================================================================
'---------------------------------------------------------------------------------------------------------------------------
'------------ GET APPLICATION & FRAMEWORK PATH -----------------------------------------------------------------------------
	'On Error Resume Next
	APPLICATION_PATH		= AUTOMATION_PATH & "\FFWF_QA_Automation\" & APPLICATION_NAME	
	FRAMEWORK_PATH			= APPLICATION_PATH & "\"
'------------ Application Library Files Path -------------------------------------------------------------------------------
	APP_CONSTANT_LIB		= APPLICATION_PATH & "\LIBRARY\" & APPLICATION_NAME & "_Constants.vbs"
	APP_FUNCTION_LIB		= APPLICATION_PATH &"\LIBRARY\" & APPLICATION_NAME & "_Functions_Lib.vbs"
	TEST_DATA_PATH			= APPLICATION_PATH & "\DATA\"
	TEST_DATA_FILE_NAME 	= APPLICATION_NAME & "_Data.xlsx" 
	TEST_DATA_SHEET_NAME 	= "Test_Data"
	TEST_STEP_SHEET_NAME	= "Test_Case"
	TEMP_TEST_SHEET_NAME	= "Template"
	OBJ_DESC_SHEET_NAME 	= "Object_Description"
	RESULTS_PATH 			= APPLICATION_PATH & "\RESULTS\"
	HTML_HEADER 			= "TESTCASE_NAME, ITERATION, STATUS, TESTER, TOTAL_STEPS, STEPS_EXECUTED, STEPS_PASSED, " &_
							  "STEPS_FAILED, ERRORS, EXEC_DATE, START_TIME, END_TIME, DURATION, TEST_BROWSER, TEST_ENV, Test_URL"
	HTML_RESULT_SUMMARY 	= RESULTS_PATH & APPLICATION_NAME & "_Result Summary.html"
	HTML_STEP_RESULT_PATH 	= RESULTS_PATH & "Step_Results\"
	ERROR_SCREEN_PATH		= RESULTS_PATH & "Error_Screens\"
	LOG_PATH				= RESULTS_PATH & "Logs\"
'------------ FRAMEWORK LIBRARY FILES PATH ---------------------------------------------------------------------------------
	FW_DRIVER_LIB			= FRAMEWORK_PATH & "FW_Driver_Lib.vbs"
	FW_COMMON_LIB			= FRAMEWORK_PATH & "FW_Common_Lib.vbs"
	WINDOWS_Keywords_Lib	= FRAMEWORK_PATH & "WINDOWS_Keywords_Lib.vbs"
	App_Specific_Lib  =  FRAMEWORK_PATH & "App_Specific_Lib.vbs"
	FILE_SYS_FUN		=FRAMEWORK_PATH & "fso_funcs.qfl"
	SCRIPT_BDD		=FRAMEWORK_PATH & "Script_Bdd.vbs"
	FEATURE_FOLDER		=FRAMEWORK_PATH & "FEATURES" & "\"
'------------ IMPORT ALL FRAMEWORK LIBRARY FILES ---------------------------------------------------------------------------
LoadFunctionLibrary FW_COMMON_LIB

	Reporter.ReportEvent micInfo, "FW_COMMON_LIB", "Successfully Imported the Framework Data Control Script Library :- " & Chr(13) & FW_COMMON_LIB	

LoadFunctionLibrary FW_DRIVER_LIB

	'If TEST_DATA_INPUT_TYPE = 'FeatureFile' then
	

Reporter.ReportEvent micInfo, "FW_DRIVER_LIB", "Successfully Imported the Framework Driver Script Library :- " & Chr(13) & FW_DRIVER_LIB

	LoadFunctionLibrary WINDOWS_Keywords_Lib

	Reporter.ReportEvent micInfo, "WINDOWS_Keywords_Lib", "Successfully Imported the Windows Keywords Library :- " & Chr(13) & WINDOWS_Keywords_Lib

	
	LoadFunctionLibrary FILE_SYS_FUN

	Reporter.ReportEvent micInfo, "FILE_SYS_FUN", "Successfully Imported the Framework File Control Script Library :- " & Chr(13) & FILE_SYS_FUN

	
	LoadFunctionLibrary SCRIPT_BDD

	Reporter.ReportEvent micInfo, "SCRIPT_BDD", "Successfully Imported the Framework File Control Script Library :- " & Chr(13) & SCRIPT_BDD
	
	
	LoadFunctionLibrary App_Specific_Lib

	Reporter.ReportEvent micInfo, "App_Specific_Lib", "Successfully Imported the Framework Data Control Script Library :- " & Chr(13) & App_Specific_Lib

	
'------------ SET EMAIL - FRAMEWORK TEST RESULTS ---------------------------------------------------------------------------
	MAIL_ALERT_FLAG 		= "ON"     		' Set "ON" to receive the email alerts. Set "OFF" to not receive the email alerts.
	TO_MAIL_LIST 			= "rekha.op@centurylink.com" 	' Add recipients by separating "; "
	CC_MAIL_LIST 			= "rekha.op@centurylink.com"
	BCC_MAIL_LIST 			= "shivam.singh@centurylink.com"
	MAIL_SUBJECT 			= APPLICATION_NAME & " - Automated Test Execution Results_" & Date & " " & Time
	MAIL_BODY  				= "This is the Auto Mail Alert from CTL KWH-A Automation Framework." & chr(13) & "Please refer attached HTML Summary result to see the Test Case status. " &_ 
	"Click on the Test Case name to see the Test Step status." & chr(13) & chr(13) & "~" & APPLICATION_NAME & " <APP_Name> Automation Test Team"
	MAIL_ATTACHMENT 		= HTML_RESULT_SUMMARY
'------------ HANDLE ERROR -------------------------------------------------------------------------------------------------
	If Err.Number <> 0 Then
		ERR_MSG = "Error # " & CStr(Err.Number) & " " & Err.Description & Err.Source
		Reporter.ReportEvent micFail, "<APP_Name>_Config", ERR_MSG
		Run_Error = Run_Error & " / " & "<APP_Name>_Config" & " - " & ERR_MSG
		Err.Clear
	End If
'---------------------------------------------------------------------------------------------------------------------------
'*******************   End of <APP_Name>_CONFIG   **************************************************************************
