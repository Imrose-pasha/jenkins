'###########################################################################################################################
'##																				

'										  ##
'##		CTL KWH-A AUTOMATION FRAMEWORK for QTP/UFT													

'					  ##
'##																				

'										  ##		
'## ----------------------------------------------------------------------------------------------------------------------##
'##																				

'										  ##
'##		SCRIPT			:	StartUp()													

'								  ##
'##																				

'										  ##
'##		DESCRIPTION		:	The following script is responsible for setting up the environment and 						
'
' ##
'##					   		then calling the appropriate functions to perform the actual work.					

'		  ##
'##																				

'										  ##
'##		PARAMETERS		:	None														

'								  ##
'##																				

'										  ##
'##		NOTE			:	Change the "<APP_Name>" with the Application/Project name.							

'		  ##
'##																				

'										  ##
'###########################################################################################################################

	APP_CONFIG_FILE = "C:\automation\APPLICATIONS\NTM_Automation_BDD\LIBRARY\NTM_Automation_BDD_Config.vbs"	' CHANGE THE <APP_Name>
	LoadFunctionLibrary (APP_CONFIG_FILE)
	
	rc = TESTRUNNER()
	
	If rc = 0 Then

		Reporter.ReportEvent micInfo, "TESTRUNNER", "Successfully Run the Test Case !" & Chr(13) & "Test Case :   " & UCASE(TESTCASE_NAME)
	
	Else

		Reporter.ReportEvent micFail, "TESTRUNNER", "Failed to Run the Test Case !" & Chr(13) & "Test Case :   " & UCASE(TESTCASE_NAME)
	
	End If

'###########################################################################################################################