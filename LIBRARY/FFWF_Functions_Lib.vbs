'###########################################################################################################################
'#
'#   <APP_Name>_Functions_Lib:	Functions Specific to <APP_Name> Application
'#
'#__________________________________________________________________________________________________________________________
'#			KEYWORDS				  	PARAMETERS
'#__________________________________________________________________________________________________________________________
'#			1.	function_1				para1, para2, para3
'#			2.	function_2				para1, para2
'#			3.	function_3				para1
'#__________________________________________________________________________________________________________________________

'___________________________________________________________________________________________________________________________
'# Function Name	:	function_2()
'# Purpose			:	
'# Parameters		:	para1		-> Description of para1
'#						para2		-> Description of para2
'# Return Code		:	0  -> Success
'#						-1 -> Failure
'# Note				:
'___________________________________________________________________________________________________________________________

Public Function function_2(para1, para2)
	On Error Resume Next
	Dim methodName
	Function_Name = 0 : methodName = "function_2"

	'Script to generate Test Step Description and Expected Result
	Step_Description = "" 
	Exp_Result = ""

	If Exec_Flag = "Y" Then
		'Write statements to execute
		Actual_Res = ""
	Else
		Actual_Res = ""
	End If

	' Handling Error
	methodName = "function_2" : rc = ErrorHandler(methodName)
End Function

Dim method_Name : method_Name = "FFWF_Functions_Libs" : Call ErrorHandler(method_Name)
'*******************   End of <APP_Name>_Functions_Lib   *******************************************************************
