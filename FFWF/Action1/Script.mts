'###########################################################################################################################

	APP_CONFIG_FILE = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\LIBRARY\FFWF_Config.vbs"	' CHANGE THE <APP_Name>
	LoadFunctionLibrary (APP_CONFIG_FILE)
			rc = TESTRUNNER()
	If rc = 0 Then

		Reporter.ReportEvent micInfo, "TESTRUNNER", "Successfully Run the Test Case !" & Chr(13) & "Test Case :   " & UCASE(TESTCASE_NAME)
	
	Else

		Reporter.ReportEvent micFail, "TESTRUNNER", "Failed to Run the Test Case !" & Chr(13) & "Test Case :   " & UCASE(TESTCASE_NAME)
	
	End If

'###########################################################################################################################

'Window("regexpwndtitle:= Parasoft SOAtest").WinTreeView("regexpwndtitle:= SysTreeView32").HighLight
		
'Window("regexpwndtitle:= Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("object class:=SysTreeView32","Location:=0").HighLight @@ hightlight id_;_65826_;_script infofile_;_ZIP::ssf191.xml_;_
'Window("regexpwndtitle:= Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("object class:=SysTreeView32","Location:=0").Click
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"


