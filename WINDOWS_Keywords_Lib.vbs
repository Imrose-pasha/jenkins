'###########################################################################################################################
'#
'#   WEB_KEYWORDS_LIB:		Contains Common functions can be used by multiple applications
'#__________________________________________________________________________________________________________________________
'#		KEYWORDS			PARAMETERS
'#__________________________________________________________________________________________________________________________
'#		1.	Login_SOA	
'#		2.  closeAllBrowser
'#		3. 	Installorder_xml
'#		4.	OPEN								TEST_ENV, pageDesc
'#		5.	MDWDesigner_Login 					UserId,Password
'#		6.	Select_Process 						ProcessName
'#		7.	Load_ProcessInstance
'#		8.	Loading_MasterRequestId 			BAN,ProcessStatus
'#		9.	Generic_CLICK1 						inputData
'#		10.	Generic_Verify						inputData,inputData1
'#		11.	INPUT								pageDesc, objectDesc, inputData
'#		12.	Buslistnerorder_xml					EventName
'#		13.	Generic_CLICK						inputData
'#		14.	Generic_Input1 						inputData,inputData1
'#		15.	Generic_Input 						inputData,inputData1
'#		16.	XMLValue_Collect 	
'#		17.	ValidateXMLTags						XmlTagName
'#		18.	ReadXMLFileAndReplaceXMLTags 		QID,FileLocation,Tagtobereplaced,TagValue
'#		19.	FFWF_DB_Connect 					TransId
'#		20.	CLOSE_								pageDesc
'#		21.	Copy_xmlvaluetoexcel				row,col,inputdata
'#		22.	ToUpdateDateinXML					XmlFileName
'#		23.	Validatecomplete_MasterRequestId 	BAN
'#		24.	CopyTextToXML
'#		25. ToUpdateTrackingNumberinXML			XmlFileName
'#		26.	Copy_ExcelvaluetoExcel				row,col
'#		27.	ReadCellvalueFromExcel				row,col
'#		28.	IOM_Login							UserId,Password,EventName,inputData3

			
'#__________________________________________________________________________________________________________________________

'___________________________________________________________________________________________________________________________
'# Function Name	: Login_SOA()
'# Purpose			: Launch SOA and Navigate to the Invoke Web Service
'# Parameters		: 
'#					  
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Login_SOA()

	On Error Resume Next
	Dim methodName, rc, Temp_Name
	methodName = "Login_SOA" : Login_SOA = 0
	Execute("Test_URL = " & TEST_ENV)

	'To generate Test Step Description and Expected Result
	'Loc_name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
	Step_Description = "Open SOA " 
	Exp_Result = " SOA should open Successfully " 
	
'	If Exec_Flag = "Y" Then
'	'Call close all open browser function
'		Call closeAllBrowser
'		If TEST_BROWSER = "IE" Then 
'			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","",3
'		'ElseIf 	TEST_BROWSER = "Firefox" Then 
'			'SystemUtil.Run "firefox.exe",Test_URL,"C:\","", 3
'		'ElseIf 	TEST_BROWSER = "GChrome" Then 
'			'SystemUtil.Run "chrome.exe",Test_URL,"C:\","",3
'		Else 
'			TEST_BROWSER = "IE"		'	Default Browser
'			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","", 3
'		End If
'
'		Browser("CreationTime:=0").Sync

Set fileSystemObj = createobject("Scripting.FileSystemObject")

'To check if the given file present'

MyFile = "C:\Program Files\Parasoft\SOAtest\9.9\soatest.exe"
'	MyFile = "C:\Program Files\Parasoft\SOAtest\9.10\soatest.exe"

		If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		Else
		    rc = "False"
		End If

		If rc = "True" Then
			Print "SOA is invoking:::"
			systemutil.Run"C:\Program Files\Parasoft\SOAtest\9.9\soatest.exe"
			'systemutil.Run"C:\Program Files\Parasoft\SOAtest\9.10\soatest.exe"
			Wait(9)
			Dialog("text:=Workspace Launcher","Location:=0").winbutton("text:=OK").Click
			'Window("text:=Eclipse Launcher").WinButton("text:=OK").Click   'This is for v9.10
			Wait(25)
			Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0").WinObject("nativeclass:=SysLink","regexpwndtitle:=.*Click here to activate license.*","attached text:=License is not active.*","index:=0").Click
			Print "License accpeted:::"
			Wait(20)
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst"
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite"
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite"
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: MDWWebServicePort"
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Select "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: MDWWebServicePort;Test 1: invokeWebService(string, string)"
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Activate "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: MDWWebServicePort;Test 1: invokeWebService(string, string)"
				
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst"
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite"
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: Test Suite"
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: Test Suite;Test Suite: MDWWebServicePort"
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Select "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: Test Suite;Test Suite: MDWWebServicePort;Test 1: invokeWebService(string, string)"
				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Activate "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: Test Suite;Test Suite: MDWWebServicePort;Test 1: invokeWebService(string, string)"
				
			Actual_Res = "Launched " & " SOA with Application Test URL -> " & Chr(13) & Test_URL & ". SOA is Launched -> " & Temp_Name
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else
			Call captureScreen
			Actual_Res = "Launched " & TEST_BROWSER & " SOA with Application Test URL -> " & Chr(13) & Test_URL & ". SOA is not Launched -> " & Temp_Name & ". Page Expected is " & Page_Name(1)
			Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
			Login_SOA = -1
		End If

wait(2)
Print "Login_SOA is completed:::"
					
'End if

Handling Error
	methodName = "Login_SOA" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function Login_SOA() -----------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name    : closeAllBrowser()
'# Purpose          : General Application close function
'# Usage	    	: close all the open browser
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function closeAllBrowser()
	On Error Resume Next
	Dim methodName, rc, Results_Msg	', oBrowser, oCol, NOB, i
	methodName = "closeAllBrowser" : closeAllBrowser = 0
	
	If TEST_BROWSER = "IE" Then 
		SystemUtil.CloseProcessByName("iexplore.exe")
	ElseIf 	TEST_BROWSER = "Firefox" Then 
		SystemUtil.CloseProcessByName("firefox.exe")
	ElseIf 	TEST_BROWSER = "GChrome" Then 
		SystemUtil.CloseProcessByName("chrome.exe")
	Else 
		TEST_BROWSER = "IE"		'	Default Browser
		SystemUtil.CloseProcessByName("iexplore.exe")
	End If
		
	rc = Browser("name:=.*").Exist(0)
	If rc = "True" Then
		Browser("name:=.*").Close
		End If
	Call clearCache
	If rc = "False" Then
		Reporter.ReportEvent micPass, methodName, "Successfully closed all the open browsers."
	Else
		Results_Msg = "Unable to close all the open browsers !! "
		Reporter.ReportEvent micfail, methodName, Results_Msg
		closeAllBrowser = -1
	End If
	' handle error
	methodName = "closeAllBrowser" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function closeAllBrowser() ------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name    :Installorder_xml()
'# Purpose          : To Load Data from XML file
'# Usage	    	:To Load Data from XML files
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
	
	Public Function Installorder_xml()
	On Error Resume Next
	Dim methodName, rc, Temp_Name
	methodName = "Installorder_xml" : Installorder_xml = 0
	
	Print "Installorder_xml function is started:::"
	Execute("Test_URL = " & TEST_ENV)


	'To generate Test Step Description and Expected Result
	Step_Description = "Load XML from the Test Data Sheet" 
	Exp_Result = " XML is loaded successfully from the Test Data Sheet"
	
If Exec_Flag = "Y" Then
	'Call close all open browser function
'		Call closeAllBrowser
'
'		If TEST_BROWSER = "IE" Then 
'			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","",3
'		'ElseIf 	TEST_BROWSER = "Firefox" Then 
'			'SystemUtil.Run "firefox.exe",Test_URL,"C:\","", 3
'		'ElseIf 	TEST_BROWSER = "GChrome" Then 
'			'SystemUtil.Run "chrome.exe",Test_URL,"C:\","",3
'		Else 
'			TEST_BROWSER = "IE"		'	Default Browser
'			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","", 3
'		End If

		Set fileSystemObj = createobject("Scripting.FileSystemObject")

		'To check if the given file present'
		
		MyFile = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\FFWF_Data.xlsx"

		If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		Else
		    rc = "False"
		End If

		If rc = "True" Then			
			Set oExcel=CreateObject("Excel.Application")
			oExcel.Visible=True
			Set oBook=oExcel.Workbooks.Open("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\FFWF_Data.xlsx")
			Set oSheet=oBook.Worksheets("Test_Data")

			rows=oSheet.UsedRange.rows.count
			Cols=oSheet.UsedRange.Columns.count

				Event_Name=oSheet.Cells(i,6).value
    
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Highlight
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").click
				Set ws=CreateObject("wscript.shell")
		      	ws.SendKeys ("^a")	
			    ws.SendKeys "{DELETE}"
			    Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Type Event_Name	
			    Set ws=Nothing 
	     
			Set xmlDoc = CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = False 
			
			Set fso=createobject("Scripting.FileSystemObject")
			
			'Set qfile=fso.OpenTextFile("C:\Automation\APPLICATIONS\FFWF\XML\AVS_Order_Fulfillment.xml",1)
			Set qfile=fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&Event_Name&".xml",1)
			Exml=qfile.ReadAll
			
			Set qfile=nothing
			Set fso=nothing
			
			Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").Highlight
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").click
				Set ws=CreateObject("wscript.shell")
		      	ws.SendKeys ("^a")	
			    ws.SendKeys "{DELETE}"
			    
		
				wait (2)
				
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","location:=2").Type Exml
	     		Set ws=Nothing 
		 
				wait (2)

				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinToolbar("nativeclass:=ToolbarWindow32","location:=8").Press 1
				
				wait (2)
					If Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("text:=Save Resource").WinButton("text:=&Yes").Exist(5) Then
					Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("text:=Save Resource").WinButton("text:=&Yes").Click
					End if 

					 	Actual_Res = "Service Name and Request Details Updated as per the testcase" 
						Reporter.ReportEvent micPass, StepName, Actual_Res
		Else
				Call captureScreen
				Actual_Res = "Service Name and Request Details NOT Updated as per the testcase" 
				Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
				Installorder_xml= -1 
			
		End if

				if Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("attached text:=Finished","regexpwndtitle:=1/1 Tests Succeeded").Exist Then
				   Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("attached text:=Finished","regexpwndtitle:=1/1 Tests Succeeded").Highlight
		 				Print "Tests Succeeded static text is validated"
		 				Actual_Res = "Tests Succeeded static text is validated" 
						Reporter.ReportEvent  micPass, StepName, Actual_Res
				Else
						Call captureScreen
						Print "Tests Succeeded static text is Not Displayed"
						Actual_Res = "Tests Succeeded static text is Not Displayed" 
						Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
						Installorder_xml= -1 
				End if
	
				if Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("nativeclass:=Static","regexpwndtitle:=No Tasks Reported").Exist Then
				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("nativeclass:=Static","regexpwndtitle:=No Tasks Reported").Highlight
						 	Actual_Res = "No Tasks Reported static text is validated" 
							Reporter.ReportEvent micPass, StepName, Actual_Res
				Else
						Call captureScreen
						Actual_Res = "No Tasks Reported static text is Not Displayed" 
						Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
						Installorder_xml= -1 
				End if
				
		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest.*","Location:=0").Highlight
		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest.*","Location:=0").Close
wait(2)
		If Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("regexpwndtitle:=Confirm Exit","text:=Confirm Exit").WinButton("text:=OK").Exist(5) Then
					Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("regexpwndtitle:=Confirm Exit","text:=Confirm Exit").WinButton("text:=OK").Click
					End if 
wait(2)


oBook.Close
oExcel.Quit

Print "Function Installorder_xml is completed:::"

End if
End Function
' --------------------------- End of Function Installorder_xml() ------------------------------------------------------------


'___________________________________________________________________________________________________________________________
'# Function Name	: OPEN()
'# Purpose			: Invoke the Browser & opens the Test URL
'# Parameters		: TEST_ENV	-> Gets the URL to be opened
'#					  pageDesc	-> Gets the Page Description of URL to be opened
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public BROWSER_NAME
Public Function OPEN(TEST_ENV, pageDesc)
	On Error Resume Next
	Dim methodName, Page_Name, pageName, rc, Temp_Name
	methodName = "OPEN" : OPEN = 0
	Execute("Test_URL = " & TEST_ENV)

	'To generate Test Step Description and Expected Result
	Page_Name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
	Step_Description = "Open Test URL in " & TEST_BROWSER & " Browser."
	Exp_Result = TEST_BROWSER & " Browser should open with the page -> " & pageName
	
	If Exec_Flag = "Y" Then
	'Call close all open browser function
'		Call closeAllBrowser
		If TEST_BROWSER = "IE" Then 
			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","",3
		ElseIf 	TEST_BROWSER = "Firefox" Then 
			SystemUtil.Run "firefox.exe",Test_URL,"C:\","", 3
		ElseIf 	TEST_BROWSER = "GChrome" Then 
			SystemUtil.Run "chrome.exe",Test_URL,"C:\","",3
		Else 
			TEST_BROWSER = "IE"		'	Default Browser
			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","", 3
		End If
	
		rc = Browser("CreationTime:=0").Exist(10)
		'msgbox rc
		rc = "True"
		
		If rc = "True" Then
			
			Actual_Res = "Opened browser" & " With Application Test URL -> " & Chr(13) & Test_URL & ". Page Opened is -> " & pageName
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else
			'Call captureScreen
			
			Actual_Res = "Opened " & TEST_BROWSER & " Browser with Application Test URL -> " & Chr(13) & Test_URL & ". Page  is not launched"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			OPEN = -1
		End If

	End If
	' Handling Error
	methodName = "OPEN" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function OPEN() -----------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: MDWDesigner_Login()
'# Purpose			: To Login to MDW Designer with the given user id and password
'# Parameters		: inputData		-> Enter the username
'#					  inputData1	-> Enter the password
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________


Public Function MDWDesigner_Login(inputData,inputData1)
	On Error Resume Next
	Dim methodName, rc, display
	methodName = "MDWDesigner_Login" : MDWDesigner_Login = 0
	Step_Description = "MDW Designer should be Launched and Logged in Successfully"
	Exp_Result = "MDW Designer is Launched and Logged in Successfully"

If Exec_Flag = "Y" Then
wait(5)

		'	Do
		'		display = JavaDialog("title:=Security Warning","toolkit class:=javax.swing.JDialog").JavaCheckBox("attached text:=I accept the risk and.*","toolkit class:=javax\.swing\.JCheckBox").Exist(10)

		'	Loop While display = False
                                  
		 If False Then
		 
		 'JavaDialog("title:=Security Warning","toolkit class:=javax.swing.JDialog").JavaCheckBox("attached text:=I accept the risk and.*","toolkit class:=javax\.swing\.JCheckBox").Exist(10) 
		 	JavaDialog("title:=Security Warning","toolkit class:=javax.swing.JDialog").JavaCheckBox("attached text:=I accept the risk and.*","toolkit class:=javax\.swing\.JCheckBox").Highlight
			JavaDialog("title:=Security Warning","toolkit class:=javax.swing.JDialog").JavaCheckBox("attached text:=I accept the risk and.*","toolkit class:=javax\.swing\.JCheckBox").Set "ON"
		    JavaDialog("title:=Security Warning","toolkit class:=javax.swing.JDialog").JavaButton("attached text:=Run","toolkit class:=javax\.swing\.JButton").Click

'	 		Else
'	 		MDWDesigner_Login=-1
		Else
	 	End If
	
	rc =JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").Exist(60) 

		If rc ="True" Then

			JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").JavaEdit("attached text:=User Name").Highlight
			JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").JavaEdit("attached text:=User Name").set inputData

			JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").JavaEdit("attached text:=Password").SetFocus
			wait(1)
			JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").JavaEdit("attached text:=Password").Set inputData1
					wait(1)
				if JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").JavaButton("attached text:=Log In").Getroproperty("enabled") = "1" Then
					JavaWindow("title:=MDW Designer","toolkit class:=com.qwest.mdw.designer.MainFrame").JavaButton("label:=Log In").Click
				End if

			Do
				display = JavaWindow("title:=MDW Designer.*","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").Exist

			Loop While display = False

			if display = True Then
					Step_Description = "Login to MDW Designer in test environment"
					Exp_Result = "Login to MDW Designer should be successfull"
					Actual_Res = "Login to MDW Designer is successfull"
					Reporter.ReportEvent micPass, StepName, Actual_Res
			Else
			Step_Description = "Login to MDW Designer in test environment"
			Exp_Result = "Login to MDW Designer should be successfull"
			Actual_Res = "Login to MDW Designer failed"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			MDWDesigner_Login=-1
			End If
	
		Else
			Step_Description = "Login window display"
			Exp_Result = "Login window should be displayed"
			Actual_Res = "Login window is displayed"
			Reporter.ReportEvent micPass, StepName, Actual_Res
			End if
	
	End If
	' #comments: Handling Error
	methodName = "MDWDesigner_Login" : rc = ErrorHandler(methodName)
End Function

' --------------------------- End of MDWDesigner_Login() ----------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: Select_Process(inputData)
'# Purpose			: To Select the Process from the Process list based on the input
'# Parameters		: inputData -> Enter Which process you want to select
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Select_Process(inputData)
	On Error Resume Next
	Dim methodName,rc, i, process_name, row_cnt, display
	methodName = "Select_Process" : Select_Process = 0
	
	If Exec_Flag = "Y" Then
	JavaWindow("MDW Designer.*").JavaTable("JTable").highlight

		JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax.swing.JTable").highlight
		rc = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax.swing.JTable").Exist(10)

			If rc = "True" Then
	
				row_cnt = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax.swing.JTable").GetROProperty("rows")
				Print "Selecting requierd process from MDW:::"
				For i = 0 To row_cnt-1 step 1
				'Print "Iteration no ="&i
					process_name = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax.swing.JTable").GetCellData(i,1)
					'print "process name: ="& process_name
						If Trim(Ucase(process_name)) = Trim(ucase(inputData)) Then
						Step_Description = "To select the required process"
							Exp_Result = "Required Process should be displayed"
							Actual_Res = "Required Process is displayed"
							Reporter.ReportEvent micPass,StepName, Actual_Res
							
							JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax.swing.JTable").ClickCell i,1
							JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax.swing.JTable").DoubleClickCell i,1
							Flag = True
						Exit for	
						Else
							
							Step_Description = "To verify update status"
							Exp_Result = "Required Process should be displayed"
							Actual_Res = "Required Process is not displayed"
							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
							Select_Process=-1
							
						End If
				Next
						
						Do
						display = JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=DesignerCanvas","toolkit class:=com.qwest.mdw.designer.pages.DesignerCanvas").Exist(5)
						Loop While display = "False"
			
						If display = "True" Then
						
							Step_Description = "To Verify if the selected process is loaded"
							Exp_Result = "Selected Process should be loaded"
							Actual_Res = "Selected Process is loaded sucessfully"
							Reporter.ReportEvent micPass, StepName, Actual_Res
						Else
							Call captureScreen
							Step_Description = "To Verify if the selected process is loaded"
							Exp_Result = "Selected Process should be loaded"
							Actual_Res = "Selected Process is not loaded"
							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
							Select_Process=-1	
						End If
						
				If  JavaWindow("title:=MDW Designer.*").JavaDialog("tagname:=Error","toolkit class:=javax.swing.JDialog").Exist(5) Then
					
					JavaWindow("title:=MDW Designer.*").JavaDialog("tagname:=Error","toolkit class:=javax.swing.JDialog").JavaButton("tagname:=OK","toolkit class:=javax.swing.JButton").Click
					JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Highlight
					rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Exist(10)
					JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Press(5)
					JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").Exist(10)  
					JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").JavaButton("attached text:=Load").Click
					wait (2)
					
					JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(5) 
					 wait (1)
'					Step_Description = "To Verify if the selected process is loaded"
'					Exp_Result = "Selected Process should be loaded"
'					Actual_Res = "Selected Process is loaded sucessfully"
'					Reporter.ReportEvent micPass, StepName, Actual_Res 
					 
					 			Else
 	Call captureScreen
	Step_Description = "To Verify if the selected process is loaded"
	Exp_Result = "Selected Process should be loaded"
	Actual_Res = "Selected Process is not loaded"
	Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
	Select_Process=-1	
	
			  End if
			Else
 	Call captureScreen
	Step_Description = "To Verify if the selected process is loaded"
	Exp_Result = "Selected Process should be loaded"
	Actual_Res = "Selected Process is not loaded"
	Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
	Select_Process=-1				
								
			End If
	End if
	' Handling Error
	methodName = "Select_Process" : Select_Process = ErrorHandler(methodName)
End function
' --------------------------- End of Select_Process() ----------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: Load_ProcessInstance()
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		: 
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Load_ProcessInstance()
	On Error Resume Next
	Dim methodName,rc, Flag
	methodName = "Load_ProcessInstance" : Load_ProcessInstance = 0
	Flag = False
	 	If Exec_Flag = "Y" Then

		'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Highlight
		'rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Exist(10)

			'If rc = "True" Then
			'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Press(5)
			'wait (5)
				'JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaButton("tagname:=table24","toolkit class:=javax\.swing\.JButton").Click
					'If  JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").Exist(2) Then
					'wait (2)
						'JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").JavaButton("tagname:=OK","toolkit class:=javax.swing.JButton").Click
						'wait (2)
						'Else
										
					'If JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").Exist(10) Then 
					'JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").JavaButton("attached text:=Load").Click
					'wait (2)
					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(5) Then
					wait (3)
							Flag = True
							Step_Description = "To Verify if Filter Process Instances Window is Opened"
							Exp_Result = "Filter Process Instances Window should be Loaded"
							Actual_Res = "Filter Process Instances Window is Loaded"
							Reporter.ReportEvent micPass, StepName, Actual_Res
						Else
							Call captureScreen
							Step_Description = "To Verify if Filter Process Instances Window is Opened"
							Exp_Result = "Filter Process Instances Window should be Loaded"
							Actual_Res = "Filter Process Instances Window is not Loaded"
							Reporter.reportevent micPass,StepName, Actual_Res, ERROR_SCREEN_FILE
							Load_ProcessInstance=-1
					End If
					
						
					Flag = True
							Step_Description = "To Verify if the Total Instances Window is Opened"
							Exp_Result = "Total Instances Window should be Loaded"
							Actual_Res = "Total Instances Window is Loaded"
							Reporter.ReportEvent micPass, StepName, Actual_Res
						Else
							Call captureScreen
							Step_Description = "To Verify if the Total Instances Window is Opened"
							Exp_Result = "Total Instances Window should be Loaded"
							Actual_Res = "Total Instances Window is not Loaded"
							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
							Load_ProcessInstance=-1
					End If
			'End If		
			'End If
	'End if
' Handling Error
	methodName = "Load_ProcessInstance" : rc = ErrorHandler(methodName)
End function

'' --------------------------- End of Load_ProcessInstance() ----------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: Loading_MasterRequestId(inputData,inputData1)
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		: 
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Loading_MasterRequestId(inputData,inputData1)
		On Error Resume Next
	Dim methodName,rc, i, master_reqid, row_cnt
	methodName = "Loading_MasterRequestId" : Loading_MasterRequestId = 0
	'	Exec_Flag = "Y"
	'inputData = "100001601"
	'inputData1 = "In Progress"
	
		If Exec_Flag = "Y" Then
	
			JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").highlight
			rc = JavaWindow("title:=MDW Designer .*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(10)
	
					If rc = "True" Then
			
						row_cnt = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
						For i = 0 To row_cnt-1
							master_reqid = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
							'msgbox master_reqid
														If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
														print "Entered master request id found:::" &inputData
															'get the value of the status code
															status_code = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,4)
																					If Trim(Ucase(status_code)) = Trim(ucase(inputData1)) Then
																					print "Order status code is found" &status_code
																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").ClickCell i,1
																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell i,1
																				Flag = True
																					else
																					'Call captureScreen
																					Step_Description = "The Status of the selected Master Request id is not in progress"
																					Exp_Result = "The Status of the selected Master Request id is in progress"
																					Actual_Res = "The Status of the selected Master Request id is not in progress"
																					Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
																					Loading_MasterRequestId=-1
																					End if
															Step_Description = "To select the required Master Request id"
															Exp_Result = "Master Request Id should be displayed"
															Actual_Res = "Master Request Id is displayed"
															Reporter.ReportEvent micPass, StepName, Actual_Res
								Exit For						
								Else
															'Call captureScreen
'															Step_Description = "To select the required Master Request id"
'															Exp_Result = "Master Request Id should be displayed"
'															Actual_Res = "Master Request Id is not displayed"
'															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'															Loading_MasterRequestId=-1
'												
														End If
						Next
								
								Do
								JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Highlight
								display = JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Exist
								Loop While display = False
					
														If display = "True" Then
															Step_Description = "To Verify if the process for the selected master request id is loaded"
															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
															Actual_Res = "Process is Loaded successfully for given Master Request Id" &inputData
															Reporter.ReportEvent micPass, StepName, Actual_Res
														Else
															'Call captureScreen
															Step_Description = "To Verify if the process for the selected master request id is loaded"
															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
															Actual_Res = "Process is not Loaded successfully for given Master Request Id" &inputData
															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
															Loading_MasterRequestId=-1	
														End If
					End If
				End if
			' Handling Error
			Print "Loading_MasterRequestId is done"
	methodName = "Loading_MasterRequestId" : rc = ErrorHandler(methodName)
End function
' --------------------------- End of Loading_MasterRequestId() ----------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: Generic_CLICK1(inputData)
'# Purpose			: To click an object in the application
'#					  objectDesc	-> Gets the Description of the Object to be clicked
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Generic_CLICK1(inputData)
	On Error Resume Next
	Dim methodName,ObjDesc_Array, Object_Arr, Object_Name, objName, rc, exec_Stmt ,exec_stmnt, Flag,val,val1,Object_Desc
	methodName = "Generic_CLICK1" : Generic_CLICK1 = 0
	Flag = False

	ObjDesc_Array = split(inputData,"=>")
	'msgbox ObjDesc_Array(0)
	ObjSplit_Array = split(ObjDesc_Array(1)," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
	Select Case (ObjSplit_Count)
		Case "0"
			Object_Desc = ObjDesc_Array(1)
			
		Case "1"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
			'msgbox Object_Desc
		Case "2"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
			
		Case "3"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
			
		Case "4"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
			
		Case Else
			Step_Description = "Object property"
			Actual_Res = "Object property exceeded more than five"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			Generic_CLICK1 = -1
	End Select

	'To generate Test Step Description and Expected Result
	Step_Description = "Click on object -> " & objName & "  in page generic page"
	Exp_Result = "Clicked on object ->"& objName & " in generic page"

	If Exec_Flag = "Y" Then
		exec_stmnt = "rc = Browser(""creationtime:=2"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Exist(""10"")"
		'msgbox exec_stmnt
		Execute(exec_stmnt)
		'msgbox rc
		If rc = "True" Then
			'Browser("CreationTime:=2").Page("title:=.*").highlight
			exec_Stmt = "Browser(""creationtime:=2"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Click"
			Execute(exec_Stmt)
			Flag = True
			Actual_Res = "Click on object -> " & objName & " in generic page is successful"
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else Generic_CLICK1 = -1 End If
If Flag = False Then
	Actual_Res = "Click on object -> " & objName & " in generic page failed"
	Reporter.ReportEvent micFail, StepName, Actual_Res
	Generic_CLICK1 = -1
End If
			
End If
	' Handling Error
	methodName = "Generic_CLICK1" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function Generic_CLICK1() ----------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: Generic_Verify(inputData,inputData1)
'# Purpose			: To click an object in the application
'#					  objectDesc	-> Gets the Description of the Object to be clicked
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Generic_Verify(inputData,inputData1)
	On Error Resume Next
	Dim methodName,ObjDesc_Array, Object_Arr, Object_Name, objName, rc, exec_Stmt ,exec_stmnt, Flag,val,val1,Object_Desc, verify_msg
	methodName = "Generic_Verify" : Generic_Verify = 0
	Flag = False

	ObjDesc_Array = split(inputData,"=>")
	'msgbox ObjDesc_Array(0)
	ObjSplit_Array = split(ObjDesc_Array(1)," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
	Select Case (ObjSplit_Count)
		Case "0"
			Object_Desc = ObjDesc_Array(1)
			
		Case "1"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
			'msgbox Object_Desc
		Case "2"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
			
		Case "3"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
			
		Case "4"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
			
		Case Else
			Step_Description = "Object property"
			Actual_Res = "Object property exceeded more than five"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			Generic_Verify = -1
	End Select

	'To generate Test Step Description and Expected Result


	If Exec_Flag = "Y" Then
		exec_stmnt = "rc = Browser(""creationtime:=0"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Exist(""10"")"
		
		'msgbox exec_stmnt
		Execute(exec_stmnt)
		
		if rc = "True" Then
		
		exec_stmnt = "verify_msg = Browser(""creationtime:=0"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Getroproperty(""innertext"")"
		Execute(exec_stmnt)
			if instr(Trim(ucase(verify_msg)),Trim(Ucase(inputData1)))<>0 Then
			         	Step_Description = "To verify-> " & ObjDesc_Array(0) & "  in  generic page"
				Exp_Result = "verification of->"& ObjDesc_Array(0) & " is successful in generic page"
				Actual_Res = "verification is successful and the verification message is-> "& 	verify_msg
			Else
				Step_Description = "To verify-> " & ObjDesc_Array(0) & "  in  generic page"
				Exp_Result = "verification of->"& ObjDesc_Array(0) & " is successful in generic page"
				Actual_Res = "verification is not successful and the verification message is-> "& verify_msg
				Generic_Verify=-1
			End if
		Else
			Generic_Verify=-1
		End if
End if
	' Handling Error
	methodName = "Generic_Verify" : rc = ErrorHandler(methodName)
End Function

' --------------------------- End of Function Generic_Verify() ----------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: INPUT(pageDesc, objectDesc, inputData)
'# Purpose			: To Enter Data to object(s) in the application
'# Parameters 		: pageDesc		-> Gets the Page Description of of the Objects(s) needs input
'#					  objectDesc	-> Gets the Description(s) of the Objects(s) needs input
'#					  inputData		-> Gets the Data for the Objects(s)
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function INPUT(pageDesc, objectDesc, inputData)
	On Error Resume Next
	Dim methodName, Page_Name, pageName, ObjDesc_Array, Object_Count, Data_Array, Data_Count, i, Object_Entered, Object_NotExist, Object_Name, Obj_Exist
	Dim Obj_Array, Object_Arr, In_Operation, exec_Stmt, rc, Input_Data, Input_Object, Data_Entered
	methodName = "INPUT" : INPUT = 0
	
	Page_Name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
	ObjDesc_Array = split(objectDesc,"++",-1,1) : Object_Count = UBound(ObjDesc_Array)
	Data_Array = split(inputData,"++",-1,1) : Data_Count = UBound(Data_Array)
	If Object_Count = Data_Count Then
		For i=0 to Object_Count
			Obj_Array = split(ObjDesc_Array(i),"=>",-1,1) : Object_Arr = split(Obj_Array(1)," + ",-1,1) : Object_Name = split(Object_Arr(0),":=",-1,1)
			If i=0 Then
				Input_Data = Data_Array(i) : Input_Object = Object_Name(1)
			Else
				Input_Data = Input_Data & ", " & Data_Array(i) : Input_Object = Input_Object & ", " & Object_Name(1)
			End If
		Next
	End If
	'To generate Test Step Description and Expected Result
	Step_Description = "Enter value for object(s) -> " & Input_Object & " in page -> " & pageName
	Exp_Result = Input_Data & " -> value(s) should be entered for object(s) -> " & Input_Object & " in page -> " & pageName
	
	If Exec_Flag = "Y" Then
		Object_Entered = "" : Object_NotExist = ""
		If Object_Count = Data_Count Then
			For i=0 to Object_Count
				If Data_Array(i) <> "" Then
					Object_Name = "" : Obj_Exist = 0
					Obj_Array = split(ObjDesc_Array(i),"=>",-1,1) : Object_Arr = split(Obj_Array(1)," + ",-1,1) : Object_Name = split(Object_Arr(0),":=",-1,1)
					rc = EXIST_(pageDesc, ObjDesc_Array(i))
					If rc = 0 Then
						If Obj_Array(0) = "WebList" OR Obj_Array(0) = "WebRadioGroup" Then
							In_Operation = "Select "
						Else
							In_Operation = "Set "
						End If
						Setting.WebPackage("ReplayType") = 2
						exec_Stmt = "Browser(""name:=" & pageName & """).Page(""title:=" & pageName & """)." & Obj_Array(0) & "(""" & Object_Desc & """)." & In_Operation & """" & Data_Array(i) & """"
						execute(exec_Stmt)
						Setting.WebPackage("ReplayType") = 1
					Else Obj_Exist = -1 End If
					If Obj_Exist = 0 Then
						If Object_Entered = "" AND Data_Entered = "" Then
							Object_Entered = Object_Name(1) : Data_Entered = Data_Array(i)
						Else
							Object_Entered = Object_Entered & ", " & Object_Name(1)
							Data_Entered = Data_Entered & ", " & Data_Array(i)
						End If
					ElseIf Obj_Exist = -1 Then
						If Object_NotExist = "" Then Object_NotExist = Object_Name(1) Else Object_NotExist = Object_NotExist & ", " & Object_Name(1) End If
					End If
			 	End If
		 	Next
			If Object_Entered <> "" And Object_NotExist = "" Then
				Actual_Res = "Entered value(s) -> " & Data_Entered & " for object(s) -> " & Object_Entered & " in page -> " & pageName & ". "
				Reporter.ReportEvent micPass, StepName, Actual_Res
			ElseIf Object_Entered <> "" And Object_NotExist <> "" Then
				Call captureScreen
				Actual_Res = "Entered value(s) -> " & Data_Entered & " for object(s) -> " & Object_Entered & ". " & vbnewline & "Object(s) -> " & Object_NotExist & " NOT identified in page -> " & pageName
				Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
				INPUT = -1
			ElseIf Object_Entered = "" And Object_NotExist <> "" Then
				Call captureScreen
				Actual_Res =  "Object(s) -> " & Object_NotExist & " NOT identified in page -> " & pageName
				Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
				INPUT = -1
			ElseIf Object_Entered = "" And Object_NotExist = "" Then
				Reporter.ReportEvent micFail, StepName, Actual_Res
				INPUT = -1
			End If  
		Else
			Actual_Res = "Object Count :-  " & Object_Count & " AND Data Count :-  " & Data_Count & "    Doesn't match ???"
			Reporter.reportevent micFail,StepName, Actual_Res
			INPUT = -1
		End If
	End If
	' Handling Error
	methodName = "INPUT" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function INPUT() ----------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name    :Buslistnerorder_xml(inputdata)
'# Purpose          : To Load Data from XML file
'# Usage	    	:To Load Data from XML file
'# Return	    	: 0  : Success
'#         	     	  -1 : Failure
'___________________________________________________________________________________________________________________________
	Public Function Buslistnerorder_xml(inputdata)
	On Error Resume Next
	Dim methodName, rc, Temp_Name,xmlDoc,filename,qfile,Responsexml,fso,ws
	methodName = "Buslistnerorder_xml" : Buslistnerorder_xml = 0
	Execute("Test_URL = " & TEST_ENV)


	'To generate Test Step Description and Expected Result
	Step_Description = "Load XML from the XML file" 
	Exp_Result = " XML should be loaded successfully from the XML file"

	
If Exec_Flag = "Y" Then
	'Call close all open browser function
'		Call closeAllBrowser
'
'		If TEST_BROWSER = "IE" Then 
'			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","",3
'		'ElseIf 	TEST_BROWSER = "Firefox" Then 
'			'SystemUtil.Run "firefox.exe",Test_URL,"C:\","", 3
'		'ElseIf 	TEST_BROWSER = "GChrome" Then 
'			'SystemUtil.Run "chrome.exe",Test_URL,"C:\","",3
'		Else 
'			TEST_BROWSER = "IE"		'	Default Browser
'			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","", 3
'		End If
'inputdata="Resp.xml"
	If Browser("name:=TMS Bus Listener","title:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("html id:=tbRequest","html tag:=TEXTAREA","name:=tbRequest").Exist(2) Then
		    rc = "True"
		   		Else
		    rc = "False"
		    		End If
	
		If rc = "True" Then			
			
			Set xmlDoc = CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = False 
			
			Set fso=createobject("Scripting.FileSystemObject")
			
			'Set qfile=fso.OpenTextFile("C:\Automation\APPLICATIONS\FFWF\XML\AVS_Order_Fulfillment.xml",1)
			Set qfile=fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&inputdata&".xml",1)
			sResponsexml=qfile.ReadAll
			print "Response xml is"&sResponsexml
			
			Set qfile=nothing
			Set fso=nothing
			
			Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Highlight
			wait (2)
			Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Click
				Set ws=CreateObject("wscript.shell")
		      	ws.SendKeys ("^a")	
			    ws.SendKeys "{DELETE}"
			    'ws.SendKeys sResponsexml
			    
				wait (2)
		
				Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Set sResponsexml
	 Actual_Res = "XML is loaded successfully from the XML file"
	Reporter.reportevent micPass, StepName, Actual_Res
	     		Set ws=Nothing 
			End if

wait(2)

'oBook.Save
'oBook.Close
'oExcel.Quit

End if

' Handling Error
	methodName = "Buslistnerorder_xml" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function Buslistnerorder_xml() ------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: Generic_CLICK(inputData)
'# Purpose			: To click an object in the application
'#					  objectDesc	-> Gets the Description of the Object to be clicked
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Generic_CLICK(inputData)
	On Error Resume Next
	Dim methodName,ObjDesc_Array, Object_Arr, Object_Name, objName, rc, exec_Stmt ,exec_stmnt, Flag,val,val1,Object_Desc
	methodName = "Generic_CLICK" : Generic_CLICK = 0
	Flag = False
	
	ObjDesc_Array = split(inputData,"=>")
	'msgbox ObjDesc_Array(0)
	ObjSplit_Array = split(ObjDesc_Array(1)," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
	Select Case (ObjSplit_Count)
		Case "0"
			Object_Desc = ObjDesc_Array(1)
			
		Case "1"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
			'msgbox Object_Desc
		Case "2"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
			
		Case "3"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
			
		Case "4"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
			
		Case Else
			Step_Description = "Object property"
			Actual_Res = "Object property exceeded more than five"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			Generic_CLICK = -1
	End Select

	'To generate Test Step Description and Expected Result
	Step_Description = "Click on object -> " & objName & "  in page generic page"
	Exp_Result = "Clicked on object ->"& objName & " in generic page"

	If Exec_Flag = "Y" Then
		exec_stmnt = "rc = Browser(""name:=TMS Bus Listener.*"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Exist(""10"")"
		'msgbox exec_stmnt
		Execute(exec_stmnt)
		'msgbox rc
		If rc = "True" Then
			exec_Stmt = "Browser(""name:=TMS Bus Listener.*"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Click"
			Execute(exec_Stmt)
			Flag = True
			Actual_Res = "Click on object -> " & objName & " in generic page is successful"
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else Generic_CLICK = -1 End If
If Flag = False Then
	Actual_Res = "Click on object -> " & objName & " in generic page failed"
	Reporter.ReportEvent micFail, StepName, Actual_Res
	Generic_CLICK = -1
End If
			
End If
	' Handling Error
	methodName = "Generic_CLICK" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function Generic_CLICK() ----------------------------------------------------------------------


'# Function Name	: Generic_Input1(inputData,inputData1)
'# Purpose			: To click an object in the application with the object details provided
'#					  objectDesc	-> Gets the Description of the Object to be clicked
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Generic_Input1(inputData,inputData1)
	On Error Resume Next
	Dim methodName,ObjDesc_Array, Object_Arr, Object_Name, objName, rc, exec_Stmt ,exec_stmnt, Flag,val,val1,Object_Desc
	methodName = "Generic_Input1" : Generic_Input1 = 0
	Flag = False

	ObjDesc_Array = split(inputData,"=>")
	'msgbox ObjDesc_Array(0)
	ObjSplit_Array = split(ObjDesc_Array(1)," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
	Select Case (ObjSplit_Count)
		Case "0"
			Object_Desc = ObjDesc_Array(1)
			
		Case "1"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
			'msgbox Object_Desc
		Case "2"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
			
		Case "3"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
			
		Case "4"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
			
		Case Else
			Step_Description = "Object property"
			Actual_Res = "Object property exceeded more than five"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			Generic_Input1 = -1
	End Select

	If Trim(ObjDesc_Array(0)) = "WebList" OR Trim(ObjDesc_Array(0)) = "WebRadioGroup" Then
			In_Operation = "Select "
		Else
			In_Operation = "Set "
	End If

	'Collect Input
	data = Trim(inputData1)
	'msgbox data
	'To generate Test Step Description and Expected Result
	Step_Description = "Click on object -> " & objName & "  in page generic page"
	Exp_Result = "Clicked on object ->"& objName & " in generic page"

	If Exec_Flag = "Y" Then
		exec_stmnt = "rc = Browser(""name:=TMS Bus Listener.*"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Exist(""10"")"
		'msgbox exec_stmnt
		Execute(exec_stmnt)
		'msgbox rc
		If rc = "True" Then
			Browser("name:=TMS Bus Listener.*").Page("title:=.*").highlight
			exec_Stmt = "Browser(""name:=TMS Bus Listener.*"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """)." & In_Operation & """" & data & """"
			'msgbox exec_Stmt
			Execute(exec_Stmt)
			Flag = True
			Actual_Res = "input of-> " & objName & " in generic page is successful"
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else Generic_Input = -1 End If
If Flag = False Then
	Actual_Res = "input of -> " & objName & " in generic page failed"
	Reporter.ReportEvent micFail, StepName, Actual_Res
	Generic_Input1 = -1
End If
			
End If
	' Handling Error
	methodName = "Generic_Input1" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function Generic_Generic_Input1() ----------------------------------------------------------------------


'# Function Name	: Generic_Input(inputData,inputData1)
'# Purpose			: To give an input to the object in the application
'#					  objectDesc	-> Gets the Description of the Object to be clicked
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Generic_Input(inputData,inputData1)
	On Error Resume Next
	Dim methodName,ObjDesc_Array, Object_Arr, Object_Name, objName, rc, exec_Stmt ,exec_stmnt, Flag,val,val1,Object_Desc
	methodName = "Generic_Input" : Generic_Input = 0
	Flag = False

	ObjDesc_Array = split(inputData,"=>")
	'msgbox ObjDesc_Array(0)
	ObjSplit_Array = split(ObjDesc_Array(1)," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
	Select Case (ObjSplit_Count)
		Case "0"
			Object_Desc = ObjDesc_Array(1)
			
		Case "1"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
			'msgbox Object_Desc
		Case "2"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
			
		Case "3"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
			
		Case "4"
			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
			
		Case Else
			Step_Description = "Object property"
			Actual_Res = "Object property exceeded more than five"
			Reporter.ReportEvent micFail, StepName, Actual_Res
			Generic_Input = -1
	End Select

	If Trim(ObjDesc_Array(0)) = "WebList" OR Trim(ObjDesc_Array(0)) = "WebRadioGroup" Then
	
			In_Operation = "Select "
		Else
			In_Operation = "Set "
	End If

	'Collect Input
	data = Trim(inputData1)
	'msgbox data
	'To generate Test Step Description and Expected Result
	Step_Description = "Click on object -> " & objName & "  in page generic page"
	Exp_Result = "Clicked on object ->"& objName & " in generic page"

	If Exec_Flag = "Y" Then
		exec_stmnt = "rc = Browser(""name:=TMS Bus Listener.*"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Exist(""10"")"
		'msgbox exec_stmnt
		Execute(exec_stmnt)
		'msgbox rc
		If rc = "True" Then
			exec_Stmt = "Browser(""name:=TMS Bus Listener.*"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """)." & In_Operation & """" & data & """"
			'msgbox exec_Stmt
			Execute(exec_Stmt)
			Flag = True
			Actual_Res = "Click on object -> " & objName & " in generic page is successful"
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else Generic_Input = -1 End If
If Flag = False Then
	Actual_Res = "Click on object -> " & objName & " in generic page failed"
	Reporter.ReportEvent micFail, StepName, Actual_Res
	Generic_Input = -1
End If
			
End If
	' Handling Error
	methodName = "Generic_Input" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function Generic_Generic_Input() -----------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: XMLValue_Collect()
'# Purpose			: To collect XML value
'# Parameter                                :inputData- > Give the index of row 
'#                                                      :inputData1-> Give the tag value
'#                                                       :inputData2 -> Give the scenario against which value needs to be written
'#                                                        :inputData3-> Give the index of the tag
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function XMLValue_Collect()
On Error Resume Next

	Dim methodName,rc,val,Validate_Bus_Listener,obj,xml_cnt,obj1,obj_xml,i,j,N,k,m,strText,xmlDoc,XMLDataFile,nodelist,tag_val,tag,REQUEST_ID,mysheet,Row,Col,myxl,qfile,fso

	methodName = "XMLValue_Collect" : XMLValue_Collect = 0
	i = inputData

		
If Exec_Flag = "Y" Then
'Set val = Browser("creationtime:=2").Page("title:=.*").WebXML("html tag:=BODY").GetData
Browser("name:=TMS Bus Listener.*").Page("title:=.*").WebEdit("name:=tbMessages").Highlight
val = Browser("name:=TMS Bus Listener.*").Page("title:=.*").WebEdit("name:=tbMessages").GetROProperty("value")
wait(2)
'msgbox val
val1 = split(val,"is:")
'msgbox val1(1)
val2 = Replace(val1(1),"_______________________________________________________","")
'msgbox val2

Step_Description = "The XML value should be collected successfully " 
Exp_Result = " The required XML value is " & val2 




'xml_val = val.ToString()
'xml_val = val
'msgbox xml_val
Set fso=createobject("Scripting.FileSystemObject")
'If qfile.FileExists("C:\Automation\APPLICATIONS\FFWF\XML\Validate_Bus_Listener.txt") Then
'Set qfile = fso.GetFile("C:\Automation\APPLICATIONS\FFWF\XML\Validate_Bus_Listener.txt")
'Else 
'qfile.Delete()
'End if

Set qfile=fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.txt",2,true)

qfile.Write val2

'#############################

'Set objFSO = CreateObject("Scripting.FileSystemObject")
Set qfile = fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.txt", 1)


Do Until qfile.AtEndOfStream
    strLine = qfile.Readline
    strLine = Trim(strLine)
    If Len(strLine) > 0 Then
        strNewContents = strNewContents & strLine & vbCrLf
    End If
Loop

Set qfile = fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.txt", 2)
qfile.Write strNewContents
qfile.Close

'#############################


fso.CopyFile "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.txt","C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.xml"

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False 
'path of XML file

XMLDataFile="C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.xml"

'Load the XML File
 xmlDoc.Load(XMLDataFile)
'get the tagname to verify
Set nodelist = xmlDoc.getElementsByTagName("RequestType")
strText = nodelist.item(0).text

'msgbox strText

If strText<>"" Then
	rc = 0
	Step_Description = "Copy the contents of the text file to xml file"
	Exp_Result = "Xml value should be collected successfully"
	Actual_Res = "XML value is collected successfully" 
	Reporter.ReportEvent micPass, StepName, Actual_Res
	
	Else
	
	Step_Description = "Copy the contents of the text file to xml file"
	Exp_Result = "Xml value should be collected successfully"
	Actual_Res = "XML value collected is not collected successfully for the requested node"
	Reporter.ReportEvent micFail, StepName, Actual_Res
	XMLValue_Collect =-1
End If

'qfile.writeline val2
'
'msgbox val2
'
'qfile.Close

Set qfile=nothing
Set fso=nothing

End if
	' Handling Error
	methodName = "XMLValue_Collect" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function XMLValue_Collect() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: ValidateXMLTags()
'# Purpose			: To collect XML value
'# Parameter                                			:inputData- > Give the index of row 
'#                                                      :inputData1-> Give the value of the xml file path
'#                                                       :inputData2 -> Give the xml tag for which the value has to be updated
'#                                                        :inputData3-> Give the tag value which has to be updated
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________

Public Function ValidateXMLTags(inputData1)
On Error Resume Next

Dim methodName,rc,val,xml_val,obj,xml_cnt,obj1,obj_xml,i,j,N,k,m,strText,xmlDoc,XMLDataFile,nodelist,session_val,rxid,RxSessionId,mysheet,Row,Col,myxl,item

	methodName = "ValidateXMLTags" : ValidateXMLTags = 0
	
If Exec_Flag = "Y" Then

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False 
'path of XML file

XMLDataFile="C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.xml"

'Load the XML File
 xmlDoc.Load(XMLDataFile)
'get the tagname to verify
Set nodelist = xmlDoc.getElementsByTagName(inputData1)
'msgbox nodelist.length
val =  nodelist.length
'msgbox val


strText = nodelist.item(0).xml

'msgbox strText

'For m = 0 to val-1
'
'       strText = nodelist.item(m).xml
'       'strText = nodelist(i).nodevalue
'       msgbox strText
'       tag_val = split(strText,">")
'       tag = split(tag_val(1),"<")
'       REQUEST_ID =  tag(0)
'       
'		msgbox REQUEST_ID
'
'       If abs(m) = abs(inputData3) Then
'	    Exit For
'       End If
'
''       If instr(trim(ucase(strText)),"GPONFTTP")<>0 Then
''       	      msgbox "success" 
'       'End If
'Next

If strText<>"" Then
	rc = 0
	Step_Description = "To Collect XML Value"
	Exp_Result = "Xml value should be collected successfully"
	Actual_Res = "XML value collected is-> " & strText 
	Reporter.ReportEvent micPass, StepName, Actual_Res
	
	Else
	
	Step_Description = "To Collect XML Value"
	Exp_Result = "Xml value should be collected successfully"
	Actual_Res = "XML value collected is not collected successfully for-> " & inputData1
	Reporter.ReportEvent micFail, StepName, Actual_Res
	ValidateXMLTags =-1
End If


'Set myxl = createobject("excel.application")
'
'myxl.Workbooks.Open "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\DATA\FFWF_Data.xlsx" 
''myxl.Application.Visible = true
' 
''this is the name of  Sheet  in Excel file "qtp.xls"   where data needs to be entered 
'set mysheet = myxl.ActiveWorkbook.Worksheets("Test_Data")
' 
'Col=mysheet.UsedRange.columns.count
''msgbox Col
'
'For k = 1 To Col
'	If Trim(mysheet.cells(1,k).value) = Trim(inputData2)  Then
'	msgbox strText
'	mysheet.cells(i,k).value = strText
'	
'	msgbox mysheet.cells(i,k).value
'	Exit For
'End If
'Next
'
''Save the Workbook
'myxl.ActiveWorkbook.Save
' 
''Close the Workbook
'myxl.ActiveWorkbook.Close
' 
''Close Excel
''myxl.Application.Quit
' 
'Set mysheet =nothing
'Set myxl = nothing

End if
'Browser("micclass:=Browser","name:=TMS Bus Listener.*").Close

' Handling Error
	methodName = "ValidateXMLTags" : rc = ErrorHandler(methodName)
End Function

'___________________________________________________________________________________________________________________________
'# Function Name	: ReadXMLFileAndReplaceXMLTags()
'# Purpose			: To collect XML value
'# Parameter                                			:inputData- > Give the index of row 
'#                                                      :inputData1-> Give the value of the xml file path
'#                                                       :inputData2 -> Give the xml tag for which the value has to be updated
'#                                                        :inputData3-> Give the tag value which has to be updated
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________

Public Function ReadXMLFileAndReplaceXMLTags(Query_ID,FileLocation,Tagtobereplaced,TagValue)
On Error Resume Next

Sheetname = "Query"
FilePath="C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx"
Datatable.AddSheet Sheetname
Datatable.ImportSheet FilePath, Sheetname, Sheetname
verify_First_Data_from_Datatable= Datatable.Value("Query",Sheetname)

Print "Method ReadXMLFileAndReplaceXMLTags, verify first data" & verify_First_Data_from_Datatable

getRowCount=Datatable.GetSheet(Sheetname).GetRowCount
getParamCount=Datatable.GetSheet(Sheetname).GetParameterCount
Print "getParamCount ="&getParamCount

Print "Total Rows ="&getRowCount
Datatable.GetSheet(Sheetname).SetCurrentRow(1)
For Iterator = 1 To getRowCount Step 1

	'Print "Iterator ="&Iterator
	Query_ID="Q"&Iterator
	'QID=Datatable.GetSheet(Sheetname).GetParameter("QID").Value
	Print "QID ="&QID
	Print "Query_ID=" &Query_ID
	 
	 If QID=Query_ID Then	
	 
	FileLocation=Datatable.GetSheet(Sheetname).GetParameter("FileLocation").Value
	Print "FileLocation ="&FileLocation
	Tagtobereplaced=Datatable.GetSheet(Sheetname).GetParameter("Tagtobereplaced").Value
	Print "Tagtobereplaced ="&Tagtobereplaced
	TagValue=Datatable.GetSheet(Sheetname).GetParameter("TagValue").Value
	Print "TagValue ="&TagValue
	Else

End If
	
	
	Dim methodName,rc,strText,xmlDoc,XMLDataFile,nodelist

	methodName = "ReadXMLFileAndReplaceXMLTags" : ReadXMLFileAndReplaceXMLTags = 0
	
	If Exec_Flag = "Y" Then
	
		Set fileSystemObj = createobject("Scripting.FileSystemObject")
		MyFile = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx"
	
			If fileSystemObj.FileExists(MyFile) then
				rc = "True"
			Else
				rc = "False"
			End If
	
			If rc = "True" Then		
			
				'QID=Queryid+1
				'msgbox QID
				'Set myxl = createobject("excel.application")
				'
				'myxl.Workbooks.Open "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx" 
				'
				'set mysheet = myxl.ActiveWorkbook.Worksheets("Query")
				'
				'Row=mysheet.UsedRange.rows.count
				'
				'FileLocation = mysheet.cells(QID,1).value
				'msgbox QID
				'FileLocation = mysheet.cells(QID,5).value
				'msgbox FileLocation
				'Tagtobereplaced = mysheet.cells(QID,6).value
				'msgbox Tagtobereplaced
				'TagValue = mysheet.cells(QID,4).value
				'msgbox TagValue
				'
				'myxl.Workbooks.Close
				'myxl.Quit
				'
				Set xmlDoc = CreateObject("Microsoft.XMLDOM")
				xmlDoc.Async = False 
				
				'path of XML file
				
				XMLDataFile=FileLocation
				
				'Load the XML File
				xmlDoc.Load(XMLDataFile)
				'get the tagname to verify
				Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)
				
				nodelist.text = TagValue
				
				xmlDoc.Save(XMLDataFile)
				
				Set xmlDoc = nothing
			
			Else
				Reporter.ReportEvent micFail, "ReadXMLFileAndReplaceXMLTags - Function failed"," rc = false - Please check"
			End if
			Else
				Reporter.ReportEvent micFail, "ReadXMLFileAndReplaceXMLTags - Function failed"," rc = false - Please check"
			End if

Datatable.SetNextRow
Step_Description = "XML tag values are updated for the Query id -> " & Query_ID
Exp_Result = Tagtobereplaced & " -> is updated with the tag value as " & TagValue &" for the Query id " & Query_ID
Actual_Res = "Tagvalue is updated successfully"
Reporter.reportevent micPass,StepName, Actual_Res


Next

'ExitTest
'	 'Handling Error
	methodName = "ReadXMLFileAndReplaceXMLTags" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function ReadXMLFileAndReplaceXMLTags() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: FFWF_DB_Connect()
'# Purpose			: To collect XML value
'# Parameter                                			:inputData- > To connect to the DB with the Connection STring details provided
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________

Public Function FFWF_DB_Connect(TransId)
On Error Resume Next

'TransId= "10293845tg123"
Exec_Flag = "Y"
Set fileSystemObj = createobject("Scripting.FileSystemObject")
MyFile = fileSystemObj.FileExists("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Details_All.xlsx")

If MyFile then
 fileSystemObj.DeleteFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Details_All.xlsx")
		    
		    
		Else
		Set ExcelObj=CreateObject("Excel.Application")
		MySourceFile = fileSystemObj.FileExists("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx")
					If  MySourceFile Then          
						ExcelObj.Workbooks.Open("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx") 
						
					    ExcelObj.ActiveWorkbook.SaveAs("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Details_All.xlsx")

					 Else                          
					   
					 End If
					 ExcelObj.Workbooks.Close
							ExcelObj.Quit
				End If
									      
set fileSystemObj = nothing
Set ExcelObj = nothing
		
Sheetname = "Query"

FilePath="C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Details_All.xlsx"
Datatable.AddSheet Sheetname
Datatable.ImportSheet FilePath, Sheetname, Sheetname

'verify_First_Data_from_Datatable= Datatable.Value("Query",Sheetname)
'Print verify_First_Data_from_Datatable

getRowCount=Datatable.GetSheet(Sheetname).GetRowCount
getParamCount=Datatable.GetSheet(Sheetname).GetParameterCount
Print "getParamCount ="&getParamCount

Print "Total Rows ="&getRowCount
Datatable.GetSheet(Sheetname).SetCurrentRow(1)
For Iterator = 1 To getRowCount Step 1
	'Print "Iterator ="&Iterator
	'Query_ID="Q"&Iterator

	Query_ID= Datatable.GetSheet(Sheetname).GetParameter("QID").Value
	Print "Query_ID ="&Query_ID
	
	getRowVal=Datatable.GetSheet(Sheetname).GetParameter("Query").Value
	Print "Query_Val ="&getRowVal
	
'	Print Query_ID
	
If (Query_ID = "Resp_RequestId")  or (Query_ID ="Redesign_Resp_RequestId" ) Then
 		 Print "TransId = "  &TransId
		 getRowVal1 = getRowVal & "'" & TransId & "')));"
		 Print "Query_Val ="&getRowVal1

ElseIf (Query_ID = "Resp_SON") or (Query_ID = "Redesign_Resp_SON") Then
		 Print "TransId = "  &TransId
		 getRowVal1 = getRowVal & "'" & TransId & "');"
		 Print "Query_Val ="&getRowVal1

ElseIf (Query_ID ="Ship_RequestId") or (Query_ID = "Redesign_Ship_RequestId") Then
		 Print "TransId = "  &TransId
		 getRowVal1 = getRowVal & "'" & TransId & "')));"
		 Print "Query_Val ="&getRowVal1

ElseIf (Query_ID ="Ship_SON") or (Query_ID = "Redesign_Ship_SON")  Then
		 Print "TransId = "  &TransId
		 getRowVal1 = getRowVal & "'" & TransId & "');"
		 Print "Query_Val ="&getRowVal1

ElseIf (Query_ID ="Ship_ItemId") or (Query_ID = "Redesign_Ship_ItemId") Then
		 Print "TransId = "  &TransId
		 getRowVal1 = getRowVal & "'" & TransId & "'"
		 Print "Query_Val ="&getRowVal1

'ElseIf (Query_ID ="DOM_RequestId") Then
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "')));"
'		 Print "Query_Val ="&getRowVal1
'
'ElseIf (Query_ID ="DOM_SON") Then
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "');"
'		 Print "Query_Val ="&getRowVal1
'
'ElseIf (Query_ID ="DOM_CAI") Then
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "'"
'		 Print "Query_Val ="&getRowVal1	

End if
	
'	 Select Case Query_ID
' Case "Q1"
' 		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "')));"
'		 Print "Query_Val ="&getRowVal1
' 
' Case "Q2"
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "');"
'		 Print "Query_Val ="&getRowVal1
' 
' Case "Q3"
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "')));"
'		 Print "Query_Val ="&getRowVal1
' 
' Case "Q4"
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "');"
'		 Print "Query_Val ="&getRowVal1
'		 
'Case "Q5"
'		 Print "TransId = "  &TransId
'		 getRowVal1 = getRowVal & "'" & TransId & "'"
'		 Print "Query_Val ="&getRowVal1
'
''Case "Q6"
''		 Print "TransId = "  &TransId
''		 getRowVal1 = getRowVal & "'" & TransId & ");"
''		 Print "Query_Val ="&getRowVal1
''		 
''Case "Q7"
''		 Print "TransId = "  &TransId
''		 getRowVal1 = getRowVal & "'" & TransId & ""
''		 Print "Query_Val ="&getRowVal1
'		 
' End Select
 

Dim myxl,mysheet,Row,Exec_Flag,Query_ID,DBQuery,Fieldname
Dim DbConn,rc,FF_Requestid_VAL,FF_ORDER_VAL,FF_Requestid_VAL2,FF_ORDER_VAL2,ITEM_ID_VAL2
Dim ConnectionString,CN,fso,port,SID,unamepasswd,host,ServiceName,Test1,E2E

methodName = "FFWF_DB_Connect" : FFWF_DB_Connect = 0

Step_Description = "Load DB Details from the Query Sheet" 
Exp_Result = " Queries are loaded successfully from the Query Sheet"
Print "Checking the db connection now:::"
If Exec_Flag = "Y" Then

Set fileSystemObj = createobject("Scripting.FileSystemObject")
MyFile = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Details_All.xlsx"

If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		Else
		    rc = "False"
		End If
		
If rc = "True" Then			
Set DbConn=CreateObject("ADODB.Connection")
Set rs=Createobject("ADODB.Recordset")
 
'ConnectionString="Driver={Oracle in OraClient 12home1_32bit};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=ffwfst1db.dev.qintra.com)(PORT=1539))(CONNECT_DATA=(TNS Service Name=ffwfst1db.dev.qintra.com:1539/ffwfst1)));Data Source=FFWF-ST1;User ID=ffwf_app;Password=ffwfst1_suomt102;"
'ConnectionString="Driver={Oracle in OraClient 12home1_32bit};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=ffwfe2edb.dev.qintra.com)(PORT=1542))(CONNECT_DATA=(TNS Service Name=ffwfe2edb.dev.qintra.com:1542/ffwfe2e)));Data Source=FFWFE2E DB;User ID=ffwf_app;Password=ffwfe2e_suomt102;"
ConnectionString="Driver={Oracle in OraClient 12home1_32bit};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=ffwfst2db.dev.qintra.com)(PORT=1541))(CONNECT_DATA=(TNS Service Name=ffwfst2db.dev.qintra.com:1541/ffwfst2)));Data Source=FFWF-ST2;User ID=ffwf_app;Password=ffwfst2_suomt102;"
'ConnectionString = E2E
'msgbox ConnectionString
DbConn.Open ConnectionString
DBQuery= getRowVal1
Print "DB query is=" & DBQuery

's rc.Open DBQuery,DbConn
rs.Open DBQuery, DbConn

vale= rs.Fields.Item(0)
'Print "Value iss:::" & vale
'rc.Open getRowVal1,DbConn

Print "Connection Status ="&DbConn.State

If DbConn.State=0 Then
      
		Actual_Res = "FFWF DB Connection Status Not Established" 
		Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
		FFWF_DB_Connect=-1
Else
		Actual_Res = "FFWF DB Connection Status Established Successfully" 
		Reporter.ReportEvent  micPass, StepName, Actual_Res
End If

'  If err.number <> 0 Then
'										Call captureScreen
'										Actual_Res = "Not Able to query the required DB" 
'										Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'										FFWF_DB_Connect=-1
'                                else
'                                		Actual_Res = "Able to query the required DB successfully" 
'										Reporter.ReportEvent micPass, StepName, Actual_Res
'
'                End If   



'If rc.EOF <> True Then
If rs.EOF <> True Then
                If (Query_ID = "Resp_RequestId")  or (Query_ID ="Redesign_Resp_RequestId" ) Then
                               ' FF_Requestid_VAL = rs("FULFILLMENT_REQUEST_ID").Value
                                'FF_Requestid_VAL= rs.Fields("FULFILLMENT_REQUEST_ID").Value
                                FF_Requestid_VAL=  rs.Fields.Item("FULFILLMENT_REQUEST_ID")
                                Print "The FULFILLMENT_REQUEST_ID for the FFWF Response is: " & FF_Requestid_VAL
                                                             
                                 inputdata=FF_Requestid_VAL
                                  Call  Copy_xmlvaluetoexcel(2,4,inputdata)
                                  Call  Copy_xmlvaluetoexcel(7,4,inputdata)   
                                
                                  
                End If

                
                If (Query_ID = "Resp_SON") or (Query_ID = "Redesign_Resp_SON") Then
                                'FF_ORDER_VAL = rc("ORDER_NUMBER").Value
                                FF_ORDER_VAL=  rs.Fields.Item("ORDER_NUMBER")
                                Print  "The ORDER_NUMBER for the FFWF Response is: "  &FF_ORDER_VAL
                                
                                inputdata=FF_ORDER_VAL
                                Call  Copy_xmlvaluetoexcel(3,4,inputdata)
                                Call  Copy_xmlvaluetoexcel(8,4,inputdata)

                End If    
                
                If (Query_ID ="Ship_RequestId") or (Query_ID = "Redesign_Ship_RequestId") Then
                                'FF_Requestid_VAL2 = rc("FULFILLMENT_REQUEST_ID").Value
                                FF_Requestid_VAL2=  rs.Fields.Item("FULFILLMENT_REQUEST_ID")
                                Print  "The FULFILLMENT_REQUEST_ID for the FFWF Shipment Response is: "  &FF_Requestid_VAL2

                                inputdata=FF_Requestid_VAL2
                                Call  Copy_xmlvaluetoexcel(4,4,inputdata)
                                Call  Copy_xmlvaluetoexcel(9,4,inputdata)

                End If
                
                If (Query_ID ="Ship_SON") or (Query_ID = "Redesign_Ship_SON") Then
                                'FF_ORDER_VAL2 = rc("ORDER_NUMBER").Value
                                FF_ORDER_VAL2=  rs.Fields.Item("ORDER_NUMBER")
                                Print  "The ORDER_NUMBER for the FFWF Shipment Response is: "  &FF_ORDER_VAL2
                               
                                inputdata=FF_ORDER_VAL2
                                Call  Copy_xmlvaluetoexcel(5,4,inputdata)
                                Call  Copy_xmlvaluetoexcel(10,4,inputdata)

                End If
                
                If (Query_ID ="Ship_ItemId") or (Query_ID = "Redesign_Ship_ItemId") Then
                                'ITEM_ID_VAL2 = rc("SHIPPABLE_INSTANCE_ID").Value
                                ITEM_ID_VAL2=  rs.Fields.Item("SHIPPABLE_INSTANCE_ID")
                                Print  "The SHIPPABLE_INSTANCE_ID(ITEM ID) for the FFWF Shipment Response is: "&ITEM_ID_VAL2
                              
                                inputdata=ITEM_ID_VAL2
                                Call  Copy_xmlvaluetoexcel(6,4,inputdata)
                                Call  Copy_xmlvaluetoexcel(11,4,inputdata)
                                 
                End If
                
'                If Query_ID="Q6" Then
'                                FF_Requestid_VAL = rc("FULFILLMENT_REQUEST_ID").Value
'                                Print "The FULFILLMENT_REQUEST_ID for the FFWF Response is: " & FF_Requestid_VAL
'                                                             
'                                inputdata=FF_Requestid_VAL
'                                  Call  Copy_xmlvaluetoexcel(2,4,inputdata)
'                                     
'                                
'                                  
'                End If
'
'                
'                If Query_ID="Q7" Then
'                                FF_ORDER_VAL = rc("ORDER_NUMBER").Value
'                                Print  "The ORDER_NUMBER for the FFWF Response is: "  &FF_ORDER_VAL
'                                
'                                inputdata=FF_ORDER_VAL
'                                Call  Copy_xmlvaluetoexcel(3,4,inputdata)
'
'                End If    
                
'                                                     Step_Description = "DB output should be updated in the DB Details excel successfully"
'                                    Exp_Result = "DB output should be updated in the DB Details excel successfully"
'                                    Actual_Res = "The DB output is  updated successfully"
'                                    Reporter.reportevent micPass,StepName, Actual_Res
                
 
 Else
'                                     Step_Description = "DB output should be updated in the DB Details excel successfully"
'                                    Exp_Result = "DB output should be updated in the DB Details excel successfully"
'                                    Actual_Res = "The DB output is not updated successfully"
'                                    Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'           FFWF_DB_Connect=-1
                  
End if
Set DbConn = Nothing
Set rc = Nothing
                                    
End if
'                                     Step_Description = "DB output should be updated in the DB Details excel successfully"
'                                    Exp_Result = "DB output should be updated in the DB Details excel successfully"
'                                    Actual_Res = "The DB output is not updated successfully"
'                                    Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'           FFWF_DB_Connect=-1
End if

	Datatable.SetNextRow
	
Next


ExitTest
	'Handling Error
	methodName = "FFWF_DB_Connect" : rc = ErrorHandler(methodName)
End Function

' --------------------------- End of Function FFWF_DB_Connect() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: CLOSE_()
'# Purpose			: To close a specific browser.
'# Parameters 		: pageDesc		-> Gets the description of the Browser.
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function CLOSE_(pageDesc)
	On Error Resume Next
	Dim methodName, Page_Name, pageName, rc
	methodName = "CLOSE_" : CLOSE_ = 0
	
	Page_Name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
	'To generate Test Step Description and Expected Result
	Step_Description = "Close the Browser with page ->  " & pageName
	Exp_Result =  "Should Close the Browser with page ->  " & pageName
	
	If Exec_Flag = "Y" Then
		rc = EXIST_(pageDesc, "")
		If rc = 0 Then
			'Browser("name:=TMS Bus Listener").Close
			Browser("title:=" & pageName).Close
			Actual_Res = "Closing the browser with page -> " & pageName
			Reporter.ReportEvent micPass, StepName, Actual_Res
		Else
			Call captureScreen
			Actual_Res = "Browser with page -> " & pageName & " NOT found."
			Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
			CLOSE_ = -1
		End If
	End If	
	' handle error
	methodName = "CLOSE_" : rc = ErrorHandler(methodName)
End Function
' --------------------------- End of Function CLOSE_() ---------------------------------------------------------------------


'___________________________________________________________________________________________________________________________
'# Function Name	: Copy_xmlvaluetoexcel()
'# Purpose			: To collect XML value
'# Parameter                                			:row- > row in which the copied value to be pasted 
'#                                                      :col-> col in which the copied value to be pasted 
'#                                                       :inputData -> to give the value which has to be pasted in the provided row and col
'#                                                       
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________

Public Function Copy_xmlvaluetoexcel(row,col,inputdata)
Flag = True
                                                                                
                                                      'msgbox row
                                                      'msgbox col
                                                      'msgbox inputdata
                                                      'On Error Resume Next
                                                      'i=1
                                                      'rc=trim(inputData)
                                                      'r = split(rc,",")
                                                      'row = r(0) 
                                                      'column =  r(1)-1
                                                 
                                                      	
                                                      
                                                      			  Set obj=Description.Create()
'                                                                 obj("Class Name").Value="JavaToolbar"
'                                                                 obj("attached text").Value="Product Home"
'                                                                 Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
'                                                                     For i = 0 To obj1.Count-1
'                                                                     obj1(i).Highlight
'                                                                     obj1(i).Press "Copy RX# to clipboard"
                                                                      Dim objCB,str_1,xl
                                                                      'Set objCB= CreateObject("Mercury.Clipboard")
                                                                      'str_1 = objCB.GetText
                                                                      'str_1 = FF_Requestid_VAL
                                                                       set xl=CreateObject("Excel.Application") 
                                                                       set workbook= xl.Workbooks.Open("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx")  
                                                                       set sheet=workbook.Sheets("Query")
                                                                       sheet.Cells(row,col).value=inputdata
                                                                       'sheet.Cells(row,column).interior.colorindex=6
                                                                        xl.ActiveWorkbook.Save
                                                                        'xl.ActiveWorkbook.Close
                                                                        'set xl.Application=Nothing
                                                                        'Set xl=Nothing
                                                                         xl.Workbooks.Close
																		 xl.Quit
									
'                                                                                Next
                                               
           ' Handling Error
	methodName = "Copy_xmlvaluetoexcel" : rc = ErrorHandler(methodName)                                                     
End Function
' --------------------------- End of Function Copy_xmlvaluetoexcel() --------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: ToUpdateDateinXML()
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		:inputData- > BAN to select the row in the list
'#					:inputData1- > status value as per the input
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function ToUpdateDateinXML(inputData)
On Error Resume Next

Dim FileLocation,sdate,xmlDoc,XMLDataFile,methodName,rc
methodName = "ToUpdateDateinXML" : ToUpdateDateinXML = 0

	Step_Description = "To update date in the XML response file -> "
	Exp_Result = "Today's Date has to be updated successfully in the XML response file"

MyFile = ("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&inputData&".xml")

		If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		Else
		    rc = "False"
		End If

		If rc = "True" Then
		
FileLocation = ("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&inputData&".xml")
print "Filelocations is:" &FileLocation
	
sdate= Year(date)&"-"&month(date)&"-"&day(date)
print "Today Date is:" &sdate

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False 

XMLDataFile=FileLocation

xmlDoc.Load(XMLDataFile)


Select Case inputData
Case "AVS_Order_Fulfillment"
Set strDateNode = xmlDoc.SelectSingleNode("ENJEventMessageRequest/bim:SendTimeStamp")

Case "DvarOm_Modem_Cancel_Return"
Set strDateNode = xmlDoc.SelectSingleNode("ReturnSTBShipmentRequest/RequestHeader/SendTimeStamp")

Case "DvarOm_Modem_Cancel_Return_Response"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")

Case "DvarOm_Modem_Cancel_Return_Response_success"
Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentReturnResponse/ReturnResponseHeader/qb:SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentReturnResponse/ReturnAuthorization/ReturnAuthorizationExpirationDate")

Case "DVAROM_Order_Fulfillment"
Set strDateNode = xmlDoc.SelectSingleNode("ENJEventMessageRequest/bim:SendTimeStamp")

Case "ENJ_Order_Fulfillment"
Set strDateNode = xmlDoc.SelectSingleNode("ENJEventMessageRequest/bim:SendTimeStamp")

Case "FF_MASTER_PROCESS_REDESIGN_FF_Response"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/SendTimeStamp")

Case "FF_MASTER_PROCESS_REDESIGN_FF_Shipment"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ShippingDate")

Case "PURETV_VENDOR_DELIVERY_PROCESS_REQUEST"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentRequest/ns2:ShipmentRequestHeader/SendTimeStamp")

Case "PURETV_VENDOR_DELIVERY_PROCESS_RESPONSE"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ShippingDate")

Case "PURETV_VENDOR_INITIATED_RETURNS_REQUEST"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnRequestHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")

Case "PURETV_VENDOR_INITIATED_RETURNS_RESPONSE"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")

Case "Resp"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/ns2:OrderReceivedDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ExpectedShippingDate")

Case "Ship"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/ns2:OrderReceivedDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:EstimatedDeliveryDate")

Case "VENDOR_DELIVERY_PROCESS_ENS_REQUEST"
Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentShipmentRequest/ShipmentRequestHeader/qb:SendTimeStamp")

Case "VENDOR_DELIVERY_PROCESS_ENS_RESPONSE"
Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/qb:SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/ns:OrderReceivedDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:EstimatedDeliveryDate")

Case "VENDOR_DELIVERY_PROCESS_IOM_REQUEST"
Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentShipmentRequest/ShipmentRequestHeader/qb:SendTimeStamp")

Case "VENDOR_DELIVERY_PROCESS_IOM_RESPONSE"
Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/qb:SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/ns:OrderReceivedDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponse/qb:ShippingDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:EstimatedDeliveryDate")

Case "VENDOR_INITIATED_RETURNS_ENS_REQUEST"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnRequestHeader/SendTimeStamp")

Case "VENDOR_INITIATED_RETURNS_ENS_RESPONSE"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnDate")
'strDateNode.text = sdate

Case "VENDOR_INITIATED_RETURNS_IOM_REQUEST"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnRequestHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")

Case "VENDOR_INITIATED_RETURNS_IOM_RESPONSE"
Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")

End Select
  
  
'get the tagname to verify
'Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)

		
strDateNode.text = sdate

xmlDoc.Save(XMLDataFile)

Set xmlDoc = nothing
Set FileLocation = nothing
	Actual_Res = "Today's Date is updated successfully in the XML response file"
	Reporter.ReportEvent micPass, StepName, Actual_Res

End if
	 'Handling Error
	methodName = "ToUpdateDateinXML" : rc = ErrorHandler(methodName)
End Function
'' --------------------------- End of Function ToReplaceDate() --------------------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: Validatecomplete_MasterRequestId(inputData,inputData1)
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		:inputData- > BAN to select the row in the list
'#					:inputData1- > status value as per the input
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function Validatecomplete_MasterRequestId(inputData)
		On Error Resume Next
	Dim methodName,rc, i, master_reqid, row_cnt
	methodName = "Validatecomplete_MasterRequestId" : Validatecomplete_MasterRequestId = 0
	
	If Exec_Flag = "Y" Then

rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Exist
'msgbox rc

If rc= "True" Then

JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Highlight
JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Press(4)

rc1 = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Exist
'msgbox rc1
JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Highlight


JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Press(4)
							wait(5)
							
rc2 = JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(10) 
'msgbox rc2

					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(1) Then
					wait (3)
							Flag = True
							Step_Description = "To Verify if Filter Process Instances Window is Opened"
							Exp_Result = "Filter Process Instances Window should be Loaded"
							Actual_Res = "Filter Process Instances Window is Loaded"
							Reporter.ReportEvent micPass, StepName, Actual_Res
						Else
							Call captureScreen
							Step_Description = "To Verify if Filter Process Instances Window is Opened"
							Exp_Result = "Filter Process Instances Window should be Loaded"
							Actual_Res = "Filter Process Instances Window is not Loaded"
							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
							Validatecomplete_MasterRequestId=-1
					End If
					
					
			
		JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").highlight
		rc = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(10)

			If rc = "True" Then
	
				row_cnt = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetROProperty("rows")
				For i = 0 To row_cnt-1
					master_reqid = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetCellData(i,1)
					'msgbox process_name
						If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
						print "Entered master request is found" &inputData
							'get the value of the status code
						
							status_code = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetCellData(i,4)
									If Trim(Ucase(status_code)) = Trim(ucase("Completed")) Then
																					print "Required status code is found" &status_code
																					
																					JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ClickCell i,2
																					
																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ClickCell i,1
																				
																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ActivateRow i
																				
																				'JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").DoubleClickCell i,1
																				
																				wait(2)
																				
																				Flag = True
																					else
																					'Call captureScreen
																					Step_Description = "The Status of the selected Master Request id is not in progress"
																					Exp_Result = "The Status of the selected Master Request id is in progress"
																					Actual_Res = "The Status of the selected Master Request id is not in progress"
																					Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
																					Validatecomplete_MasterRequestId=-1
																					End if
															Step_Description = "To select the required Master Request id"
															Exp_Result = "Master Request Id should be displayed"
															Actual_Res = "Master Request Id is displayed"
															Reporter.ReportEvent micPass, StepName, Actual_Res
								Exit For						
								Else
															'Call captureScreen
'															Step_Description = "To select the required Master Request id"
'															Exp_Result = "Master Request Id should be displayed"
'															Actual_Res = "Master Request Id is not displayed"
'															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'															Loading_MasterRequestId=-1
'												
														End If
						Next
								
'								Do
'								JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Highlight
'								display = JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Exist
'								Loop While display = False
					
														If display = "True" Then
															Step_Description = "To Verify if the process for the selected master request id is loaded"
															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
															Actual_Res = "Process is Loaded successfully for given Master Request Id" &inputData
															Reporter.ReportEvent micPass, StepName, Actual_Res
														Else
															'Call captureScreen
															Step_Description = "To Verify if the process for the selected master request id is loaded"
															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
															Actual_Res = "Process is not Loaded successfully for given Master Request Id" &inputData
															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
															Validatecomplete_MasterRequestId=-1	
														End If
			End if
			End if
End If

wait(5)
JavaWindow("title:=MDW Designer.*").Close

		 'Handling Error
	methodName = "Validatecomplete_MasterRequestId" : rc = ErrorHandler(methodName)
End function

' --------------------------- End of Function Validatecomplete_MasterRequestId() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: CopyTextToXML(inputData,inputData1)
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		:inputData- > BAN to select the row in the list
'#					:inputData1- > status value as per the input
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function CopyTextToXML()
		On Error Resume Next
	Dim methodName,rc, i, master_reqid, row_cnt
	methodName = "CopyTextToXML" : CopyTextToXML = 0
	
	If Exec_Flag = "Y" Then
	
	MyFile = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.txt"

		If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		Else
		    rc = "False"
		End If

		If rc = "True" Then
		

fso.CopyFile "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.txt","C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.xml"

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False 
'path of XML file

XMLDataFile="C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\Validate_Bus_Listener.xml"

'Load the XML File
 xmlDoc.Load(XMLDataFile)
'get the tagname to verify
Set nodelist = xmlDoc.getElementsByTagName("RequestType")
strText = nodelist.item(0).text

If strText<>"" Then
	rc = 0
	Step_Description = "Copy the contents of the text file to xml file"
	Exp_Result = "Xml value should be collected successfully"
	Actual_Res = "XML value collected is-> "& strText 
	Reporter.ReportEvent micPass, StepName, Actual_Res
	
	Else
	
	Step_Description = "Copy the contents of the text file to xml file"
	Exp_Result = "Xml value should be collected successfully"
	Actual_Res = "XML value collected is not collected successfully for the requested node->"
	Reporter.ReportEvent micFail, StepName, Actual_Res
	XMLValue_Collect =-1
End If

'qfile.writeline val2
'
'msgbox val2
'
'qfile.Close

Set qfile=nothing
Set fso=nothing

End if
End if
		 'Handling Error
	methodName = "CopyTextToXML" : rc = ErrorHandler(methodName)
End function

' --------------------------- End of Function CopyTextToXML() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: ToUpdateTrackingNumberinXML()
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		:inputData- > BAN to select the row in the list
'#					:inputData1- > status value as per the input
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
Public Function ToUpdateTrackingNumberinxml(inputData)
On Error Resume Next

Dim FileLocation,sdate,xmlDoc,XMLDataFile,methodName,rc,nodelist,strText,str1,remval,str2,strTextnew,strTrNum,val1,val2,strTrNumUrl
methodName = "ToUpdateTrackingNumberinxml" : ToUpdateTrackingNumberinxml = 0

	Step_Description = "To update ToUpdateTrackingNumberinxml in the given XML response file -> "
	Exp_Result = " ToUpdateTrackingNumberinxml is updated successfully in the XML response file"

MyFile = ("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&inputData&".xml")

		'If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		'Else
		'    rc = "False"
		'End If

		If rc = "True" Then
		
	
FileLocation = ("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&inputData&".xml")
print "Filelocations is:" &FileLocation
	
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False 

XMLDataFile=FileLocation

xmlDoc.Load(XMLDataFile)

If (inputData = "Ship")  or (inputData ="FF_MASTER_PROCESS_REDESIGN_FF_Shipment" ) Then

		Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
		strText = nodelist.item(0).text
				
			print "The TrackingNumber number is "&strText
				
			str1 = left(strText,14)
			print "First 14 digits of the TrackingNumber is "&str1
			
			remval = right(strText,4)
			print "Last 4 digits of the TrackingNumber is "&remval
			
			str2 = remval+10
			print "Updating the last 3 digits as "&str2
			
			strTextnew = str1&str2
			print "The new last 4 digits of the TrackingNumber is "&strTextnew
			
			Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingNumber")
			'msgbox strTrNum.text
			strTrNum.text = strTextnew
			
			Actual_Res = "Updated TrackingNumber in the XML file" 
			Reporter.ReportEvent micPass, StepName, Actual_Res
		 
			 
			'To update the TrackingURL tag in the Ship.xml response file
			Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
			strText = nodelist.item(0).text
					
			print "The TrackingURL is"&strText
					
			val1 = split(strText,"trackNums=")
			print "The Tracking Number in the TrackingURL is "&val1(1)
			
			val2 = Replace(val1(1),remval,str2)
			print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
			
			strTextnew = Replace(strText,val1(1),val2)
			print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
				
		
			Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingURL")
			'msgbox strTrNumUrl.text
			strTrNumUrl.text = strTextnew

ElseIf (inputData = "VENDOR_DELIVERY_PROCESS_IOM_RESPONSE") or (inputData = "VENDOR_DELIVERY_PROCESS_IOM_RESPONSE") Then

Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
strText = nodelist.item(0).text
		
	print "The TrackingNumber number is "&strText
		
	str1 = left(strText,14)
	print "First 14 digits of the TrackingNumber is "&str1
	
	remval = right(strText,4)
	print "Last 4 digits of the TrackingNumber is "&remval
	
	str2 = remval+10
	print "Updating the last 3 digits as "&str2
	
	strTextnew = str1&str2
	print "The new last 4 digits of the TrackingNumber is "&strTextnew
	
	Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:TrackingInfo/qb:TrackingNumber")
	'msgbox strTrNum.text
	strTrNum.text = strTextnew
	
	Actual_Res = "Updated TrackingNumber in the XML file" 
	Reporter.ReportEvent micPass, StepName, Actual_Res
 
	 
	'To update the TrackingURL tag in the Ship.xml response file
	Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
	strText = nodelist.item(0).text
			
	print "The TrackingURL is"&strText
			
	val1 = split(strText,"InquiryNumber1=")
	print "The Tracking Number in the TrackingURL is "&val1(1)
	
	val2 = Replace(val1(1),remval,str2)
	print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
	
	strTextnew = Replace(strText,val1(1),val2)
	print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
		

	Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:TrackingInfo/qb:TrackingURL")
	'msgbox strTrNumUrl.text
	strTrNumUrl.text = strTextnew


ElseIf inputData ="PURETV_VENDOR_DELIVERY_PROCESS_RESPONSE"  Then

Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
strText = nodelist.item(0).text
		
	print "The TrackingNumber number is "&strText
		
	str1 = left(strText,14)
	print "First 14 digits of the TrackingNumber is "&str1
	
	remval = right(strText,4)
	print "Last 4 digits of the TrackingNumber is "&remval
	
	str2 = remval+10
	print "Updating the last 3 digits as "&str2
	
	strTextnew = str1&str2
	print "The new last 4 digits of the TrackingNumber is "&strTextnew
	
	Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingNumber")
	'msgbox strTrNum.text
	strTrNum.text = strTextnew
	
	Actual_Res = "Updated TrackingNumber in the XML file" 
	Reporter.ReportEvent micPass, StepName, Actual_Res
 
	 
	'To update the TrackingURL tag in the Ship.xml response file
	Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
	strText = nodelist.item(0).text
			
	print "The TrackingURL is"&strText
			
	val1 = split(strText,"InquiryNumber1=")
	print "The Tracking Number in the TrackingURL is "&val1(1)
	
	val2 = Replace(val1(1),remval,str2)
	print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
	
	strTextnew = Replace(strText,val1(1),val2)
	print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
		

	Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingURL")
	'msgbox strTrNumUrl.text
	strTrNumUrl.text = strTextnew
	
End if


	Actual_Res = "Updated TrackingNumber in the XML file as-> "& strTextnew 
	Reporter.ReportEvent micPass, StepName, Actual_Res
	
xmlDoc.Save(XMLDataFile)

Set xmlDoc = nothing
End if
	 'Handling Error
	methodName = "ToUpdateTrackingNumberinxml" : rc = ErrorHandler(methodName)
End Function

'' --------------------------- End of Function ToUpdateTrackingNumberinxml() ---------------------------------------------------------
'___________________________________________________________________________________________________________________________
'# Function Name	: Copy_ExcelvaluetoExcel()
'# Purpose			: To collect XML value
'# Parameter                                			:row- > row in which the copied value to be pasted 
'#                                                      :col-> col in which the copied value to be pasted 
'#                                                       :inputData -> to give the value which has to be pasted in the provided row and col
'#                                                       
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________

Public Function Copy_ExcelvaluetoExcel(row,col)
Flag = True
                                                                                
                                                      'msgbox row 2,4 7,4
                                                      'msgbox col
                                                      'msgbox inputdata
                                                      'On Error Resume Next
                                                      'i=1
                                                      'rc=trim(inputData)
                                                      'r = split(rc,",")
                                                      'row = r(0) 
                                                      'column =  r(1)-1
                                                 
                                                      	
                                                      
                                                      			  Set obj=Description.Create()
'                                                                 obj("Class Name").Value="JavaToolbar"
'                                                                 obj("attached text").Value="Product Home"
'                                                                 Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
'                                                                     For i = 0 To obj1.Count-1
'                                                                     obj1(i).Highlight
'                                                                     obj1(i).Press "Copy RX# to clipboard"
                                                                      Dim objCB,str_1,xl
                                                                      'Set objCB= CreateObject("Mercury.Clipboard")
                                                                      'str_1 = objCB.GetText
                                                                      'str_1 = FF_Requestid_VAL
                                                                       set xl=CreateObject("Excel.Application") 
                                                                       set workbook= xl.Workbooks.Open("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx")  
                                                                       set sheet=workbook.Sheets("Query")
                                                                       inputdata = sheet.cells(17,4).value
                                                                       inputdata1="abcd"&inputdata
                                                                       sheet.Cells(row,col).value=inputdata1
                                                                       'sheet.Cells(row,column).interior.colorindex=6
                                                                        xl.ActiveWorkbook.Save
                                                                        'xl.ActiveWorkbook.Close
                                                                        'set xl.Application=Nothing
                                                                        'Set xl=Nothing
                                                                         xl.Workbooks.Close
																		 xl.Quit
									
'                                                                                Next
                                               
           ' Handling Error
	methodName = "Copy_ExcelvaluetoExcel" : rc = ErrorHandler(methodName)                                                     
End Function
' --------------------------- End of Function Copy_xmlvaluetoexcel() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: ReadCellvalueFromExcel()
'# Purpose			: To collect XML value
'# Parameter                                			:row- > row in which the copied value to be pasted 
'#                                                      :col-> col in which the copied value to be pasted 
'#                                                       :inputData -> to give the value which has to be pasted in the provided row and col
'#                                                       
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________

Public Function ReadCellvalueFromExcel(row,col,Filename)
Flag = True
                                                                                
                                                      'msgbox row
                                                      'msgbox col
                                                      'msgbox inputdata
                                                      'On Error Resume Next
                                                      'i=1
                                                      'rc=trim(inputData)
                                                      'r = split(rc,",")
                                                      'row = r(0) 
                                                      'column =  r(1)-1
                                                 
                                                      	
                                                      
          Set obj=Description.Create()
'         obj("Class Name").Value="JavaToolbar"
'         obj("attached text").Value="Product Home"
'         Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
'         For i = 0 To obj1.Count-1
'         obj1(i).Highlight
'         obj1(i).Press "Copy RX# to clipboard"
          Dim objCB,str_1,xl
          'Set objCB= CreateObject("Mercury.Clipboard")
          'str_1 = objCB.GetText
          'str_1 = FF_Requestid_VAL
          set xl=CreateObject("Excel.Application") 
		' set workbook= xl.Workbooks.Open("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\DB_Detail.xlsx")  
		   set workbook= xl.Workbooks.Open("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\Data\"&Filename&".xlsx",1) 
		If Filename = "DB_Detail" Then
			set sheet=workbook.Sheets("Query")
			ElseIf Filename = "FFWF_Data" Then
			set sheet=workbook.Sheets("Test_Data")
		End If
		inputdata = sheet.cells(row,col).value
		
		leng =  len(inputdata)
		'msgbox leng
		
		str1 = right(inputdata,4)
		'msgbox str1
		
		str1lenght = len(str1)
		'msgbox str1lenght
		
		lengrem = leng - str1lenght
		'msgbox lengrem
		
		remval = left(inputdata,lengrem)
		'msgbox remval
		
		str2 = str1+10
		'msgbox str2
		
		inputdata1 = remval&str2
		print "old data = " &inputdata & " Now the new val is = " &inputdata1
		sheet.Cells(row,col).value=inputdata1
		'sheet.Cells(row,column).interior.colorindex=6
         xl.ActiveWorkbook.Activate
         xl.ActiveWorkbook.Save
         xl.ActiveWorkbook.Close
        'set xl.Application=Nothing

        ' xl.Workbooks.Close
		 xl.Quit
        Set xl=Nothing
        systemutil.CloseProcessByName "EXCEL.EXE"
        wait(2)
		'  Next
                                               
           ' Handling Error
	methodName = "ReadCellvalueFromExcel" : rc = ErrorHandler(methodName)                                                     
End Function
' --------------------------- End of Function ReadCellvalueFromExcel() --------------------------------------------------------------------

'___________________________________________________________________________________________________________________________
'# Function Name	: IOM_Login()
'# Purpose			: To Login to MDW Designer with the given user id and password
'# Parameters		: inputData		-> Enter the username
'#					  inputData1	-> Enter the password
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________


Public Function IOM_Login(inputData,inputData1,inputData2,inputData3)
	On Error Resume Next
	Dim methodName, rc, display, i
	methodName = "IOM_Login" : MDWDesigner_Login = 0
	Step_Description = "IOM should be Launched and Logged in Successfully"
	Exp_Result = "IOM is Launched and Logged in Successfully"

If Exec_Flag = "Y" Then
	If Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("html id:=user","html tag:=INPUT","name:=user").Exist(2) Then
		    rc = "True"
		    		Else
		    rc = "False"
	End If
                                  
	If rc ="True" Then
		 Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("html id:=user","html tag:=INPUT","name:=user").Highlight
		 Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("html id:=user","html tag:=INPUT","name:=user").Set inputData
		 
		 Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("html id:=password","html tag:=INPUT","name:=password").set inputData1
		 
 		Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebButton("html id:=ctGoButton","html tag:=INPUT","type:=submit","value:=Login").Click 
 		Else
 		msgbox "end"
	End If
	 	
	 	
	Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").Link("name:=Event Trigger").Highlight
	Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").Link("name:=Event Trigger").Click	
	Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebList("name:=mainHelperForm:detailForm:eventNameSelect").Select "SERVICE_ORDER_EVENT_BUNDLING"	
	
	
	'for the split functionality
Dim firstLine,tempFile,qfile, qFile1,sResponsexml

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.Async = False 
			
	Set fso=createobject("Scripting.FileSystemObject")
			
	'Set qfile=fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\AVS_Order_Fulfillment.xml",1)
	Set qfile=fso.OpenTextFile("C:\jenkins\workspace\FFWF_QA_Automation\FFWF\XML\"&inputData2&".xml",1)
	qFile1 = qfile.ReadAll
	
	tempFile= Split(qFile1, vbCrLf)
	

	For i = 0 To 0
		firstLine = tempFile(i)
	Next	
	 Print firstLine
	
	'a = Split(objFSO.OpenTextFile(strDesktop & "\folder\blabla.vbs", ForReading).ReadAll, vbCrLf)

LineCount = UBound(tempFile) + 1

For i = 2 To LineCount
    If UBound(tempFile) >= i Then sResponsexml = sResponsexml & tempFile(i) & vbCrLf
Next

	print "Response xml is"&sResponsexml
			
	Set qfile=nothing
	Set fso=nothing
			
	Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:eventMessageTextarea").Highlight
			
	
	wait (2)
    Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:eventMessageTextarea").Click
	Set ws=CreateObject("wscript.shell")
		      	ws.SendKeys ("^a")	
			    ws.SendKeys "{DELETE}"
			    'ws.SendKeys sResponsexml
			    
				wait (2)
		
    Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:eventMessageTextarea").Set sResponsexml
	Actual_Res = "XML is loaded successfully from the XML file"
	Reporter.reportevent micPass, StepName, Actual_Res
	     		Set ws=Nothing 
	     		
	Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:processParametersTextarea").Set firstLine 	 	
	    		
    Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=text","html tag:=INPUT","name:=mainHelperForm:detailForm:masterRequestIdInput").Set inputData3	

    Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebButton("html id:=mainHelperForm:sendMessageButton","html tag:=INPUT","type:=submit","value:=Send Message").Click 

End if	
	 	
	' #comments: Handling Error
	methodName = "IOM_Login" : rc = ErrorHandler(methodName)
End Function

' --------------------------- End of IOM_Login() ----------------------------------------------------------------------



'___________________________________________________________________________________________________________________________
'# Function Name	: TestDataSetUp()
'# Purpose			: To Login to MDW Designer with the given user id and password
'# Parameters		: inputData		-> Enter the username
'#					  inputData1	-> Enter the password
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________


Public Function TestDataSetUp()
	On Error Resume Next
	Dim methodName, rc, display
	methodName = "TestDataSetUp" : MDWDesigner_Login = 0
	Step_Description = "New Test Data should be uploaded in the DataSheet successfully"
	Exp_Result = "New Test Data should be uploaded in the DataSheet successfully"

If Exec_Flag = "Y" Then

ReadCellvalueFromExcel	3,11,"FFWF_Data"
ReadCellvalueFromExcel	3,12,"FFWF_Data"
ReadCellvalueFromExcel	4,11,"FFWF_Data"
ReadCellvalueFromExcel	4,12,"FFWF_Data"
ReadCellvalueFromExcel	5,11,"FFWF_Data"
ReadCellvalueFromExcel	6,11,"FFWF_Data"
ReadCellvalueFromExcel	7,11,"FFWF_Data"
ReadCellvalueFromExcel	8,11,"FFWF_Data"
ReadCellvalueFromExcel	9,11,"FFWF_Data"
ReadCellvalueFromExcel	10,11,"FFWF_Data"
ReadCellvalueFromExcel	10,12,"FFWF_Data"
ReadCellvalueFromExcel	14,11,"FFWF_Data"
ReadCellvalueFromExcel	14,12,"FFWF_Data"
ReadCellvalueFromExcel	17,11,"FFWF_Data"
ReadCellvalueFromExcel	17,12,"FFWF_Data"
ReadCellvalueFromExcel	18,11,"FFWF_Data"
ReadCellvalueFromExcel	15,3,"DB_Detail"
ReadCellvalueFromExcel	16,3,"DB_Detail"
ReadCellvalueFromExcel	17,3,"DB_Detail"
ReadCellvalueFromExcel	18,3,"DB_Detail"
ReadCellvalueFromExcel	19,3,"DB_Detail"
ReadCellvalueFromExcel	20,3,"DB_Detail"
ReadCellvalueFromExcel	21,3,"DB_Detail"
ReadCellvalueFromExcel	22,3,"DB_Detail"
ReadCellvalueFromExcel	23,3,"DB_Detail"
ReadCellvalueFromExcel	24,3,"DB_Detail"
ReadCellvalueFromExcel	25,3,"DB_Detail"
ReadCellvalueFromExcel	26,3,"DB_Detail"
ReadCellvalueFromExcel	27,3,"DB_Detail"
ReadCellvalueFromExcel	28,3,"DB_Detail"
ReadCellvalueFromExcel	29,3,"DB_Detail"
ReadCellvalueFromExcel	30,3,"DB_Detail"
ReadCellvalueFromExcel	31,3,"DB_Detail"
ReadCellvalueFromExcel	32,3,"DB_Detail"
ReadCellvalueFromExcel	33,3,"DB_Detail"
ReadCellvalueFromExcel	34,3,"DB_Detail"

	Step_Description = "To Set Up new Test Data in all the Data Sheet for this Run"
	Exp_Result = "The Test Data has been set in all the Data Sheet successfully"
	Actual_Res = "The Test Data has been set in all the Data Sheet successfully" 
	Reporter.ReportEvent micPass, StepName, Actual_Res
	
	Else
	
	Step_Description = "To Set Up new Test Data in all the Data Sheet for this Run"
	Exp_Result = "The Test Data has NOT been set in all the Data Sheet successfully"
	Actual_Res = "The Test Data has NOT been set in all the Data Sheet successfully" 
	Reporter.ReportEvent micFail, StepName, Actual_Res
	TestDataSetUp =-1

                         
	
End if	
	 	
	' #comments: Handling Error
	methodName = "TestDataSetUp" : rc = ErrorHandler(methodName)
End Function

' --------------------------- End of TestDataSetUp() ----------------------------------------------------------------------

Dim method_Name : method_Name = "WINDOWS_KEYWORDS_LIB" : Call ErrorHandler(method_Name)

'#******************   End of WEB_KEYWORDS_LIB   ***************************************************************************
