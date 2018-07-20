'###########################################################################################################################

<<<<<<< HEAD
	APP_CONFIG_FILE = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\LIBRARY\FFWF_Config.vbs"	' CHANGE THE <APP_Name>
	LoadFunctionLibrary (APP_CONFIG_FILE)
			rc = TESTRUNNER()
=======
'										  ##
'##		CTL KWH-A AUTOMATION FRAMEWORK for QTP/UFT													

 @@ hightlight id_;_65894_;_script infofile_;_ZIP::ssf182.xml_;_
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

	APP_CONFIG_FILE = "C:\jenkins\workspace\FFWF_QA_Automation\FFWF\LIBRARY\FFWF_Config.vbs"	' CHANGE THE <APP_Name>
	LoadFunctionLibrary (APP_CONFIG_FILE)
	
	
	
	rc = TESTRUNNER()
 @@ hightlight id_;_1606351601_;_script infofile_;_ZIP::ssf183.xml_;_
 
	
>>>>>>> 6f0c0c5309f46d67885a570fba43c8643d5f027d
	If rc = 0 Then

		Reporter.ReportEvent micInfo, "TESTRUNNER", "Successfully Run the Test Case !" & Chr(13) & "Test Case :   " & UCASE(TESTCASE_NAME)
	
	Else

		Reporter.ReportEvent micFail, "TESTRUNNER", "Failed to Run the Test Case !" & Chr(13) & "Test Case :   " & UCASE(TESTCASE_NAME)
	
	End If

'###########################################################################################################################
<<<<<<< HEAD
=======




'
'Set fileSystemObj = createobject("Scripting.FileSystemObject")
'
'To check if the given file present'
'
'MyFile = "C:\Program Files\Parasoft\SOAtest\9.9\soatest.exe"
'
'If fileSystemObj.FileExists(MyFile) then
'    rc = "True"
'Else
'    rc = "False"
'End If
'
'		If rc = "True" Then
'			systemutil.Run"C:\Program Files\Parasoft\SOAtest\9.9\soatest.exe"
'			wait 2
'			Dialog("text:=Workspace Launcher","Location:=0").winbutton("text:=OK").Click
'			Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0").WinObject("nativeclass:=SysLink","regexpwndtitle:=.*Click here to activate license.*","attached text:=License is not active.*","index:=0").Click
'			wait 7
'
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst"
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite"
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Select "FFWF;FFWF.tst;Test Suite: Test Suite;Test 1: invokeWebService(string, string)"
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Activate "FFWF;FFWF.tst;Test Suite: Test Suite;Test 1: invokeWebService(string, string)"
'
'End if
'
'
'
'Public Function EXCEL()
'	On Error Resume Next
'	Dim methodName, rc, Temp_Name
'	methodName = "EXCEL" : EXCEL = 0
'	Execute("Test_URL = " & TEST_ENV)
'
' @@ hightlight id_;_327834_;_script infofile_;_ZIP::ssf101.xml_;_
'	'To generate Test Step Description and Expected Result
''	Step_Description = "Load XML from the Test Data Sheet" 
''	Exp_Result = " XML is loaded successfully from the Test Data Sheet"
'	
'	If Exec_Flag = "Y" Then
'	'Call close all open browser function
''		Call closeAllBrowser
'' @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf43.xml_;_
''		If TEST_BROWSER = "IE" Then 
''			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","",3
''		'ElseIf 	TEST_BROWSER = "Firefox" Then 
''			'SystemUtil.Run "firefox.exe",Test_URL,"C:\","", 3
''		'ElseIf 	TEST_BROWSER = "GChrome" Then 
''			'SystemUtil.Run "chrome.exe",Test_URL,"C:\","",3
''		Else 
''			TEST_BROWSER = "IE"		'	Default Browser
''			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","", 3
''		End If
'
'		'Browser("CreationTime:=0").Sync
'		'rc = EXIST_(pageDesc, "")
'
'
'Set fileSystemObj = createobject("Scripting.FileSystemObject")
'
''To check if the given file present'
'
'MyFile = "C:\automation\APPLICATIONS\FFWF\Data\FFWF_Data.xlsx"
'
'If fileSystemObj.FileExists(MyFile) then
'    rc = "True"
'Else
'    rc = "False"
'End If
'
'		If rc = "True" Then			
'
'Set oExcel=CreateObject("Excel.Application")
'Set oBook=oExcel.Workbooks.Open("C:\automation\APPLICATIONS\FFWF\Data\FFWF_Data.xlsx")
'Set oSheet=oBook.Worksheets("Test_Data")
'
'rows=oSheet.UsedRange.rows.count
'Cols=oSheet.UsedRange.Columns.count
'
'For i = 2 To rows Step 1
'	
'	Event_Name=oSheet.Cells(i,6).value
'    Exml=oSheet.Cells(i,7).value
'    
'    msgbox Event_Name
'    
'    msgbox Exml
'		
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Highlight
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").click
'		Set ws=CreateObject("wscript.shell")
'      	ws.SendKeys ("^a")	
'	    ws.SendKeys "{DELETE}"
'	    'ws.SendKeys Event_Name
'		'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Type Event_Name	    
'	    Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Type Event_Name	
'	    Set ws=Nothing 
'	     
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").Highlight
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").click
'		Set ws=CreateObject("wscript.shell")
'      	 ws.SendKeys ("^a")	
'	     ws.SendKeys "{DELETE}"
'	     'ws.SendKeys Exml
''	    Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinEditor("nativeclass:=Edit","location:=2").click
''		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinEditor("nativeclass:=Edit","location:=2").Type micCtrlDwn + "a" + micCtrlUp
''		ws.SendKeys "{BACKSPACE}"
'		
'		wait (2)
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","location:=2").Type Exml
'	     Set ws=Nothing 
'
'wait (2)
'
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinToolbar("nativeclass:=ToolbarWindow32","location:=8").Press 1 @@ hightlight id_;_2426152_;_script infofile_;_ZIP::ssf10.xml_;_
'wait 2
'If Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("text:=Save Resource").WinButton("text:=&Yes").Exist(5) Then
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("text:=Save Resource").WinButton("text:=&Yes").Click
'End if 
'
'
'if Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("attached text:=Finished","regexpwndtitle:=1/1 Tests Succeeded").Exist Then
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("attached text:=Finished","regexpwndtitle:=1/1 Tests Succeeded").Highlight
'msgbox "Tests Succeeded successfully"
'End if
'Next
'
''oBook.Save
'oBook.Close
'oExcel.Quit
'
'End if
'End if
'End Function

'set user name and password'


'___________________________________________________________________________________________________________________________
'# Function Name	: Select_Process()
'# Purpose			: To Select the Process from the Process list
'# Parameters		: inputData -> Enter Which process you want to select
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
'Public Function Select_Process(inputData)
'	On Error Resume Next
'	Dim methodName,rc, i, process_name, row_cnt
'	methodName = "Select_Process" : Select_Process = 0
'	
'inputData = "PureTvOrderFulfillmentProcess"
'Exec_Flag = "Y"
'	If Exec_Flag = "Y" Then
'
'		JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").highlight
'		rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").Exist(10)
'
'msgbox rc
'			If rc = "True" Then
'	
'				row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
'				For i = 0 To row_cnt-1
'					process_name = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
'					msgbox process_name
'						If Trim(Ucase(process_name)) = Trim(ucase(inputData)) Then
'						msgbox "found" &inputData
'							JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").ClickCell i,1
'							JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell i,1
'							Flag = True
'							Step_Description = "To select the required process"
'							Exp_Result = "Required Process should be displayed"
'							Actual_Res = "Required Process is displayed"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'				'			Call captureScreen
'							Step_Description = "To verify update status"
'							Exp_Result = "Required Process should be displayed"
'							Actual_Res = "Required Process is not displayed"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Select_Process=-1
'				Exit For
'						End If
'						Next
'						
'						Do
'						display = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaObject("tagname:=DesignerCanvas","toolkit class:=com\.qwest\.mdw\.designer\.pages\.DesignerCanvas").Exist
'						Loop While display = False
'			
'						If display = "True" Then
'							Step_Description = "To Verify if the selected process is loaded"
'							Exp_Result = "Selected Process should be loaded"
'							Actual_Res = "Selected Process is loaded sucessfully"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'				'			Call captureScreen
'							Step_Description = "To Verify if the selected process is loaded"
'							Exp_Result = "Selected Process should be loaded"
'							Actual_Res = "Selected Process is not loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Select_Process=-1	
'						End If
'			End If
'	End if
''End function

'Public Function Generic_CLICK1(inputData)
'	On Error Resume Next
'	Dim methodName,ObjDesc_Array, Object_Arr, Object_Name, objName, rc, exec_Stmt ,exec_stmnt, Flag,val,val1,Object_Desc
'	methodName = "Generic_CLICK1" : Generic_CLICK1 = 0
'	Flag = False
'
'	ObjDesc_Array = split(inputData,"=>")
'	'msgbox ObjDesc_Array(0)
'	ObjSplit_Array = split(ObjDesc_Array(1)," + ",-1,1) : ObjSplit_Count = UBOUND(ObjSplit_Array)
'	Select Case (ObjSplit_Count)
'		Case "0"
'			Object_Desc = ObjDesc_Array(1)
'			
'		Case "1"
'			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1)
'			'msgbox Object_Desc
'		Case "2"
'			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2)
'			
'		Case "3"
'			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3)
'			
'		Case "4"
'			Object_Desc = ObjSplit_Array(0) & """, """ & ObjSplit_Array(1) & """, """ & ObjSplit_Array(2) & """, """ & ObjSplit_Array(3) & """, """ & ObjSplit_Array(4)
'			
'		Case Else
'			Step_Description = "Object property"
'			Actual_Res = "Object property exceeded more than five"
'			Reporter.ReportEvent micFail, StepName, Actual_Res
'			Generic_CLICK1 = -1
'	End Select
'
'	'To generate Test Step Description and Expected Result
'	Step_Description = "Click on object -> " & objName & "  in page generic page"
'	Exp_Result = "Clicked on object ->"& objName & " in generic page"
'
'	If Exec_Flag = "Y" Then
'		exec_stmnt = "rc = Browser(""creationtime:=2"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Exist(""10"")"
'		'msgbox exec_stmnt
'		Execute(exec_stmnt)
'		'msgbox rc
'		If rc = "True" Then
'			Browser("CreationTime:=2").Page("title:=.*").highlight
'			exec_Stmt = "Browser(""creationtime:=2"")."&"Page(""title:=.*"&""")." & ObjDesc_Array(0) & "(""" & Object_Desc & """).Click"
'			Execute(exec_Stmt)
'			Flag = True
'			Actual_Res = "Click on object -> " & objName & " in generic page is successful"
'			Reporter.ReportEvent micInfo, StepName, Actual_Res
'		Else Generic_CLICK1 = -1 End If
'If Flag = False Then
'	Actual_Res = "Click on object -> " & objName & " in generic page failed"
'	Reporter.ReportEvent micFail, StepName, Actual_Res
'	Generic_CLICK1 = -1
'End If
'			
'End If
'	' Handling Error
'	methodName = "Generic_CLICK1" : rc = ErrorHandler(methodName)
'End Function
'	
'	
'	
'***************Opening TMS bus listener in IE*******

'systemutil.Run "iexplore.exe", "http://x7009075/TMSWebUtilities/TMSBusTester.aspx"
'Wait(35)
'Dim bp
'Set bp=Browser("title:=TMS Bus Listener - Internet Explorer provided by CenturyLink","name:=TMS Bus Listener").Page("title:=TMS Bus Listener","url:=http://x7009075/TMSWebUtilities/TMSBusTester.aspx")
'bp.WebButton("name:=Show Request","html tag:=INPUT","html id:=btnSubmitSection").Click
'Wait(2)
''**************Setting up the Test Env**************
'
'bp.WebList("html id:=ddlBusSettings").select "BusITV1"
'
'wait(3)
 @@ hightlight id_;_Browser("TMS Bus Listener").Page("TMS Bus Listener").WbfGrid("MenuTable")_;_script infofile_;_ZIP::ssf170.xml_;_


'ExitTest
'Call Triggering()


'Call UpdateXmlwithcurrentRequestidandDate("FF_Requestid_VAL")
'
'Function UpdateXmlwithcurrentRequestidandDate(FF_Requestid_VAL)
'    
'sdate= Year(date)&"-"&month(date)&"-"&day(date)
'Set xl=CreateObject("Excel.Application")
'
'Set wb=xl.Workbooks.Open("C:\Users\ab51923\Desktop\FFWF_Automation\Bus.xlsx")
'Set sheet=wb.Worksheets("Sheet1")
'var= sheet.Cells(2,2).value
'
''***********************************Request id updation*************************************
'If instr(1,var,"<RequestId>")<>0 Then
'sReq1= instr(1,var,"<RequestId>") 
'sReq2= instr(1,var,"</RequestId>")
'
'For i = sReq1+11 To sReq2-1 Step 1
'     req=mid(var,i,1)
'sReq=sReq&req
'Next    
'
'End If
'
'sheet.Cells(2,2).value= Replace(var,sReq,FF_Requestid_VAL)
'
'If instr(1,sheet.Cells(2,2).value,FF_Requestid_VAL)<>0 Then
'	Print "Following Request ID: "&FF_Requestid_VAL&"  is successfully Updated in xml code"
'	Else
'	Print "Request ID is not successfully Updated in xml code"
'End If
'
'wb.Save
'wb.Close
'
'xl.Quit
'Set xl=Nothing
'Set wb=Nothing
'Set sheet=Nothing
'
'Call ExcelApp()
'
'End Function





'Browser("name:=TMS Bus Listener.*").Page("title:=.*").WebEdit("name:=tbMessages").Highlight
'val = Browser("name:=TMS Bus Listener.*").Page("title:=.*").WebEdit("name:=tbMessages").GetROProperty("value")
'msgbox val
'val1 = split(val,"is:")
'msgbox val1(1)
'val2 = Replace(val1(1),"_______________________________________________________"," ")
'msgbox val2
'Set fso=createobject("Scripting.FileSystemObject")
'Set qfile=fso.OpenTextFile("C:\Users\ab65745\Desktop\testing_xml.txt",2,True)
'qfile.writeline val2
'
'qfile.Close
'
''Release the allocated objects
'Set qfile=nothing
'Set fso=nothing
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
'
'XMLDataFile="C:\Users\ab65745\Desktop\testing_xml.txt"
'
' xmlDoc.Load(XMLDataFile)
'
'Set nodelist = xmlDoc.getElementsByTagName("bim:RequestId")
'
'strText = nodelist.item(0).text
'  
'msgbox strText



'xml_val = val
''msgbox xml_val
'Set obj = createobject("scripting.filesystemobject")
'
'set xml_cnt = obj.OpenTextFile("C:\automation\APPLICATIONS\FFWF\XML\Xml_val.txt",2,true)
'
'xml_cnt.Writeline xml_val
'xml_cnt.Close
'
'
'
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
''path of XML file
'
'XMLDataFile="C:\automation\APPLICATIONS\FFWF\XML\Xml_val.xml"
'
''Load the XML File
' xmlDoc.Load(XMLDataFile)
''get the tagname to verify
'Set nodelist = xmlDoc.getElementsByTagName("bim:RequestId")
''msgbox nodelist.length
'val =  nodelist.length
'
'For m = 0 to val
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

'
'
'Function ToUpdateDateinXML(FileName)
'
'FileLocation = "C:\automation\APPLICATIONS\FFWF\XML\" &FileName &".XML"	
'	If FileName = "Resp" Then
'		Tagtobereplaced = "ns2:FulfillmentResponse/ns2:ResponseHeader/SendTimeStamp"
'		Else
'		 Tagtobereplaced = "ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp"
'	End If
'sdate= Year(date)&"-"&month(date)&"-"&day(date)
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
'XMLDataFile=FileLocation
'
'xmlDoc.Load(XMLDataFile)
''get the tagname to verify
'Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)
'
'nodelist.text = sdate
'
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
'
''End if
''	 Handling Error
''	methodName = "ToUpdateDateinXML" : rc = ErrorHandler(methodName)
'End Function


'Set DbConn=CreateObject("ADODB.Connection")
'Set rc=Createobject("ADODB.Recordset") 
''Protocol="TCP"
'
''
' DBQuery = "SELECT * FROM FFWF.fulfillment_request WHERE request_owner_id = (SELECT delivery_detail_id FROM FFWF.delivery_detail WHERE order_id = ( SELECT Order_id FROM FFWF.ordr WHERE customer_id =( SELECT DISTINCT customer_id FROM shippable_instance WHERE eno_trans_id= '0120333444600966667773')));"
'
''ConnectionString="Driver={Oracle in OraClient 12home1_32bit};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=RXTEST11DB.DEV.QINTRA.COM)(PORT=1569))(CONNECT_DATA=(TNS Service Name=RXTEST11DB.DEV.QINTRA.COM:1569/RXTEST11)));Data Source=Rx DB;User ID=DRDSLAPP;Password=drdslapp;"
'ConnectionString="Driver={Oracle in OraClient 12home1_32bit};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=ffwfst1db.dev.qintra.com)(PORT=1539))(CONNECT_DATA=(TNS Service Name=ffwfst1db.dev.qintra.com:1539/ffwfst1)));Data Source=FFWF DB;User ID=ffwf_app;Password=ffwfst1_suomt102;"
'
''On error resume next
'DbConn.Open ConnectionString 
'    
'rc.Open DBQuery,DbConn
'
'Do while not rc.eof
'
' msgbox rc("MOD_USER").value    
'stop
'
''msgbox ucid
'
'rc.movenext
'Loop
'
'Set DbConn = Nothing
'Set rc = Nothing


'Public Function FFWF_DB_Connect(Query_ID,DBQuery,Fieldname)
'On Error Resume Next

Rem Creating a datatable and storing the data from the Excel sheet



'Sheetname = "Query"
'FilePath="C:\automation\APPLICATIONS\FFWF\Data\DB_Details.xlsx" 
'Datatable.AddSheet Sheetname
'Datatable.ImportSheet FilePath, Sheetname, Sheetname
'
'
'verify_First_Data_from_Datatable= Datatable.Value("Query",Sheetname)
'Print verify_First_Data_from_Datatable
'
'getRowCount=Datatable.GetSheet(Sheetname).GetRowCount
'getParamCount=Datatable.GetSheet(Sheetname).GetParameterCount
'Print "getParamCount ="&getParamCount
'
'Print "Total Rows ="&getRowCount
'Datatable.GetSheet(Sheetname).SetCurrentRow(1)
'For Iterator = 1 To getRowCount Step 1
'	Print "Iterator ="&Iterator
'	Query_ID="Q"&Iterator
'	Print "Query_ID ="&Query_ID
'	getRowVal=Datatable.GetSheet(Sheetname).GetParameter("Query").Value
'	Print "Query_Val ="&getRowVal
''	Datatable.SetNextRow
'
''Next
''
''
''ExitTest
'
'
'Dim myxl,mysheet,Row,Exec_Flag,Query_ID,DBQuery,Fieldname
''Set myxl = createobject("excel.application")
'
''myxl.Workbooks.Open "C:\automation\APPLICATIONS\FFWF\Data\DB_Details.xlsx" 
'
''set mysheet = myxl.ActiveWorkbook.Worksheets("Query")
'
''Row=mysheet.UsedRange.rows.count
'
''For i = 1 To Row
'    'If Trim(ucase(mysheet.cells(i,1).value)) <> ""  Then
'        'QID = abs(i)+1
'        
''msgbox QID
'
'
'Exec_Flag = "Y"
'
''Query_ID = mysheet.cells(QID,1).value
''Print "Query_ID ="&Query_ID
'
''DBQuery = mysheet.cells(QID,2).value
'
''Fieldname = mysheet.cells(QID,3).value
'
'
'
''End if
'
''myxl.Workbooks.Close
''myxl.Quit
'
'
'Sheetname = "Query"
'FilePath="C:\Automation\APPLICATIONS\FFWF\Data\DB_Details.xlsx"
'Datatable.AddSheet Sheetname
'Datatable.ImportSheet FilePath, Sheetname, Sheetname
'verify_First_Data_from_Datatable= Datatable.Value("Query",Sheetname)
'Print verify_First_Data_from_Datatable
'
'getRowCount=Datatable.GetSheet(Sheetname).GetRowCount
'getParamCount=Datatable.GetSheet(Sheetname).GetParameterCount
'Print "getParamCount ="&getParamCount
'
'Print "Total Rows ="&getRowCount
'Datatable.GetSheet(Sheetname).SetCurrentRow(1)
'For Iterator = 1 To getRowCount Step 1
'	Print "Iterator ="&Iterator
'	Query_ID="Q"&Iterator
'	Print "Query_ID ="&Query_ID
'	getRowVal=Datatable.GetSheet(Sheetname).GetParameter("Query").Value
'	Print "Query_Val ="&getRowVal
'
'Dim myxl,mysheet,Row,Exec_Flag,Query_ID,DBQuery,Fieldname
'Dim DbConn,rc
'Dim ConnectionString,CN,fso,port,SID,unamepasswd,host,ServiceName
'
'
'Step_Description = "Load DB Details from the Query Sheet" 
'Exp_Result = " Queries are loaded successfully from the Query Sheet"
'Exec_Flag = "Y"
'If Exec_Flag = "Y" Then
'
'Set fileSystemObj = createobject("Scripting.FileSystemObject")
'MyFile = "C:\automation\APPLICATIONS\FFWF\Data\DB_Details.xlsx"
'
'If fileSystemObj.FileExists(MyFile) then
'		    rc = "True"
'		Else
'		    rc = "False"
'		End If
'
'If rc = "True" Then			
'
'			
'		
'Set DbConn=CreateObject("ADODB.Connection")
'Set rc=Createobject("ADODB.Recordset") 
'
'												Actual_Res = "FFWF DB is Not connected Currently... Will establish the connection" 
'												Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'												FFWF_DB_Connect=-1
'											
'
'
'
'ConnectionString="Driver={Oracle in OraClient 12home1_32bit};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=ffwfst1db.dev.qintra.com)(PORT=1539))(CONNECT_DATA=(TNS Service Name=ffwfst1db.dev.qintra.com:1539/ffwfst1)));Data Source=FFWF DB;User ID=ffwf_app;Password=ffwfst1_suomt102;"
'
'
'DbConn.Open ConnectionString
'    
''rc.Open DBQuery,DbConn
'rc.Open getRowVal,DbConn
'
'Print "Connection Status ="&DbConn.State
'
'If DbConn.State=0 Then
'      
'		Actual_Res = "FFWF DB Connection Status Not Established" 
'		Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'		FFWF_DB_Connect=-1
'Else
'		Actual_Res = "FFWF DB Connection Status Established Successfully" 
'		Reporter.ReportEvent micInfo, StepName, Actual_Res
'End If
'
'  If err.number <> 0 Then
'										Call captureScreen
'										Actual_Res = "Not Able to query the required DB" 
'										Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'										FFWF_DB_Connect=-1
'                                else
'                                		Actual_Res = "Able to query the required DB successfully" 
'										Reporter.ReportEvent micInfo, StepName, Actual_Res
'
'                End If   
'
'
'
'If rc.EOF <> True Then
'                If Query_ID="Q1" Then
'                                FF_Requestid_VAL = rc("FULFILLMENT_REQUEST_ID").Value
'                                msgbox "The FULFILLMENT_REQUEST_ID for the FFWF Response is: " & FF_Requestid_VAL
'                               
'                                inputdata=FF_Requestid_VAL
'                                  Call  Copy_xmlvaluetoexcel(2,4,inputdata)
'                                  
'                End If
'
'                
'                If Query_ID="Q2" Then
'                                FF_ORDER_VAL = rc("ORDER_NUMBER").Value
'                                msgbox  "The ORDER_NUMBER for the FFWF Response is: "  &FF_ORDER_VAL
'                                
'                                inputdata=FF_ORDER_VAL
'                                Call  Copy_xmlvaluetoexcel(3,4,inputdata)
'                End If    
'                
'                If Query_ID="Q3" Then
'                                FF_Requestid_VAL2 = rc("FULFILLMENT_REQUEST_ID").Value
'                                msgbox  "The FULFILLMENT_REQUEST_ID for the FFWF Shipment Response is: "  &FF_Requestid_VAL2
'
'                                inputdata=FF_Requestid_VAL2
'                                Call  Copy_xmlvaluetoexcel(4,4,inputdata)
'                End If
'                
'                If Query_ID="Q4" Then
'                                FF_ORDER_VAL2 = rc("ORDER_NUMBER").Value
'                                msgbox  "The ORDER_NUMBER for the FFWF Shipment Response is: "  &FF_ORDER_VAL2
'                               
'                                inputdata=FF_ORDER_VAL2
'                                Call  Copy_xmlvaluetoexcel(5,4,inputdata)
'                End If
'                
'                If Query_ID="Q5" Then
'                                ITEM_ID_VAL2 = rc("SHIPPABLE_INSTANCE_ID").Value
'                                msgbox  "The SHIPPABLE_INSTANCE_ID(ITEM ID) for the FFWF Shipment Response is: "&ITEM_ID_VAL2
'                              
'                                inputdata=ITEM_ID_VAL2
'                                Call  Copy_xmlvaluetoexcel(6,4,inputdata)
'                End If
' 
' Else
'
'           
'                  
'End if
'
'
'
'Set DbConn = Nothing
'Set rc = Nothing
'
'End if
'End if
''Next
'
'	Datatable.SetNextRow
'Next
'
'
'ExitTest
'
'
''End Function
'
'
'Public Function Copy_xmlvaluetoexcel(row,col,inputdata)
'                                                                                
'                                                      msgbox row
'                                                      msgbox col
'                                                      msgbox inputdata
'                                                      'On Error Resume Next
'                                                      'i=1
'                                                      'rc=trim(inputData)
'                                                      'r = split(rc,",")
'                                                      'row = r(0) 
'                                                      'column =  r(1)-1
'                                                      			  Set obj=Description.Create
''                                                                 obj("Class Name").Value="JavaToolbar"
''                                                                 obj("attached text").Value="Product Home"
''                                                                 Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
''                                                                     For i = 0 To obj1.Count-1
''                                                                     obj1(i).Highlight
''                                                                     obj1(i).Press "Copy RX# to clipboard"
'                                                                      Dim objCB,str_1,xl
'                                                                      'Set objCB= CreateObject("Mercury.Clipboard")
'                                                                      'str_1 = objCB.GetText
'                                                                      'str_1 = FF_Requestid_VAL
'                                                                       set xl=CreateObject("Excel.Application")                                                                                           
'                                                                       set workbook= xl.Workbooks.Open("C:\automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx")  
'                                                                       set sheet=workbook.Sheets("Query")
'                                                                       sheet.Cells(row,col).value=inputdata
'                                                                       'sheet.Cells(row,column).interior.colorindex=6
'                                                                        xl.ActiveWorkbook.Save
'                                                                        'xl.ActiveWorkbook.Close
'                                                                        'set xl.Application=Nothing
'                                                                        'Set xl=Nothing
'                                                                         xl.Workbooks.Close
'																		 xl.Quit
''                                                                                Next
'                                                Step_Description = "To click on javatree"
'                                                Exp_Result = "Session Id should be copied"
'                                                Actual_Res = "Session Id for the current Login is "& str_1
'                                                Reporter.ReportEvent micInfo, StepName, Actual_Res
'                                                                
'End Function
'
''
''
''
'''Public Function ReadXMLFileAndReplaceXMLTags(Queryid,FileLocation,Tagtobereplaced,TagValue)
'
'QID
'QID=Queryid+1
'
'Set myxl = createobject("excel.application")
'
'myxl.Workbooks.Open "C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx" 
'
'set mysheet = myxl.ActiveWorkbook.Worksheets("Query")
'
'Row=mysheet.UsedRange.rows.count
'
'Exec_Flag = "Y"
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
'
'On Error Resume Next
'
'	Dim methodName,rc,strText,xmlDoc,XMLDataFile,nodelist
'
'	'methodName = "ReadXMLFileAndReplaceXMLTags" : ReadXMLFileAndReplaceXMLTags = 0
'	
'
''If Exec_Flag = "Y" Then
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
''path of XML file
'
'XMLDataFile=FileLocation
'
''Load the XML File
' xmlDoc.Load(XMLDataFile)
''get the tagname to verify
'Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)
'
'nodelist.text = TagValue
'
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
'
'End if
''	 Handling Error
''	methodName = "ReadXMLFileAndReplaceXMLTags" : rc = ErrorHandler(methodName)
'End Function


	

'
'JavaWindow("MDW Designer (ecomt199.dev.qintra.c").JavaDialog("Filter Process Instances").JavaButton("Load").Click @@ hightlight id_;_886710079_;_script infofile_;_ZIP::ssf172.xml_;_
'Exec_Flag = "Y" 
'Flag = False
'	If Exec_Flag = "Y" Then
'
'		JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Highlight
'		rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Exist(10)
'
'			'If rc = "True" Then
'			JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Press(5)
'			wait (2)
'				'JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaButton("tagname:=table24","toolkit class:=javax\.swing\.JButton").Click
'					If  JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").Exist(2) Then
'						JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").JavaButton("tagname:=OK","toolkit class:=javax.swing.JButton").Click
'						Else
'						
'					
'					
'					If JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=JPanel","toolkit class:=javax.swing.JPanel").Exist(10) Then 
'							Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							'Call captureScreen
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'					JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").JavaButton("attached text:=Load").Click
'					wait (2)
'					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(1) Then
'					wait (3)
'					Flag = True
'							Step_Description = "To Verify if the Total Instances Window is Opened"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							'Call captureScreen
'							Step_Description = "To Verify if the Total Instances Window is Opened"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'			End If		
			'End If
'	End if
	
	
	'Public Function Loading_MasterRequestId(inputData,inputData1)
	'	On Error Resume Next
	'Dim methodName,rc, i, master_reqid, row_cnt
	'methodName = "Loading_MasterRequestId" : Loading_MasterRequestId = 0
'	Exec_Flag = "Y"
'	inputData = "100001666"
'	inputData1 = "In Progress"
'	
'		If Exec_Flag = "Y" Then
'	
'			JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").highlight
'			rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(10)
'	
'					If rc = "True" Then
'			
'						row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
'						For i = 0 To row_cnt-1
'							master_reqid = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
'							'msgbox master_reqid
'														If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
'														'msgbox "found" &inputData
'															'get the value of the status code
'															status_code = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,4)
'																					If Trim(Ucase(status_code)) = Trim(ucase(inputData1)) Then
'																					'msgbox "found" &status_code
'																				JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").ClickCell i,1
'																				JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell i,1
'																				Flag = True
'																					else
'																					'Call captureScreen
'																					Step_Description = "The Status of the selected Master Request id is not in progress"
'																					Exp_Result = "The Status of the selected Master Request id is in progress"
'																					Actual_Res = "The Status of the selected Master Request id is not in progress"
'																					Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'																					Loading_MasterRequestId=-1
'																					End if
'															Step_Description = "To select the required Master Request id"
'															Exp_Result = "Master Request Id should be displayed"
'															Actual_Res = "Master Request Id is displayed"
'															Reporter.ReportEvent micInfo, StepName, Actual_Res
'								Exit For						
'								Else
'															'Call captureScreen
''															Step_Description = "To select the required Master Request id"
''															Exp_Result = "Master Request Id should be displayed"
''															Actual_Res = "Master Request Id is not displayed"
''															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
''															Loading_MasterRequestId=-1
''												
'														End If
'						Next
'								
'								Do
'								JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Highlight
'								display = JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Exist
'								Loop While display = False
'					
'														If display = "True" Then
'															Step_Description = "To Verify if the process for the selected master request id is loaded"
'															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
'															Actual_Res = "Process is Loaded successfully for given Master Request Id" &inputData
'															Reporter.ReportEvent micInfo, StepName, Actual_Res
'														Else
'															'Call captureScreen
'															Step_Description = "To Verify if the process for the selected master request id is loaded"
'															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
'															Actual_Res = "Process is not Loaded successfully for given Master Request Id" &inputData
'															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'															Loading_MasterRequestId=-1	
'														End If
'					End If
'				End if
			
'End function

'Call Login_SOA()

'Public Function Login_SOA()

'	On Error Resume Next
'	Dim methodName, rc, Temp_Name
'	methodName = "Login_SOA" : Login_SOA = 0
'	Execute("Test_URL = " & TEST_ENV)
'	
'	Test_URL = "http://ecomt200.dev.qintra.com:7622/FulfillmentWFMDWWeb/MDWWebService?WSDL"
'
'	'To generate Test Step Description and Expected Result
'	'Loc_name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
'	Step_Description = "Open SOA " 
'	Exp_Result = " SOA should open Successfully " 
'	Exec_Flag = "Y"
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
'
'Set fileSystemObj = createobject("Scripting.FileSystemObject")
'
''To check if the given file present'
'
'MyFile = "C:\Program Files\Parasoft\SOAtest\9.9\soatest.exe"
'
'		If fileSystemObj.FileExists(MyFile) then
'		    rc = "True"
'		Else
'		    rc = "False"
'		End If
'
'		If rc = "True" Then
'			systemutil.Run"C:\Program Files\Parasoft\SOAtest\9.9\soatest.exe"
'			wait 2
'			Dialog("text:=Workspace Launcher","Location:=0").winbutton("text:=OK").Click
'			Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0").WinObject("nativeclass:=SysLink","regexpwndtitle:=.*Click here to activate license.*","attached text:=License is not active.*","index:=0").Click
'			wait 7
'				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: MDWWebServicePort"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Select "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: MDWWebServicePort;Test 1: invokeWebService(string, string)"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Activate "FFWF;FFWF.tst;Test Suite: Test Suite;Test Suite: MDWWebServicePort;Test 1: invokeWebService(string, string)"
'				
'			Actual_Res = "Launched " & " SOA with Application Test URL -> " & Chr(13) & Test_URL & ". SOA is Launched -> " & Temp_Name
'			Reporter.ReportEvent micInfo, StepName, Actual_Res
'		Else
'			'Call captureScreen
'			Actual_Res = "Launched " & TEST_BROWSER & " SOA with Application Test URL -> " & Chr(13) & Test_URL & ". SOA is not Launched -> " & Temp_Name & ". Page Expected is " & Page_Name(1)
'			Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'			Login_SOA = -1
'		End If
'
'wait(2)
'
'					
'End if

'Handling Error
'	methodName = "Login_SOA" : rc = ErrorHandler(methodName)
'End Function

'Exec_Flag = "Y" 
'Flag = False
'	'If Exec_Flag = "Y" Then
'	
'	
'	If  JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").Exist(2) Then
'						JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").JavaButton("tagname:=OK","toolkit class:=javax.swing.JButton").Click
'						End if
'
'		JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Highlight
'		rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Exist(10)
'
'			If rc = "True" Then
'			JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Press(5)
'			wait (2)
'				JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaButton("tagname:=table24","toolkit class:=javax\.swing\.JButton").Click
'					
'						
'					
'					
'					If JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=JPanel","toolkit class:=javax.swing.JPanel").Exist(10) Then 
'							Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							Call captureScreen
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'					JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").JavaButton("attached text:=Load").Click
'					wait (2)
'					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(1) Then
'					wait (3)
'					Flag = True
'							Step_Description = "To Verify if the Total Instances Window is Opened"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							Call captureScreen
'							Step_Description = "To Verify if the Total Instances Window is Opened"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'			End If		
'			
	



''Public Function Load_ProcessInstance()
''	On Error Resume Next
''	Dim methodName,rc, Flag
''	methodName = "Load_ProcessInstance" : Load_ProcessInstance = 0
'	Flag = False
'	 Exec_Flag = "Y"
'	If Exec_Flag = "Y" Then

'		JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Highlight
'		rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Exist(10)
'
'			If rc = "True" Then
'			JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Press(5)
'			wait (2)
'				'JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaButton("tagname:=table24","toolkit class:=javax\.swing\.JButton").Click
'					If  JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").Exist(2) Then
'						JavaWindow("title:=MDW Designer.*").JavaDialog("title:=Error","toolkit class:=javax.swing.JDialog").JavaButton("tagname:=OK","toolkit class:=javax.swing.JButton").Click
'						Else
'						
'					
'					
'					If JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").Exist(10) Then 
'					JavaWindow("title:=MDW Designer.*").JavaDialog("label:=Filter Process Instances","toolkit class:=com.qwest.mdw.designer.dialogs.FilterDialog").JavaButton("attached text:=Load").Click
'					wait (2)
'					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(1) Then
'					wait (3)
'							Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							Call captureScreen
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'					
'					Flag = True
'							Step_Description = "To Verify if the Total Instances Window is Opened"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							Call captureScreen
'							Step_Description = "To Verify if the Total Instances Window is Opened"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'			End If		
'			End If
' @@ hightlight id_;_Browser("TMS Bus Listener").Page("TMS Bus Listener")_;_script infofile_;_ZIP::ssf175.xml_;_

''End function


''Public Function CopyRespxmltoBusListener()
''	On Error Resume Next
'	Dim methodName, rc, Temp_Name
''	methodName = "Buslistnerorder_xml" : Buslistnerorder_xml = 0
''	Execute("Test_URL = " & TEST_ENV)
''
'
'	'To generate Test Step Description and Expected Result
'	Step_Description = "Load XML from the Test Data Sheet" 
'	Exp_Result = " XML is loaded successfully from the Test Data Sheet"
'Exec_Flag = "Y"	
'If Exec_Flag = "Y" Then
'
'Set fileSystemObj = createobject("Scripting.FileSystemObject")
'
'		'To check if the given file present'
'		
'		MyFile = "C:\automation\APPLICATIONS\FFWF\Data\FFWF_Data.xlsx"
'
'		If fileSystemObj.FileExists(MyFile) then
'		    rc = "True"
'		Else
'		    rc = "False"
'		End If
'
'		If rc = "True" Then			
'			Set oExcel=CreateObject("Excel.Application")
'			Set oBook=oExcel.Workbooks.Open("C:\automation\APPLICATIONS\FFWF\Data\FFWF_Data.xlsx")
'			Set oSheet=oBook.Worksheets("Test_Data")
'
'			rows=oSheet.UsedRange.rows.count
'			Cols=oSheet.UsedRange.Columns.count
'
'			Event_Name=oSheet.Cells(i,6).value
'			msgbox Event_Name
'		   ' Exml=oSheet.Cells(i,7).value
'		   'oBook.Save
'oBook.Close
'oExcel.Quit
'    
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Highlight
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").click
'				Set ws=CreateObject("wscript.shell")
'		      	ws.SendKeys ("^a")	
'			    ws.SendKeys "{DELETE}"
'			    'ws.SendKeys Event_Name
'				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Type Event_Name	    
'			    Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=1").Type Event_Name	
'			    Set ws=Nothing 
'		
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
'Set fso=createobject("Scripting.FileSystemObject")
'
'Set qfile=fso.OpenTextFile("C:\automation\APPLICATIONS\FFWF\XML\ENJ_NewInstall.xml",1)
'Exml=qfile.ReadAll
'msgbox Exml
'Set qfile=nothing
'Set fso=nothing
'
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").Highlight
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").click
'				Set ws=CreateObject("wscript.shell")
'		      	ws.SendKeys ("^a")	
'			    ws.SendKeys "{DELETE}"
'			    'ws.SendKeys Exml
'			    'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinEditor("nativeclass:=Edit","location:=2").click
'				'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinEditor("nativeclass:=Edit","location:=2").Type micCtrlDwn + "a" + micCtrlUp
'				'ws.SendKeys "{BACKSPACE}"
'		
'				wait (2)
'				
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","location:=2").Type Exml
'	     		Set ws=Nothing 
'		 
'				wait (2)
'
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinToolbar("nativeclass:=ToolbarWindow32","location:=8").Press 1
'				
'				wait (2)
'					If Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("text:=Save Resource").WinButton("text:=&Yes").Exist(5) Then
'					Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("text:=Save Resource").WinButton("text:=&Yes").Click
'					End if 
'
'					 	Actual_Res = "Service Name and Request Details Updated as per the testcase" 
'						Reporter.ReportEvent micInfo, StepName, Actual_Res
'		Else
'				Call captureScreen
'				Actual_Res = "Service Name and Request Details NOT Updated as per the testcase" 
'				Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'				Installorder_xml= -1 
'			
'		End if
'
'				if Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("attached text:=Finished","regexpwndtitle:=1/1 Tests Succeeded").Exist Then
'				   Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("attached text:=Finished","regexpwndtitle:=1/1 Tests Succeeded").Highlight
'		 				Actual_Res = "Tests Succeeded static text is validated" 
'						Reporter.ReportEvent micInfo, StepName, Actual_Res
'				Else
'						Call captureScreen
'						Actual_Res = "Tests Succeeded static text is Not Displayed" 
'						Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'						Installorder_xml= -1 
'				End if
'	
'				if Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("nativeclass:=Static","regexpwndtitle:=No Tasks Reported").Exist Then
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Static("nativeclass:=Static","regexpwndtitle:=No Tasks Reported").Highlight
'						 	Actual_Res = "No Tasks Reported static text is validated" 
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'				Else
'						Call captureScreen
'						Actual_Res = "No Tasks Reported static text is Not Displayed" 
'						Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'						Installorder_xml= -1 
'				End if
'				
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest.*","Location:=0").Highlight
'		Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest.*","Location:=0").Close
'wait(2)
'		If Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("regexpwndtitle:=Confirm Exit","text:=Confirm Exit").WinButton("text:=OK").Exist(5) Then
'					Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").Dialog("regexpwndtitle:=Confirm Exit","text:=Confirm Exit").WinButton("text:=OK").Click
'					End if 
'wait(2)
'
'
'
'	End if	    





'End Function	



'
'If instr(1,val,"REKHA",0) Then
'	msgbox "pass"
'	else
'	msgbox "fail"
'End If

'Public Function Buslistnerorder_xml(inputdata)
'	On Error Resume Next
'	Dim methodName, rc, Temp_Name
'	methodName = "Buslistnerorder_xml" : Buslistnerorder_xml = 0
'	Execute("Test_URL = " & TEST_ENV)


	'To generate Test Step Description and Expected Result
'	Step_Description = "Load XML from the Test Data Sheet" 
'	Exp_Result = " XML is loaded successfully from the Test Data Sheet"
'	Exec_Flag = "Y"
''If Exec_Flag = "Y" Then
'	'Call close all open browser function
''		Call closeAllBrowser
''
''		If TEST_BROWSER = "IE" Then 
''			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","",3
''		'ElseIf 	TEST_BROWSER = "Firefox" Then 
''			'SystemUtil.Run "firefox.exe",Test_URL,"C:\","", 3
''		'ElseIf 	TEST_BROWSER = "GChrome" Then 
''			'SystemUtil.Run "chrome.exe",Test_URL,"C:\","",3
''		Else 
''			TEST_BROWSER = "IE"		'	Default Browser
''			SystemUtil.Run "iexplore.exe",Test_URL,"C:\","", 3
''		End If
'
'
'inputdata="Resp.xml"
'			If Browser("name:=TMS Bus Listener","title:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("html id:=tbRequest","html tag:=TEXTAREA","name:=tbRequest").Exist(2) Then
'		    rc = "True"
'		Else
'		    rc = "False"
'		End If
'	
'		If rc = "True" Then			
'			
'			Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'			xmlDoc.Async = False 
'			
'			Set fso=createobject("Scripting.FileSystemObject")
'			
''			Set qfile=fso.OpenTextFile("C:\automation\APPLICATIONS\FFWF\XML\AVS_Order_Fulfillment.xml",1)
''			Responsexml=qfile.ReadAll
'			
'			filename = "C:\automation\APPLICATIONS\FFWF\XML\"&inputdata
'			msgbox filename
'			Set qfile=fso.OpenTextFile(filename,1)
'			Responsexml=qfile.ReadAll
'			
'			Set qfile=nothing
'			Set fso=nothing
'			
'			Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Highlight
'			Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Click
'				Set ws=CreateObject("wscript.shell")
'		      	ws.SendKeys ("^a")	
'			    ws.SendKeys "{DELETE}"
'			    
'		
'				wait (2)
'				
'				Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Set Responsexml
'	     		Set ws=Nothing 
'			End if
'
'wait(2)
'
''oBook.Save
'oBook.Close
'oExcel.Quit

'End if

' Handling Error
'	methodName = "Buslistnerorder_xml" : rc = ErrorHandler(methodName)
'End Function

''
''
'
'
'
'inputdata="Resp.xml"
''	If Browser("name:=TMS Bus Listener","title:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("html id:=tbRequest","html tag:=TEXTAREA","name:=tbRequest").Exist(2) Then
''		    rc = "True"
''		    msgbox rc
''		Else
''		    rc = "False"
''		    msgbox rc
''		End If
''	
''		If rc = "True" Then			
'			
'			Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'			xmlDoc.Async = False 
'			
'			Set fso=createobject("Scripting.FileSystemObject")
'			
'			Set qfile=fso.OpenTextFile("C:\automation\APPLICATIONS\FFWF\XML\AVS_Order_Fulfillment.xml",1)
'			filename = "C:\automation\APPLICATIONS\FFWF\XML\"&inputdata
'			msgbox filename
'			Set qfile=fso.OpenTextFile(filename,1)
'			Responsexml=qfile.ReadAll
'			msgbox Responsexml
'			
'			Set qfile=nothing
'			Set fso=nothing
'			
'			Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Highlight
'			Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Click
'				Set ws=CreateObject("wscript.shell")
'		      	ws.SendKeys ("^a")	
'			    ws.SendKeys "{DELETE}"
'			    
'				wait (2)
'				
'				'ws.SendKeys Responsexml
'				
'				Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Set Responsexml
'	     		Set ws=Nothing 
'			'End if
'
'wait(2)

'oBook.Save
'oBook.Close
'oExcel.Quit

'Set fso=createobject("Scripting.FileSystemObject")
'
''Set qfile=fso.OpenTextFile("C:\Automation\APPLICATIONS\FFWF\XML\Validate_Bus_Listener.txt",2,true)
'
'fso.CopyFile "C:\Automation\APPLICATIONS\FFWF\XML\Validate_Bus_Listener.txt","C:\Automation\APPLICATIONS\FFWF\XML\Validate_Bus_Listener.xml"
'
''qfile.writeline val2
''
''msgbox val2
''
''qfile.Close
'
'Set qfile=nothing
'Set fso=nothing


'Public Function ValidateXMLTags(inputData,inputData1,inputData2,inputData3)
'On Error Resume Next

'inputData1="bim:ErrorCode"
'inputData2="ErrorCode"
'inputData3=0
'Dim methodName,rc,val,xml_val,obj,xml_cnt,obj1,obj_xml,i,j,N,k,m,strText,xmlDoc,XMLDataFile,nodelist,session_val,rxid,RxSessionId,mysheet,Row,Col,myxl,item
'
'	'methodName = "ValidateXMLTags" : ValidateXMLTags = 0
'	
'	Exec_Flag = "Y" 
'	'If Exec_Flag = "Y" Then
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
''path of XML file
'
'XMLDataFile="C:\Automation\APPLICATIONS\FFWF\XML\Validate_Bus_Listener.xml"
'
''Load the XML File
' xmlDoc.Load(XMLDataFile)
''get the tagname to verify
'Set nodelist = xmlDoc.getElementsByTagName(inputData1)
''msgbox nodelist.length
'val =  nodelist.length
'msgbox val
'
'
'strText = nodelist.item(0).xml
'
'msgbox strText
'
''For m = 0 to val-1
''
''       strText = nodelist.item(m).xml
''       'strText = nodelist(i).nodevalue
''       msgbox strText
''       tag_val = split(strText,">")
''       tag = split(tag_val(1),"<")
''       REQUEST_ID =  tag(0)
''       
''		msgbox REQUEST_ID
''
''       If abs(m) = abs(inputData3) Then
''	    Exit For
''       End If
''
'''       If instr(trim(ucase(strText)),"GPONFTTP")<>0 Then
'''       	      msgbox "success" 
''       'End If
''Next
'
'If strText<>"" Then
'	rc = 0
'	Step_Description = "To Collect XML Value"
'	Exp_Result = "Xml value should be collected successfully"
'	Actual_Res = "XML value collected is->"& strText 
'	Reporter.ReportEvent micInfo, StepName, Actual_Res
'	
'	Else
'	
'	Step_Description = "To Collect XML Value"
'	Exp_Result = "Xml value should be collected successfully"
'	Actual_Res = "XML value collected is not collected successfully for->"& inputData1
'	Reporter.ReportEvent micFail, StepName, Actual_Res
''	ValidateXMLTags =-1
'End If
'
'
'Set myxl = createobject("excel.application")
'
'myxl.Workbooks.Open "C:\automation\APPLICATIONS\FFWF\DATA\FFWF_Data.xlsx" 
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
'
'End if
''Browser("micclass:=Browser","name:=TMS Bus Listener.*").Close
'
'' Handling Error
''	methodName = "ValidateXMLTags" : rc = ErrorHandler(methodName)
''End Function
'
''rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Exist
''msgbox rc
''
''If rc= "True" Then
''
''JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Highlight
''JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Press(4)
'
'rc1 = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Exist
'msgbox rc1
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Highlight
'
'
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Press(4)
'							wait(5)
'							
'rc2 = JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(10) 
'msgbox rc2
'
'					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(1) Then
'					wait (3)
'							Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							Call captureScreen
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Validatecomplete_MasterRequestId=-1
'					End If
'					
'					
'			inputData=	"100001542"
'			master_reqid = "100001542"
'			status_code = "Completed"
'			inputData1 = "Completed"
'			
'		JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").highlight
'		rc = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(10)
'
'			If rc = "True" Then
'	
'				row_cnt = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetROProperty("rows")
'				For i = 0 To row_cnt-1
'					master_reqid = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetCellData(i,1)
'					'msgbox process_name
'						If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
'						'msgbox "found" &inputData
'							'get the value of the status code
'						
'							status_code = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetCellData(i,4)
'									If Trim(Ucase(status_code)) = Trim(ucase(inputData1)) Then
'																					msgbox "found" &status_code
'																					
'																					JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ClickCell i,2
'																					
'																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ClickCell i,1
'																				
'																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ActivateRow i
'																				
'																				'JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").DoubleClickCell i,1
'																				
'																				
'																				Flag = True
'																					else
'																					'Call captureScreen
'																					Step_Description = "The Status of the selected Master Request id is not in progress"
'																					Exp_Result = "The Status of the selected Master Request id is in progress"
'																					Actual_Res = "The Status of the selected Master Request id is not in progress"
'																					Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'																					Validatecomplete_MasterRequestId=-1
'																					End if
'															Step_Description = "To select the required Master Request id"
'															Exp_Result = "Master Request Id should be displayed"
'															Actual_Res = "Master Request Id is displayed"
'															Reporter.ReportEvent micInfo, StepName, Actual_Res
'								Exit For						
'								Else
'															'Call captureScreen
''															Step_Description = "To select the required Master Request id"
''															Exp_Result = "Master Request Id should be displayed"
''															Actual_Res = "Master Request Id is not displayed"
''															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
''															Loading_MasterRequestId=-1
''												
'														End If
'						Next
'								
'								Do
'								JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Highlight
'								display = JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Exist
'								Loop While display = False
'					
'														If display = "True" Then
'															Step_Description = "To Verify if the process for the selected master request id is loaded"
'															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
'															Actual_Res = "Process is Loaded successfully for given Master Request Id" &inputData
'															Reporter.ReportEvent micInfo, StepName, Actual_Res
'														Else
'															'Call captureScreen
'															Step_Description = "To Verify if the process for the selected master request id is loaded"
'															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
'															Actual_Res = "Process is not Loaded successfully for given Master Request Id" &inputData
'															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'															Validatecomplete_MasterRequestId=-1	
'														End If
'			End if
'			
''End If @@ hightlight id_;_685929527_;_script infofile_;_ZIP::ssf179.xml_;_


'Public Function ReadXMLFileAndReplaceXMLTags(Queryid,FileLocation,Tagtobereplaced,TagValue)
'On Error Resume Next

'Sheetname = "Query"
'FilePath="C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx"
'Datatable.AddSheet Sheetname
'Datatable.ImportSheet FilePath, Sheetname, Sheetname
'verify_First_Data_from_Datatable= Datatable.Value("Query",Sheetname)
'Print verify_First_Data_from_Datatable
'
'getRowCount=Datatable.GetSheet(Sheetname).GetRowCount
'getParamCount=Datatable.GetSheet(Sheetname).GetParameterCount
'Print "getParamCount ="&getParamCount
'
'Print "Total Rows ="&getRowCount
'Datatable.GetSheet(Sheetname).SetCurrentRow(1)
'For Iterator = 1 To getRowCount Step 1
'	Print "Iterator ="&Iterator
'	Query_ID="Q"&Iterator
'	Print "Query_ID ="&Query_ID
'	FileLocation=Datatable.GetSheet(Sheetname).GetParameter("FileLocation").Value
'	Print "FileLocation ="&FileLocation
'	Tagtobereplaced=Datatable.GetSheet(Sheetname).GetParameter("Tagtobereplaced").Value
'	Print "Tagtobereplaced ="&Tagtobereplaced
'	TagValue=Datatable.GetSheet(Sheetname).GetParameter("TagValue").Value
'	Print "TagValue ="&TagValue
'	
'	Dim methodName,rc,strText,xmlDoc,XMLDataFile,nodelist
'
'	'methodName = "ReadXMLFileAndReplaceXMLTags" : ReadXMLFileAndReplaceXMLTags = 0
'	Exec_Flag = "Y"
'If Exec_Flag = "Y" Then
'
''QID=Queryid+1
''msgbox QID
'
''Set myxl = createobject("excel.application")
''
''myxl.Workbooks.Open "C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx" 
''
''set mysheet = myxl.ActiveWorkbook.Worksheets("Query")
''
''Row=mysheet.UsedRange.rows.count
''
''FileLocation = mysheet.cells(QID,1).value
''msgbox QID
''FileLocation = mysheet.cells(QID,5).value
''msgbox FileLocation
''Tagtobereplaced = mysheet.cells(QID,6).value
''msgbox Tagtobereplaced
''TagValue = mysheet.cells(QID,4).value
''msgbox TagValue
''
''myxl.Workbooks.Close
''myxl.Quit
''
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
''path of XML file
'
'XMLDataFile=FileLocation
'
''Load the XML File
' xmlDoc.Load(XMLDataFile)
''get the tagname to verify
'Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)
'
'nodelist.text = TagValue
'
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
'
'End if
'Datatable.SetNextRow
'Next
'
'ExitTest
'	 'Handling Error
'	methodName = "ReadXMLFileAndReplaceXMLTags" : rc = ErrorHandler(methodName)
'End Function

			' JavaDialog("title:=Security Warning").SetFocus
		     'JavaDialog("title:=Security Warning").Highlight
                          
'		 If JavaDialog("title:=Security Warning").JavaCheckBox("attached text:=I accept the risk and.*").Exist(10) Then
'		 	JavaDialog("title:=Security Warning").JavaCheckBox("attached text:=I accept the risk and.*").Highlight
'			JavaDialog("title:=Security Warning").JavaCheckBox("attached text:=I accept the risk and.*").Set "ON"
'		    JavaDialog("title:=Security Warning").JavaButton("attached text:=Run").Click
'
'	 		Else
'	 		MDWDesigner_Login=-1
'	 	End If
'	
'	rc =JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").Exist(60) 
'
'		If rc ="True" Then
'
'			JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").JavaEdit("attached text:=User Name").Highlight
'			JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").JavaEdit("attached text:=User Name").set inputData
'
'			JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").JavaEdit("attached text:=Password").SetFocus
'			wait(1)
'			JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").JavaEdit("attached text:=Password").Set inputData1
'					wait(1)
'				if JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").JavaButton("attached text:=Log In").Getroproperty("enabled") = "1" Then
'					JavaWindow("title:=MDW Designer","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").JavaButton("label:=Log In").Click
'				End if
'
'			Do
'				display = JavaWindow("title:=MDW Designer.*","toolkit class:=com\.qwest\.mdw\.designer\.MainFrame").Exist
'
'			Loop While display = False
'
'			if display = True Then
'					Step_Description = "Login to MDW Designer in test environment"
'					Exp_Result = "Login to MDW Designer should be successfull"
'					Actual_Res = "Login to MDW Designer is successfull"
'					Reporter.ReportEvent micInfo, StepName, Actual_Res
'			Else
'			Step_Description = "Login to MDW Designer in test environment"
'			Exp_Result = "Login to MDW Designer should be successfull"
'			Actual_Res = "Login to MDW Designer failed"
'			Reporter.ReportEvent micFail, StepName, Actual_Res
'			MDWDesigner_Login=-1
'			End If
'	
'		Else
'			Step_Description = "Login window display"
'			Exp_Result = "Login window should be displayed"
'			Actual_Res = "Login window is displayed"
'			Reporter.ReportEvent micFail, StepName, Actual_Res
'			MDWDesigner_Login=-1



'	Dim methodName,rc, i, master_reqid, row_cnt
'	Exec_Flag = "Y"
'	
'	If Exec_Flag = "Y" Then
'
'rc = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Exist
''msgbox rc
'
'If rc= "True" Then
'
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Highlight
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstancePage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=JToolBar").Press(4)
'
'rc1 = JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Exist
''msgbox rc1
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Highlight
'
'
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=JToolBar;ProcessInstanceListPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=Page 1 of 9").Press(4)
'							wait(5)
'							
'rc2 = JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(10) 
''msgbox rc2
'
'					If JavaWindow("title:=MDW Designer.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(1) Then
'					wait (3)
'							Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micPass, StepName, Actual_Res
'						Else
'							Call captureScreen
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Validatecomplete_MasterRequestId=-1
'					End If
'					
'					
'			
'		JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").highlight
'		rc = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").Exist(10)
'
'			If rc = "True" Then
'	
'				row_cnt = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetROProperty("rows")
'				For i = 0 To row_cnt-1
'					master_reqid = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetCellData(i,1)
'					'msgbox process_name
'						If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
'						print "Entered master request is found" &inputData
'							'get the value of the status code
'						
'							status_code = JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").GetCellData(i,4)
'									If Trim(Ucase(status_code)) = Trim(ucase(inputData1)) Then
'																					print "Required status code is found" &status_code
'																					
'																					JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ClickCell i,2
'																					
'																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ClickCell i,1
'																				
'																				JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").ActivateRow i
'																				
'																				'JavaWindow("title:=MDW Designer.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax.swing.JTable").DoubleClickCell i,1
'																				
'																				
'																				Flag = True
'																					else
'																					'Call captureScreen
'																					Step_Description = "The Status of the selected Master Request id is not in progress"
'																					Exp_Result = "The Status of the selected Master Request id is in progress"
'																					Actual_Res = "The Status of the selected Master Request id is not in progress"
'																					Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'																					Validatecomplete_MasterRequestId=-1
'																					End if
'															Step_Description = "To select the required Master Request id"
'															Exp_Result = "Master Request Id should be displayed"
'															Actual_Res = "Master Request Id is displayed"
'															Reporter.ReportEvent micPass, StepName, Actual_Res
'								Exit For						
'								Else
'															'Call captureScreen
''															Step_Description = "To select the required Master Request id"
''															Exp_Result = "Master Request Id should be displayed"
''															Actual_Res = "Master Request Id is not displayed"
''															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
''															Loading_MasterRequestId=-1
''												
'														End If
'						Next
'								Flag = True
'															Step_Description = "To Verify if the process for the selected master request id is loaded"
'															Exp_Result = "Process Should be Loaded successfully for given Master Request Id" 
'															Actual_Res = "Process is Loaded successfully for given Master Request Id" &inputData
'															Reporter.ReportEvent micPass, StepName, Actual_Res
'														
'			End if
'			End if
'End If
'
'wait(5)
'JavaWindow("title:=MDW Designer.*").Close
'
'row=17
'col=4
'inputdata="013304467370"
'Flag = True
'                                                                                
'                                                      'msgbox row
'                                                      'msgbox col
'                                                      'msgbox inputdata
'                                                      'On Error Resume Next
'                                                      'i=1
'                                                      'rc=trim(inputData)
'                                                      'r = split(rc,",")
'                                                      'row = r(0) 
'                                                      'column =  r(1)-1
'                                                 
'                                                      	
'                                                      
'                                                      			  Set obj=Description.Create()
''                                                                 obj("Class Name").Value="JavaToolbar"
''                                                                 obj("attached text").Value="Product Home"
''                                                                 Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
''                                                                     For i = 0 To obj1.Count-1
''                                                                     obj1(i).Highlight
''                                                                     obj1(i).Press "Copy RX# to clipboard"
'                                                                      Dim objCB,str_1,xl
'                                                                      'Set objCB= CreateObject("Mercury.Clipboard")
'                                                                      'str_1 = objCB.GetText
'                                                                      'str_1 = FF_Requestid_VAL
'                                                                       set xl=CreateObject("Excel.Application") 
'                                                                       set workbook= xl.Workbooks.Open("C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx")  
'                                                                       set sheet=workbook.Sheets("Query")
'                                                                       inputdata1="abcd"&inputdata
'                                                                       sheet.Cells(row,col).value=inputdata1
'                                                                       'sheet.Cells(row,column).interior.colorindex=6
'                                                                        xl.ActiveWorkbook.Save
'                                                                        'xl.ActiveWorkbook.Close
'                                                                        'set xl.Application=Nothing
'                                                                        'Set xl=Nothing
'                                                                         xl.Workbooks.Close
'																		 xl.Quit
'																		 
'																		 
'																		 
'																		 
'
'
' xl.Workbooks.Close
' xl.Quit

	
'inputData2 = "AVS_Order_Fulfillment"
'		
'	Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebList("name:=mainHelperForm:detailForm:eventNameSelect").Select "SERVICE_ORDER_EVENT_BUNDLING"	
'	
'				Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'			xmlDoc.Async = False 
'			
'			Set fso=createobject("Scripting.FileSystemObject")
'			
'			'Set qfile=fso.OpenTextFile("C:\Automation\APPLICATIONS\FFWF\XML\AVS_Order_Fulfillment.xml",1)
'			Set qfile=fso.OpenTextFile("C:\Automation\APPLICATIONS\FFWF\XML\"&inputData2&".xml",1)
'			sResponsexml=qfile.ReadAll
'			print "Response xml is"&sResponsexml
'			
'			Set qfile=nothing
'			Set fso=nothing
'			
'			Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:eventMessageTextarea").Highlight
'			
'	
'			wait (2)
'Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:eventMessageTextarea").Click
'				Set ws=CreateObject("wscript.shell")
'		      	ws.SendKeys ("^a")	
'			    ws.SendKeys "{DELETE}"
'			    'ws.SendKeys sResponsexml
'			    
'				wait (2)
'		
'Browser("name:=MDW Web.*","title:=MDW Web.*").Page("title:=MDW Web.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=mainHelperForm:detailForm:eventMessageTextarea").Set sResponsexml
'	 Actual_Res = "XML is loaded successfully from the XML file"
'	Reporter.reportevent micPass, StepName, Actual_Res
'	     		Set ws=Nothing 
	     		

'Query_ID = "VDP_RequestId"
''FileLocation,Tagtobereplaced,TagValue
'
'
'Sheetname = "Query"
'FilePath="C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx"
'Datatable.AddSheet Sheetname
'Datatable.ImportSheet FilePath, Sheetname, Sheetname
'verify_First_Data_from_Datatable= Datatable.Value("Query",Sheetname)
''Print verify_First_Data_from_Datatable
'
'getRowCount=Datatable.GetSheet(Sheetname).GetRowCount
'getParamCount=Datatable.GetSheet(Sheetname).GetParameterCount
'Print "getParamCount ="&getParamCount
'
'Print "Total Rows ="&getRowCount
'Datatable.GetSheet(Sheetname).SetCurrentRow(1)
'For Iterator = 1 To getRowCount Step 1
'
'	Print "Iterator ="&Iterator
'	'Query_ID="Q"&Iterator
'	QID=Datatable.GetSheet(Sheetname).GetParameter("QID").Value
'	Print "QID ="&QID
'	 
'	 If QID=Query_ID Then	
'	 
'	FileLocation=Datatable.GetSheet(Sheetname).GetParameter("FileLocation").Value
'	Print "FileLocation ="&FileLocation
'	Tagtobereplaced=Datatable.GetSheet(Sheetname).GetParameter("Tagtobereplaced").Value
'	Print "Tagtobereplaced ="&Tagtobereplaced
'	TagValue=Datatable.GetSheet(Sheetname).GetParameter("TagValue").Value
'	Print "TagValue ="&TagValue
'	Else
'
'End If
'	
'	
'	Dim methodName,rc,strText,xmlDoc,XMLDataFile,nodelist
'
'	methodName = "ReadXMLFileAndReplaceXMLTags" : ReadXMLFileAndReplaceXMLTags = 0
'	
'Exec_Flag = "Y"
'rc = "True"
'	
'	If Exec_Flag = "Y" Then
'	
'		Set fileSystemObj = createobject("Scripting.FileSystemObject")
'		MyFile = "C:\automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx"
'	
'			If fileSystemObj.FileExists(MyFile) then
'				rc = "True"
'			Else
'				rc = "False"
'			End If
'	
'			If rc = "True" Then		
'			
'				'QID=Queryid+1
'				'msgbox QID
'				'Set myxl = createobject("excel.application")
'				'
'				'myxl.Workbooks.Open "C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx" 
'				'
'				'set mysheet = myxl.ActiveWorkbook.Worksheets("Query")
'				'
'				'Row=mysheet.UsedRange.rows.count
'				'
'				'FileLocation = mysheet.cells(QID,1).value
'				'msgbox QID
'				'FileLocation = mysheet.cells(QID,5).value
'				'msgbox FileLocation
'				'Tagtobereplaced = mysheet.cells(QID,6).value
'				'msgbox Tagtobereplaced
'				'TagValue = mysheet.cells(QID,4).value
'				'msgbox TagValue
'				'
'				'myxl.Workbooks.Close
'				'myxl.Quit
'				'
'				Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'				xmlDoc.Async = False 
'				
'				'path of XML file
'				
'				XMLDataFile=FileLocation
'				
'				'Load the XML File
'				xmlDoc.Load(XMLDataFile)
'				'get the tagname to verify
'				Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)
'				
'				nodelist.text = TagValue
'				
'				xmlDoc.Save(XMLDataFile)
'				
'				Set xmlDoc = nothing
'			
'			Else
'				Reporter.ReportEvent micFail, "ReadXMLFileAndReplaceXMLTags - Function failed"," rc = false - Please check"
'			End if
'			Else
'				Reporter.ReportEvent micFail, "ReadXMLFileAndReplaceXMLTags - Function failed"," rc = false - Please check"
'			End if
'
'Datatable.SetNextRow
'Step_Description = "XML tag values are updated for the Query id -> " & Query_ID
'Exp_Result = Tagtobereplaced & " -> is updated with the tag value as " & TagValue &" for the Query id " & Query_ID
'Actual_Res = "Tagvalue is updated successfully"
'Reporter.reportevent micPass,StepName, Actual_Res
'
'
'Next



' inputData = "Resp"
' rc = "True"

''MyFile = ("C:\Automation\APPLICATIONS\FFWF\XML\"&inputData&".xml")
''
''		If fileSystemObj.FileExists(MyFile) then
''		    rc = "True"
''		Else
''		    rc = "False"
''		End If
'
'		If rc = "True" Then
'		
'FileLocation = ("C:\Automation\APPLICATIONS\FFWF\XML\"&inputData&".xml")
'print "Filelocations is:" &FileLocation
'	
'sdate= Year(date)&"-"&month(date)&"-"&day(date)
'print "Today Date is:" &sdate
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
'XMLDataFile=FileLocation
'
'xmlDoc.Load(XMLDataFile)
'
'
'Select Case inputData
'Case "AVS_Order_Fulfillment"
'Set strDateNode = xmlDoc.SelectSingleNode("ENJEventMessageRequest/bim:SendTimeStamp")
'
'Case "DvarOm_Modem_Cancel_Return"
'Set strDateNode = xmlDoc.SelectSingleNode("ReturnSTBShipmentRequest/RequestHeader/SendTimeStamp")
'
'Case "DvarOm_Modem_Cancel_Return_Response"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
'
'Case "DvarOm_Modem_Cancel_Return_Response_success"
'Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentReturnResponse/ReturnResponseHeader/qb:SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentReturnResponse/ReturnAuthorization/ReturnAuthorizationExpirationDate")
'
'Case "DVAROM_Order_Fulfillment"
'Set strDateNode = xmlDoc.SelectSingleNode("ENJEventMessageRequest/bim:SendTimeStamp")
'
'Case "ENJ_Order_Fulfillment"
'Set strDateNode = xmlDoc.SelectSingleNode("ENJEventMessageRequest/bim:SendTimeStamp")
'
'Case "FF_MASTER_PROCESS_REDESIGN_FF_Response"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/SendTimeStamp")
'
'Case "FF_MASTER_PROCESS_REDESIGN_FF_Shipment"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ShippingDate")
'
'Case "PURETV_VENDOR_DELIVERY_PROCESS_REQUEST"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentRequest/ns2:ShipmentRequestHeader/SendTimeStamp")
'
'Case "PURETV_VENDOR_DELIVERY_PROCESS_RESPONSE"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ShippingDate")
'
'Case "PURETV_VENDOR_INITIATED_RETURNS_REQUEST"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnRequestHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")
'
'Case "PURETV_VENDOR_INITIATED_RETURNS_RESPONSE"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnDate")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")
'
'Case "Resp"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/SendTimeStamp")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/ns2:OrderReceivedDate")
'strDateNode.text = sdate
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ExpectedShippingDate")
'
'Case "Ship"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/ns2:OrderReceivedDate")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:EstimatedDeliveryDate")
'
'Case "VENDOR_DELIVERY_PROCESS_ENS_REQUEST"
'Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentShipmentRequest/ShipmentRequestHeader/qb:SendTimeStamp")
'
'Case "VENDOR_DELIVERY_PROCESS_ENS_RESPONSE"
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/qb:SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/ns:OrderReceivedDate")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:EstimatedDeliveryDate")
'
'Case "VENDOR_DELIVERY_PROCESS_IOM_REQUEST"
'Set strDateNode = xmlDoc.SelectSingleNode("FulfillmentShipmentRequest/ShipmentRequestHeader/qb:SendTimeStamp")
'
'Case "VENDOR_DELIVERY_PROCESS_IOM_RESPONSE"
'Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/qb:SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponseHeader/ns:OrderReceivedDate")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponse/qb:ShippingDate")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:EstimatedDeliveryDate")
'
'Case "VENDOR_INITIATED_RETURNS_ENS_REQUEST"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnRequestHeader/SendTimeStamp")
'
'Case "VENDOR_INITIATED_RETURNS_ENS_RESPONSE"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnDate")
''strDateNode.text = sdate
'
'Case "VENDOR_INITIATED_RETURNS_IOM_REQUEST"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnRequestHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnRequest/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")
'
'Case "VENDOR_INITIATED_RETURNS_IOM_RESPONSE"
'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnResponseHeader/SendTimeStamp")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnDate")
''strDateNode.text = sdate
''Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentReturnResponse/ns2:ReturnAuthorization/ns2:ReturnAuthorizationExpirationDate")
'
'End Select
'  
'  
''get the tagname to verify
''Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)
'
'		
'strDateNode.text = sdate
'
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
'Set FileLocation = nothing
'	Actual_Res = "Today's Date is updated successfully in the XML response file"
'	Reporter.ReportEvent micPass, StepName, Actual_Res
'

'
'inputData="PURETV_VENDOR_DELIVERY_PROCESS_RESPONSE"
'
'
'
'MyFile = ("C:\Automation\APPLICATIONS\FFWF\XML\"&inputData&".xml")
'
'		'If fileSystemObj.FileExists(MyFile) then
'		    rc = "True"
'		'Else
'		'    rc = "False"
'		'End If
'
'		If rc = "True" Then
'		
'	
'FileLocation = ("C:\Automation\APPLICATIONS\FFWF\XML\"&inputData&".xml")
'print "Filelocations is:" &FileLocation
'	
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
'XMLDataFile=FileLocation
'
'xmlDoc.Load(XMLDataFile)
'
'If (inputData = "Ship")  or (inputData ="FF_MASTER_PROCESS_REDESIGN_FF_Shipment" ) Then
'
'		Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
'		strText = nodelist.item(0).text
'				
'			print "The TrackingNumber number is "&strText
'				
'			str1 = left(strText,14)
'			print "First 14 digits of the TrackingNumber is "&str1
'			
'			remval = right(strText,4)
'			print "Last 4 digits of the TrackingNumber is "&remval
'			
'			str2 = remval+10
'			print "Updating the last 3 digits as "&str2
'			
'			strTextnew = str1&str2
'			print "The new last 4 digits of the TrackingNumber is "&strTextnew
'			
'			Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingNumber")
'			'msgbox strTrNum.text
'			strTrNum.text = strTextnew
'			
'			Actual_Res = "Updated TrackingNumber in the XML file" 
'			Reporter.ReportEvent micPass, StepName, Actual_Res
'		 
'			 
'			'To update the TrackingURL tag in the Ship.xml response file
'			Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
'			strText = nodelist.item(0).text
'					
'			print "The TrackingURL is"&strText
'					
'			val1 = split(strText,"trackNums=")
'			print "The Tracking Number in the TrackingURL is "&val1(1)
'			
'			val2 = Replace(val1(1),remval,str2)
'			print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
'			
'			strTextnew = Replace(strText,val1(1),val2)
'			print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
'				
'		
'			Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingURL")
'			'msgbox strTrNumUrl.text
'			strTrNumUrl.text = strTextnew
'
'ElseIf inputData = "VENDOR_DELIVERY_PROCESS_IOM_RESPONSE"  Then
'
'Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
'strText = nodelist.item(0).text
'		
'	print "The TrackingNumber number is "&strText
'		
'	str1 = left(strText,14)
'	print "First 14 digits of the TrackingNumber is "&str1
'	
'	remval = right(strText,4)
'	print "Last 4 digits of the TrackingNumber is "&remval
'	
'	str2 = remval+10
'	print "Updating the last 3 digits as "&str2
'	
'	strTextnew = str1&str2
'	print "The new last 4 digits of the TrackingNumber is "&strTextnew
'	
'	Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:TrackingInfo/qb:TrackingNumber")
'	'msgbox strTrNum.text
'	strTrNum.text = strTextnew
'	
'	Actual_Res = "Updated TrackingNumber in the XML file" 
'	Reporter.ReportEvent micPass, StepName, Actual_Res
' 
'	 
'	'To update the TrackingURL tag in the Ship.xml response file
'	Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
'	strText = nodelist.item(0).text
'			
'	print "The TrackingURL is"&strText
'			
'	val1 = split(strText,"InquiryNumber1=")
'	print "The Tracking Number in the TrackingURL is "&val1(1)
'	
'	val2 = Replace(val1(1),remval,str2)
'	print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
'	
'	strTextnew = Replace(strText,val1(1),val2)
'	print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
'		
'
'	Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns:ShipmentResponse/ns:TrackingInfo/qb:TrackingURL")
'	'msgbox strTrNumUrl.text
'	strTrNumUrl.text = strTextnew
'
'
'ElseIf inputData ="PURETV_VENDOR_DELIVERY_PROCESS_RESPONSE"  Then
'
'Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
'strText = nodelist.item(0).text
'		
'	print "The TrackingNumber number is "&strText
'		
'	str1 = left(strText,14)
'	print "First 14 digits of the TrackingNumber is "&str1
'	
'	remval = right(strText,4)
'	print "Last 4 digits of the TrackingNumber is "&remval
'	
'	str2 = remval+10
'	print "Updating the last 3 digits as "&str2
'	
'	strTextnew = str1&str2
'	print "The new last 4 digits of the TrackingNumber is "&strTextnew
'	
'	Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingNumber")
'	'msgbox strTrNum.text
'	strTrNum.text = strTextnew
'	
'	Actual_Res = "Updated TrackingNumber in the XML file" 
'	Reporter.ReportEvent micPass, StepName, Actual_Res
' 
'	 
'	'To update the TrackingURL tag in the Ship.xml response file
'	Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
'	strText = nodelist.item(0).text
'			
'	print "The TrackingURL is"&strText
'			
'	val1 = split(strText,"InquiryNumber1=")
'	print "The Tracking Number in the TrackingURL is "&val1(1)
'	
'	val2 = Replace(val1(1),remval,str2)
'	print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
'	
'	strTextnew = Replace(strText,val1(1),val2)
'	print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
'		
'
'	Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingURL")
'	'msgbox strTrNumUrl.text
'	strTrNumUrl.text = strTextnew
'	
'End if
'
'
'	Actual_Res = "Updated TrackingNumber in the XML file as-> "& strTextnew 
'	Reporter.ReportEvent micPass, StepName, Actual_Res
'	
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
'End if

'Set fileSystemObj = createobject("Scripting.FileSystemObject")
'MyFile = fileSystemObj.FileExists("C:\automation\APPLICATIONS\FFWF\Data\DB_Details_All.xlsx")
'
'If MyFile then
' fileSystemObj.DeleteFile("C:\automation\APPLICATIONS\FFWF\Data\DB_Details_All.xlsx")
'		    
'		    
'		Else
'		Set ExcelObj=CreateObject("Excel.Application")
'		MySourceFile = fileSystemObj.FileExists("C:\automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx")
'					If  MySourceFile Then          
'						ExcelObj.Workbooks.Open("C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx") 
'						
'					    ExcelObj.ActiveWorkbook.SaveAs("C:\automation\APPLICATIONS\FFWF\Data\DB_Details_All.xlsx")
'
'					 Else                          
'					   
'					 End If
'				End If
'									      ExcelObj.Workbooks.Close
'							ExcelObj.Quit
'set fileSystemObj = nothing
'Set ExcelObj = nothing



'   row = 3
'   col = 12
'   Filename = "FFWF_Data"
'                                                      	
'                                                      
'          Set obj=Description.Create()
''         obj("Class Name").Value="JavaToolbar"
''         obj("attached text").Value="Product Home"
''         Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
''         For i = 0 To obj1.Count-1
''         obj1(i).Highlight
''         obj1(i).Press "Copy RX# to clipboard"
'          Dim objCB,str_1,xl
'          'Set objCB= CreateObject("Mercury.Clipboard")
'          'str_1 = objCB.GetText
'          'str_1 = FF_Requestid_VAL
'          set xl=CreateObject("Excel.Application") 
'		' set workbook= xl.Workbooks.Open("C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx")  
'		   set workbook= xl.Workbooks.Open("C:\Automation\APPLICATIONS\FFWF\Data\"&Filename&".xlsx",1) 
'		If Filename = "DB_Detail" Then
'			set sheet=workbook.Sheets("Query")
'			ElseIf Filename = "FFWF_Data" Then
'			set sheet=workbook.Sheets("Test_Data")
'		End If
'		
'																	   inputdata = sheet.cells(row,col).value
'																	   msgbox inputdata
'																	   
'																	  leng =  len(inputdata)
'																	   msgbox leng
'																	   														   
'																		str1 = right(inputdata,4)
'																		msgbox str1
'																		
'																		lengrem = leng-4
'																		msgbox lengrem
'																		
'																		remval = left(inputdata,lengrem)
'																		msgbox remval
'																		
'																		str2 = str1+10
'																		msgbox str2
'																		
'																		inputdata1 = remval&str2
'																		msgbox inputdata1
'			
'																		sheet.Cells(row,col).value=inputdata1
'			
'																																	
'                                                                       'sheet.Cells(row,column).interior.colorindex=6
'                                                                        xl.ActiveWorkbook.Save
'                                                                        'xl.ActiveWorkbook.Close
'                                                                        'set xl.Application=Nothing
'                                                                        'Set xl=Nothing
'                                                                         xl.Workbooks.Close
'																		 xl.Quit



'row =3
'col = 11
'Filename = "FFWF_Data"

'row = 15
'col = 3 
'Filename = "DB_Detail"
'                                                                                
'                                                      'msgbox row
'                                                      'msgbox col
'                                                      'msgbox inputdata
'                                                      'On Error Resume Next
'                                                      'i=1
'                                                      'rc=trim(inputData)
'                                                      'r = split(rc,",")
'                                                      'row = r(0) 
'                                                      'column =  r(1)-1
'                                                 
'                                                      	
'                                                      
'          Set obj=Description.Create()
''         obj("Class Name").Value="JavaToolbar"
''         obj("attached text").Value="Product Home"
''         Set obj1 = JavaWindow("title:=Rx - RX#.*").ChildObjects(obj)
''         For i = 0 To obj1.Count-1
''         obj1(i).Highlight
''         obj1(i).Press "Copy RX# to clipboard"
'          Dim objCB,str_1,xl
'          'Set objCB= CreateObject("Mercury.Clipboard")
'          'str_1 = objCB.GetText
'          'str_1 = FF_Requestid_VAL
'          set xl=CreateObject("Excel.Application") 
'		' set workbook= xl.Workbooks.Open("C:\Automation\APPLICATIONS\FFWF\Data\DB_Detail.xlsx")  
'		   set workbook= xl.Workbooks.Open("C:\Automation\APPLICATIONS\FFWF\Data\"&Filename&".xlsx",1) 
'		If Filename = "DB_Detail" Then
'			set sheet=workbook.Sheets("Query")
'			ElseIf Filename = "FFWF_Data" Then
'			set sheet=workbook.Sheets("Test_Data")
'		End If
'		
'	
'		inputdata = sheet.cells(row,col).value
'		msgbox inputdata
'	
'
'		leng =  len(inputdata)
'		msgbox leng
'		
'		str1 = right(inputdata,4)
'		msgbox str1
'		
'		str1lenght = len(str1)
'		msgbox str1lenght
'		
'		lengrem = leng - str1lenght
'		msgbox lengrem
'		
'		remval = left(inputdata,lengrem)
'		msgbox remval
'		
'		str2 = str1+10
'		msgbox str2
'		
'		inputdata1 = remval&str2
'		msgbox inputdata1
'		sheet.Cells(row,col).value=inputdata1
'		'sheet.Cells(row,column).interior.colorindex=6
'		 xl.ActiveWorkbook.Activate
'         xl.ActiveWorkbook.Save
'        'xl.ActiveWorkbook.Close
'        'set xl.Application=Nothing
'
'         xl.Workbooks.Close
'		 xl.Quit
'        Set xl=Nothing
'		'  NextJavaDialog("Security Warning_4").Activate
>>>>>>> 6f0c0c5309f46d67885a570fba43c8643d5f027d

'Window("regexpwndtitle:= Parasoft SOAtest").WinTreeView("regexpwndtitle:= SysTreeView32").HighLight
		
'Window("regexpwndtitle:= Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("object class:=SysTreeView32","Location:=0").HighLight @@ hightlight id_;_65826_;_script infofile_;_ZIP::ssf191.xml_;_
'Window("regexpwndtitle:= Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("object class:=SysTreeView32","Location:=0").Click
'Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Expand "FFWF"



 @@ hightlight id_;_612553465_;_script infofile_;_ZIP::ssf189.xml_;_
