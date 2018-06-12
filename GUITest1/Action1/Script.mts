'''RowNumber,Tag_Value,Scenario,Index
''
''
'''Public Function XMLValue_Collect(inputData,inputData1,inputData2,inputData3)
'''On Error Resume Next
'''
'''	Dim methodName,rc,val,xml_val,obj,xml_cnt,obj1,obj_xml,i,j,N,k,m,strText,xmlDoc,XMLDataFile,nodelist,session_val,rxid,RxSessionId,mysheet,Row,Col,myxl
'''
'''	methodName = "XMLValue_Collect" : XMLValue_Collect = 0
'''	i = inputData
''
''Exec_Flag = "Y"
''Dim rc,val,xml_val,obj,xml_cnt,obj1,obj_xml,i,j,N,k,m,strText,xmlDoc,XMLDataFile,nodelist,tag_val,tag,REQUEST_ID,mysheet,Row,Col,myxl
''
''If Exec_Flag = "Y" Then
'''Set val = Browser("creationtime:=2").Page("title:=.*").WebXML("html tag:=BODY").GetData
''Browser("name:=TMS Bus Listener.*").Page("title:=.*").WebEdit("name:=tbMessages").Highlight
''val = Browser("name:=TMS Bus Listener.*").Page("title:=.*").WebEdit("name:=tbMessages").GetROProperty("value")
''msgbox val
''
''xml_val = val
'''xml_val = val.ToString()
'''msgbox xml_val
''Set obj = createobject("scripting.filesystemobject")
''
''set xml_cnt = obj.OpenTextFile("C:\automation\APPLICATIONS\FFWF\XML\Xml_val.txt",2,true)
''
''
''xml_cnt.Writeline xml_val
''xml_cnt.Close
''
''Set xmlDoc = CreateObject("Microsoft.XMLDOM")
''xmlDoc.Async = False 
''
'''path of XML file
''
''XMLDataFile="C:\automation\APPLICATIONS\FFWF\XML\Xml_val.txt"
''
'''Load the XML File
'' xmlDoc.Load(XMLDataFile)
'''get the tagname to verify
''Set nodelist = xmlDoc.getElementsByTagName(REQUEST_ID)
'''msgbox nodelist.length
''val =  nodelist.length
''
''For m = 0 to val-1
''
''       strText = nodelist.item(m).xml
''       'strText = nodelist(i).nodevalue
''      ' msgbox strText
''       tag_val = split(strText,">")
''       tag = split(tag_val(1),"<")
''       REQUEST_ID =  tag(0)
''       
''       msgbox REQUEST_ID
''
''       If abs(m) = abs("Success") Then
''	    Exit For
''       End If
''
'''       If instr(trim(ucase(strText)),"GPONFTTP")<>0 Then
'''       	      msgbox "success" 
''       'End If
''Next
''
''If REQUEST_ID<>"" Then
''	rc = 0
''	Step_Description = "To Collect XML Value"
''	Exp_Result = "Xml value should be collected successfully"
''	Actual_Res = "XML value collected is->"& RxSessionId 
''	Reporter.ReportEvent micInfo, StepName, Actual_Res
''	
''	Else
''	
''	Step_Description = "To Collect XML Value"
''	Exp_Result = "Xml value should be collected successfully"
''	Actual_Res = "XML value collected is not collected successfully for->"& inputData1
''	Reporter.ReportEvent micFail, StepName, Actual_Res
''	XMLValue_Collect =-1
''End If
''
''
''Set myxl = createobject("excel.application")
''
''myxl.Workbooks.Open "C:\automation\APPLICATIONS\FFWF\DATA\FFWF_Data.xlsx" 
'''myxl.Application.Visible = true
'' 
'''this is the name of  Sheet  in Excel file "qtp.xls"   where data needs to be entered 
''set mysheet = myxl.ActiveWorkbook.Worksheets("Test_Data")
'' 
''Col=mysheet.UsedRange.columns.count
''
''msgbox Col
''
''For k = 1 To Col
''	If Trim(mysheet.cells(1,k).value) = Trim("bim:RequestId")  Then
''	mysheet.cells(i,k).value = REQUEST_ID
''	Exit For
''End If
''Next
''
'''Save the Workbook
''myxl.ActiveWorkbook.Save
'' 
'''Close the Workbook
''myxl.ActiveWorkbook.Close
'' 
'''Close Excel
'''myxl.Application.Quit
'' 
''Set mysheet =nothing
''Set myxl = nothing
''
''Browser("micclass:=Browser","name:=TMS Bus Listener.*").Close
''
''End if
'''	' Handling Error
'''	methodName = "XMLValue_Collect" : rc = ErrorHandler(methodName)
'''End Function

'
'
''Public Function Select_Process(inputData)
'inputData = "PureTvOrderFulfillmentProcess"
'Exec_Flag = "Y"
'	'On Error Resume Next
'	'Dim methodName,rc, i, process_name, row_cnt, display
'	'methodName = "Select_Process" : Select_Process = 0
'	msgbox inputData
'	If Exec_Flag = "Y" Then
'
'		JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").highlight
'		rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").Exist(10)
'			msgbox rc
'			If rc = "True" Then
'	
'				row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
'				msgbox row_cnt
'				For i = 0 To row_cnt-1 step 1
'				Print "Iteration no ="&i
'					process_name = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
'					print "process name: ="& process_name
'						If Trim(Ucase(process_name)) = Trim(ucase(inputData)) Then
'						Step_Description = "To select the required process"
'							Exp_Result = "Required Process should be displayed"
'							Actual_Res = "Required Process is displayed"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'							
'						msgbox "found" &inputData
'							JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").ClickCell i,1
'							JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=ID;Name.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell i,1
'							Flag = True
'						Exit for	
'						Else
'							'Call captureScreen
'							
'						'Exit For
'							Step_Description = "To verify update status"
'							Exp_Result = "Required Process should be displayed"
'							Actual_Res = "Required Process is not displayed"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Select_Process=-1
'							
'						End If
'				Next
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
'							'Call captureScreen
'							Step_Description = "To Verify if the selected process is loaded"
'							Exp_Result = "Selected Process should be loaded"
'							Actual_Res = "Selected Process is not loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Select_Process=-1	
'						End If
'			End If
'	End if
''End function
'
''Public Function Load_ProcessInstance()
'	'On Error Resume Next
''	'Dim methodName,rc, Flag
''	'methodName = "Load_ProcessInstance" : Load_ProcessInstance = 0
'	Flag = False
'	Exec_Flag = "Y" 
'	If Exec_Flag = "Y" Then
'	
'	rc=JavaWindow("title:=MDW Designer (ecomt199.dev.qintra.com.*").JavaButton("toolkit class:=javax.swing.JButton","label:=table24").Exist(2)
'	msgbox rc
'	
'	JavaWindow("title:=MDW Designer.*").Highlight
'		
'		rc = JavaWindow("title:=MDW Designer (ecomt199.dev.qintra.com.*").JavaButton("path:=JButton;ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","toolkit class:=javax.swing.JButton","attached text:=table24").Exist(10)
'		msgbox rc
'
'			If rc = "True" Then
'			
'				JavaWindow("title:=MDW Designer (ecomt199.dev.qintra.com.*").JavaButton("tagname:=table24","toolkit class:=javax.swing.JButton").Click
'					If JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaObject("tagname:=JPanel","toolkit class:=javax\.swing\.JPanel").Exist(10) Then 
'							Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Loaded"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							'Call captureScreen
'							Step_Description = "To Verify if Filter Process Instances Window is Loaded"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'					
'					JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaButton("tagname:=Load","toolkit class:=javax\.swing\.JButton").Click
'					wait (2)
'					If JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaTable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(1) Then
'					Flag = True
'							Step_Description = "To Verify if the Total Instances Window is Loaded"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'						Else
'							'Call captureScreen
'							Step_Description = "To Verify if the Total Instances Window is Loaded"
'							Exp_Result = "Total Instances Window should be Loaded"
'							Actual_Res = "Total Instances Window is not Loaded"
'							Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'							Load_ProcessInstance=-1
'					End If
'					
'			End If
'	End if
'''End function
''		
'''Public Function Loading_MasterRequestId(inputData)
''	inputData="100001541"
''		'On Error Resume Next
''	'Dim methodName,rc, i, master_reqid, row_cnt
''	'methodName = "Loading_MasterRequestId" : Loading_MasterRequestId = 0
''	Exec_Flag = "Y"
''	If Exec_Flag = "Y" Then
''
''		JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").highlight
''		rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(10)
''
''			If rc = "True" Then
''	
''				row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
''				For i = 0 To row_cnt-1
''					master_reqid = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
''					msgbox master_reqid
''						If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
''						msgbox "found" &inputData
''							'get the value of the status code
''							status_code = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,4)
''								msgbox status_code
''								'If Trim(Ucase(status_code)) = "In Progress" Then
''
''							rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").ClickCell (i,1)
''							msgbox rc
''							rc1=JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell (i,1)
''							msgbox rc1
''							
''							Flag = True
''								'else
''								
''								'End if
''						
''							Exit For
''						Else
'							
'						End If
'						Next
'						
'						Do
'						display = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaObject("tagname:=DesignerCanvas","toolkit class:=com\.qwest\.mdw\.designer\.pages\.DesignerCanvas").Exist
'						Loop While display = False
'			
'						If display = "True" Then
'				
'						Else
'							
'						End If
'			End If
'	End if
'End function


	
	
'	Set obj=Description.Create
'	obj("micclass").value="JavaButton"
'	obj("class description").value="push_button"
'	Set childObj=JavaWindow("title:=MDW Designer.*").ChildObjects(obj)
'	For i=0 To childObj.count-1
'	     sName=obj(i).GetROProperty("attached text")
'	    msgbox  sName
'	     
'	Next
	
	
	
	
'	JavaWindow("title:=MDW Designer.*").JavaButton("to_class:=JavaWindow","label:=table24").Highlight @@ hightlight id_;_1447709996_;_script infofile_;_ZIP::ssf10.xml_;_
'	
'	
'	
'	Set obj= Description.Create
'		obj("class name").value = "JavaButton"
'		'obj("toolkit class").value="javax.swing.JButton"
'	Set obj1=JavaWindow("title:=MDW Designer.*").ChildObjects(obj)
'	msgbox obj1.count
'	
'	
'	'Public Function Loading_MasterRequestId(inputData)
'	inputData="100001541"
'		'On Error Resume Next
'	'Dim methodName,rc, i, master_reqid, row_cnt
'	'methodName = "Loading_MasterRequestId" : Loading_MasterRequestId = 0
'	Exec_Flag = "Y"
'	If Exec_Flag = "Y" Then
'
'		JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").highlight
'		rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(10)
'
'			If rc = "True" Then
'	
'				row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
'				For i = 0 To row_cnt-1
'					master_reqid = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
'					msgbox master_reqid
'						If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
'						msgbox "found" &inputData
'							'get the value of the status code
'							status_code = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,4)
'								If Trim(Ucase(status_code)) = "In Progress" Then
'							JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").ClickCell i,1
'							JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell i,1
'							Flag = True
'								else
'							
'								End if
'				
'						Else
'						
'				Exit For
'						End If
'						Next
'						
'						Do
'						display = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaObject("tagname:=DesignerCanvas","toolkit class:=com\.qwest\.mdw\.designer\.pages\.DesignerCanvas").Exist
'						Loop While display = False
'			
'						If display = "True" Then
'							
'						Else
'							
'						End If
'			End If
'	End if
	

	
'JavaWindow("title:=MDW Designer.*").Javatoolbar("path:=ToolPane;FlowchartPage;JPanel;JLayeredPane;JRootPane;MainFrame;","tagname:=ToolPane","index:=0").Press(5)


'rx0p 123
'Public Function Loading_MasterRequestId(inputData)
'		On Error Resume Next
'	Dim methodName,rc, i, master_reqid, row_cnt
'	methodName = "Loading_MasterRequestId" : Loading_MasterRequestId = 0
'	Exec_Flag = "Y"
'	inputData = "100001541"
'	inputData1 = "In Progress"
'	
'		If Exec_Flag = "Y" Then
'	
'			JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").highlight
'			rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(10)
'	
'					If rc = "True" Then
			
'						row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
'						For i = 0 To row_cnt-1
'							master_reqid = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
'							msgbox master_reqid
'														If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
'														msgbox "found" &inputData
'															'get the value of the status code
'															status_code = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,4)
'																					If Trim(Ucase(status_code)) = Trim(ucase(inputData1)) Then
'																					msgbox "found" &status_code
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
'															Step_Description = "To select the required Master Request id"
'															Exp_Result = "Master Request Id should be displayed"
'															Actual_Res = "Master Request Id is not displayed"
'															Reporter.reportevent micFail,StepName, Actual_Res, ERROR_SCREEN_FILE
'															Loading_MasterRequestId=-1
'												
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


'JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=RunTimeDesignerCanvas","toolkit class:=com.qwest.mdw.designer.runtime.RunTimeDesignerCanvas").Highlight


'display = JavaWindow("title:=MDW Designer.*").JavaObject("tagname:=DesignerCanvas","toolkit class:=com\.qwest\.mdw\.designer\.pages\.DesignerCanvas").Exist
'msgbox display


'Public Function Loading_MasterRequestId(inputData)
'	inputData="100001541"
'	inputData1="In Progress"
		'On Error Resume Next
	'Dim methodName,rc, i, master_reqid, row_cnt
	'methodName = "Loading_MasterRequestId" : Loading_MasterRequestId = 0
'	Exec_Flag = "Y"
'	If Exec_Flag = "Y" Then
'
'		JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").highlight
'		rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").Exist(10)
'
'			If rc = "True" Then
'	
'				row_cnt = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetROProperty("rows")
'				For i = 0 To row_cnt-1
'					master_reqid = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,1)
'					msgbox master_reqid
'						If Trim(Ucase(master_reqid)) = Trim(ucase(inputData)) Then
'						msgbox "found" &inputData
'							'get the value of the status code
'							status_code = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").GetCellData(i,4)
'								msgbox status_code
'								If Trim(Ucase(status_code)) = Trim(ucase(inputData1)) Then
'
'							rc = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").ClickCell (i,1)
'							msgbox rc
'							rc1=JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").Javatable("columns_names:=Process Instance Id;Master Request Id.*","toolkit class:=javax\.swing\.JTable").DoubleClickCell (i,1)
'							msgbox rc1
'							End if
'							Flag = True
'								'else
'								
'								'End if
'						
'							Exit For
'						Else
'							
'						End If
'						Next
'						
'						Do
'						display = JavaWindow("title:=MDW Designer \(ecomt199\.dev\.qintra\.com.*").JavaObject("tagname:=DesignerCanvas","toolkit class:=com\.qwest\.mdw\.designer\.pages\.DesignerCanvas").Exist
'						Loop While display = False
'			
'						If display = "True" Then
'				
'						Else
'							
'						End If
'			End If
'	End if


'Public Function Login_SOA()

'	On Error Resume Next
'	Dim methodName, rc, Temp_Name
'	methodName = "Login_SOA" : Login_SOA = 0
'	Execute("Test_URL = " & TEST_ENV)
'
'	'To generate Test Step Description and Expected Result
'	'Loc_name = split(pageDesc,"=>",-1,1) : pageName = Page_Name(1)
'	Step_Description = "Open SOA " 
'	Exp_Result = " SOA should open Successfully " 
'	
''	If Exec_Flag = "Y" Then
''	'Call close all open browser function
''		Call closeAllBrowser
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
''
''		Browser("CreationTime:=0").Sync
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
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Select "FFWF;FFWF.tst;Test Suite: Test Suite;Test 1: invokeWebService(string, string)"
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","text:=Parasoft SOAtest - Parasoft SOAtest.*","Location:=0").WinTreeView("nativeclass:=SysTreeView32","Location:=0").Activate "FFWF;FFWF.tst;Test Suite: Test Suite;Test 1: invokeWebService(string, string)"
'				
'			Actual_Res = "Launched " & " SOA with Application Test URL -> " & Chr(13) & Test_URL & ". SOA is Launched -> " & Temp_Name
'			Reporter.ReportEvent micInfo, StepName, Actual_Res
'		Else
'			Call captureScreen
'			Actual_Res = "Launched " & TEST_BROWSER & " SOA with Application Test URL -> " & Chr(13) & Test_URL & ". SOA is not Launched -> " & Temp_Name & ". Page Expected is " & Page_Name(1)
'			Reporter.ReportEvent micFail, StepName, Actual_Res, ERROR_SCREEN_FILE
'			Login_SOA = -1
'		End If
'
'wait(2)
'
'					
''End if
'
'	'Handling Error
''	methodName = "Login_SOA" : rc = ErrorHandler(methodName)
''End Function
'
''Public Function Installorder_xml()
''	On Error Resume Next
''	Dim methodName, rc, Temp_Name
''	methodName = "Installorder_xml" : Installorder_xml = 0
'	Execute("Test_URL = " & TEST_ENV)
'
'
'	'To generate Test Step Description and Expected Result
'	Step_Description = "Load XML from the Test Data Sheet" 
'	Exp_Result = " XML is loaded successfully from the Test Data Sheet"
'	
'If Exec_Flag = "Y" Then
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
'		'Browser("CreationTime:=0").Sync
'		'rc = EXIST_(pageDesc, "")
'
'		Set fileSystemObj = createobject("Scripting.FileSystemObject")
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
'		    Exml=oSheet.Cells(i,7).value
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
'				Window("regexpwndtitle:=Parasoft SOAtest","regexpwndclass:=SWT_Window0","Location:=0").WinEditor("nativeclass:=Edit","Location:=2").Highlight
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
''oBook.Save
'oBook.Close
'oExcel.Quit
'
'End if
''End Function 
'
'
'FileName="Resp.xml"
'
'Dim sdate,FileLocation,FileName,Tagtobereplaced,strDateNode
' 
'FileLocation = "C:\automation\APPLICATIONS\FFWF\XML\" &FileName
'	
'msgbox FileLocation
'sdate= Year(date)&"-"&month(date)&"-"&day(date)
'msgbox sdate
'
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.Async = False 
'
'XMLDataFile=FileLocation
'
'xmlDoc.Load(XMLDataFile)
''get the tagname to verify
'If Trim(Ucase(FileName)) = Trim(ucase("Resp.xml")) Then
'			Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentResponse/ns2:ResponseHeader/SendTimeStamp")	
'		ElseIf Trim(Ucase(FileName)) = Trim(ucase("Ship.xml")) Then
'		 	 Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")	
'	End If
'
'strDateNode.Text = sdate
'
'
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
'Set XMLDataFile=Nothing
'Set FileLocation = Nothing
'
'
'
'Public Function Buslistnerorder(inputdata)
'	On Error Resume Next
'	Dim methodName, rc, Temp_Name,xmlDoc,filename,qfile,Responsexml,fso,ws
'	methodName = "Buslistnerorder" : Buslistnerorder = 0
'	Execute("Test_URL = " & TEST_ENV)
'
'
'	'To generate Test Step Description and Expected Result
'	Step_Description = "Load XML from the Test Data Sheet" 
'	Exp_Result = " XML is loaded successfully from the Test Data Sheet"
'	
'If Exec_Flag = "Y" Then
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
'inputdata="Resp.xml"
'	If Browser("name:=TMS Bus Listener","title:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("html id:=tbRequest","html tag:=TEXTAREA","name:=tbRequest").Exist(2) Then
'		    rc = "True"
'		    msgbox rc
'		Else
'		    rc = "False"
'		    msgbox rc
'		End If
'	
'		If rc = "True" Then			
'			
'			Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'			xmlDoc.Async = False 
'			
'			Set fso=createobject("Scripting.FileSystemObject")
'			
'			'Set qfile=fso.OpenTextFile("C:\automation\APPLICATIONS\FFWF\XML\AVS_Order_Fulfillment.xml",1)
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
'				Browser("name:=TMS Bus Listener.*").Page("title:=TMS Bus Listener.*").WebEdit("type:=textarea","html tag:=TEXTAREA","name:=tbRequest").Set Responsexml
'	     		Set ws=Nothing 
'			End if
'
'wait(2)
'
''oBook.Save
''oBook.Close
''oExcel.Quit
'
'End if
'
'' Handling Error
'	methodName = "Buslistnerorder" : rc = ErrorHandler(methodName)
'End Function
'
'
'Call Buslistnerorder(inputdata)
'

'___________________________________________________________________________________________________________________________
'# Function Name	: ToUpdateTrackingNumberinXML(inputData,inputData1)
'# Purpose			: To Load the Master request id from the total instance table
'# Parameters		:inputData- > BAN to select the row in the list
'#					:inputData1- > status value as per the input
'# Return			: 0  : Success
'#         			  -1 : Failure
'___________________________________________________________________________________________________________________________
'Public Function ToUpdateTrackingNumberinXML()
'On Error Resume Next

'Dim FileLocation,sdate,xmlDoc,XMLDataFile,methodName,rc,nodelist,strString,strNewString,strStringnumber,strNewFirstString,intAdd,strFinalID,strTrNum,strFinalID,strTrNumUrl
'methodName = "ToUpdateTrackingNumberinXML" : ToUpdateTrackingNumberinXML = 0

	Step_Description = "To update ToUpdateTrackingNumberinXML in the Ship XML response file -> "
	Exp_Result = " ToUpdateTrackingNumberinXML is updated successfully in the XML response file"

'MyFile = "C:\Automation\APPLICATIONS\FFWF\XML\Ship.xml"

		'If fileSystemObj.FileExists(MyFile) then
		    rc = "True"
		'Else
		   ' rc = "False"
		'End If

		If rc = "True" Then
		
	
FileLocation = "C:\Automation\APPLICATIONS\FFWF\XML\Ship.xml"
'print "Filelocations is:" &FileLocation
	
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False 

XMLDataFile=FileLocation

xmlDoc.Load(XMLDataFile)
'get the tagname to verify
'Set nodelist = xmlDoc.selectsinglenode(Tagtobereplaced)

 'Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/TrackingNumber")

'To update the TrackingNumber tag in the Ship.xml response file
Set nodelist = xmlDoc.getElementsByTagName("TrackingNumber")
strString = nodelist.item(0).text
		
		msgbox "the tracking number is "&strString
		
'strString="1Z75X59E0301077780"
For i=len(strString) To 1 Step-1
           If isNumeric(mid(strString,i,1)) Then
              strNewString=(mid(strString,i,1))&strNewString
           Else 
           strStringnumber=i
           Exit for  
           End If
       
           
           
Next

msgbox "The integer part of the tracking number is "&strNewString

strNewFirstString=mid(strString,1,len(strString)-len(strNewString))
intAdd=strNewString+10
msgbox "After adding 10 to the integer part, the new value is "&intAdd

strFinalID=strNewFirstString&intAdd
msgbox "The new tracking number is "&strFinalID


	
	Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingNumber")
	'msgbox strTrNum.text
	strTrNum.text = strFinalID
	
	msgbox "the tracking number is updated in the xml file"
	
	
	Actual_Res = "Updated TrackingNumber in the XML file as->"& strTextnew 
	Reporter.ReportEvent micInfo, StepName, Actual_Res
' 
 
	'To update the TrackingURL tag in the Ship.xml response file
	Set nodelist = xmlDoc.getElementsByTagName("TrackingURL")
	strText = nodelist.item(0).text
			
	msgbox "The TrackingURL is "&strText
			
	val = split(strText,"trackNums=")
	msgbox "The Tracking Number in the TrackingURL is "&val(1)
	
	strString=val(1)
	msgbox "the tracking number in the tracking url is "&strString
	
'strString="1Z75X59E0301077760"
For i=len(strString) To 1 Step-1
           If isNumeric(mid(strText,i,1)) Then
              strNewString=(mid(strText,i,1))&strNewString
           Else 
           strStringnumber=i
           Exit for  
           End If
           
           
           
Next

msgbox "The integer part of the tracking number in the tracking url is "&strNewString

strNewFirstString=mid(strString,1,len(strString)-len(strNewString))
intAdd=strNewString+10
msgbox "After adding 10 to the integer part, the new value is "&intAdd

strFinalID=strNewFirstString&intAdd
msgbox "The new tracking number is "&strFinalID
		
	strFinalUrl = Replace(strText,strString,strFinalID)
    print "The TrackingURL is updated with the new Tracking Number as "&strFinalUrl
    
    
    
	Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingURL")
	'msgbox strTrNumUrl.text
	strTrNumUrl.text = strFinalUrl
	
	




msgbox "the tracking number is updated in the tracking url of the xml file"

	Actual_Res = "Updated TrackingNumber in the XML file as->"& strTextnew 
	Reporter.ReportEvent micInfo, StepName, Actual_Res
	
xmlDoc.Save(XMLDataFile)

Set xmlDoc = nothing
End if
	 'Handling Error
	'methodName = "ToUpdateTrackingNumberinXML" : rc = ErrorHandler(methodName)
'End Function
'' --------------------------- End of Function ToUpdateTrackingNumberinXML() --------------------------------------------------------------------



	
'strText = 1Z75X59E0301077780
'
'str1 = left(strText,14)
'    print "First 14 digits of the TrackingNumber is"&str1
'    
'    remval = right(strText,4)
'    print "Last 4 digits of the TrackingNumber is"&remval
'    
'    str2 = remval+10
'    print "Updating the last 3 digits as"&str2
'    
'    strTextnew = str1&str2
'    print "The new last 4 digits of the TrackingNumber is"&strTextnew
'    
'    Set strTrNum = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingNumber")
'    msgbox strTrNum.text
'    strTrNum.text = strTextnew
'
'
'strText= http://wwwapps.ups.com/WebTracking/track?track=yes&amp;trackNums=1Z75X59E0301077780
'
'val1 = split(strText,"trackNums=")
'    print "The Tracking Number in the TrackingURL is"&val1(1)
'    
'    val2 = Replace(val1(1),remval,str2)
'    print "The Tracking Number in the TrackingURL fr the last 4 digits as "&val2
'    
'    strTextnew = Replace(strText,val1(1),val2)
'    print "The TrackingURL is updated with the new Tracking Number as "&strTextnew
'        
'
'    Set strTrNumUrl = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponse/ns2:TrackingInfo/TrackingURL")
'    msgbox strTrNumUrl.text
'    strTrNumUrl.text = strTextnew
'
