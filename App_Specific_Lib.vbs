'Public Function ToUpdateDateinXML(FileName)
'
'On Error Resume Next
'Dim sdate,FileLocation,FileName,strDateNode,methodName,Flag
'
' methodName = "ToUpdateDateinXML" : ToUpdateDateinXML = 0
' Flag = False
' If Exec_Flag = "Y" Then
'FileLocation = "C:\automation\APPLICATIONS\FFWF\XML\" &FileName
'	
''msgbox FileLocation
'sdate= Year(date)&"-"&month(date)&"-"&day(date)
''msgbox sdate
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
'				Flag = True
'							Step_Description = "To Verify if Filter Process Instances Window is Opened"
'							Exp_Result = "Filter Process Instances Window should be Loaded"
'							Actual_Res = "Filter Process Instances Window is Loaded"
'							Reporter.ReportEvent micInfo, StepName, Actual_Res
'							
'		ElseIf Trim(Ucase(FileName)) = Trim(ucase("Ship.xml")) Then
'		 	 Set strDateNode = xmlDoc.SelectSingleNode("ns2:FulfillmentShipmentResponse/ns2:ShipmentResponseHeader/SendTimeStamp")	
'				End If
'
'strDateNode.Text = sdate
'
'
'xmlDoc.Save(XMLDataFile)
'
'Set xmlDoc = nothing
''Set XMLDataFile=Nothing
''Set FileLocation = Nothing
''
'	
'End If
'	 'Handling Error
'	methodName = "ToUpdateDateinXML" : rc = ErrorHandler(methodName)
'End Function
