'###########################################################################################################################
'#
'#   <APP_Name>_Constants_File:	Contains Constants Descriptions for <APP_Name>
'#
'#############################################################TestCase Status##############################################################
' --------------------------------------------------------------------------------------------------------------------------
' 	Test Environment Variables:-
' --------------------------------------------------------------------------------------------------------------------------
'	Test URLs:

	Dim SOATest1 : SOATest1 = "http://ecomt200.dev.qintra.com:7622/FulfillmentWFMDWWeb/MDWWebService?WSDL"
	Dim SOATest2 : SOATest2 = "http://ffwf-itv2.dev.qintra.com/FulfillmentWFMDWWeb/MDWWebService?WSDL"
	Dim SOAE2E : SOAE2E = "http://ffwf-e2e.dev.qintra.com/FulfillmentWFMDWWeb/MDWWebService?WSDL"
	
	Dim IOMTest1 : IOMTest1 = "https://iom-itv1.dev.qintra.com/IOMWeb///system/systemInformation.jsf"
	Dim IOMTest2 : IOMTest2 = "https://iom-itv2.dev.qintra.com/IOMWeb/authentication/login.jsf"
	Dim IOME2E : IOME2E = "https://iom-e2e.dev.qintra.com/IOMWeb/authentication/login.jsf"
	
	Dim MDWTest1 : MDWTest1 = "http://ecomt199.dev.qintra.com:7622/FulfillmentWFMDWDesignerWeb"
	Dim MDWTest2 : MDWTest2 = "http://ecomt200.dev.qintra.com:7623/FulfillmentWFMDWDesignerWeb"
	Dim MDWE2E : MDWE2E = "http://ecomt199.dev.qintra.com:7621/FulfillmentWFMDWDesignerWeb"
	
	Dim TMSBusTester : TMSBusTester = "http://x7009075/TMSWebUtilities/TMSBusTester.aspx"
	Dim E2E  : E2E  = "http://ntm-e2ew.dev.intranet/arsys/shared/login.jsp"
' --------------------------------------------------------------------------------------------------------------------------
'	Object Loading Time Constants
	
	Dim TIMEOUT_1 : TIMEOUT_1 = 10		'	Application Page Loading TimeOut Value	-	Change only value
	Dim TIMEOUT_2 : TIMEOUT_2 = 20		'	Object Loading TimeOut Value			-	Change only value
	Dim TIMEOUT_3 : TIMEOUT_3 = 120		'	Object Loading Maximum TimeOut Value	-	Change only value
	
	
	'FFWF-ST1 wsdl-  http://ecomt200.dev.qintra.com:7622/FulfillmentWFMDWWeb/MDWWebService?WSDL 
'	DB Connection
	'DBConnection_String = "User ID = test;Password = test;Data Source = test_e2e"
' --------------------------------------------------------------------------------------------------------------------------
' 	Handle error:-
	'Dim method_Name : method_Name = "FFWF_Constants_File" : Call ErrorHandler(method_Name)
' --------------------------------------------------------------------------------------------------------------------------

'*******************   End of <APP_Name>_Constants_File   ******************************************************************
