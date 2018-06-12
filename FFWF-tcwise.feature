#=========================================================================
Feature: 
1.	Placing order
2.	selecting the processs and validating the db
3.	successfully setting the FFWF resp and shipment response and valiate the response
#=========================================================================    
	##########01. Place a Pure TV Order and successfully validate the KGP return and Shipment Response##########
				@Execute	
				Scenario: Place a Pure TV Order and successfully validate the KGP return and Shipment Response
				'Given		Place_Installorder
				Given		MDW_Designer
				When		Selecting_Process
				And			Loading_Request
				And			DB_Validation
				And			Updating_XML
				And			Updating_Date
				And			TrackingNumberUpdate
				And			TMS_Buslistener_FFWF_RESP
				And			Validate_XML_Response
				And			CLOSE_Browser
				And			TMS_Buslistener_FFWF_SHIPMENT_RESP
				And			Validate_XML_Response
				And			CLOSE_Browser
				Then		Validate_MDW_Complete

	##########02. Place a ENJ Order and successfully validate the KGP return and Shipment Response##########
				@ExecuteNot	
				Scenario: Place a ENJ Order and successfully validate the KGP return and Shipment Response
				Given		Place_Installorder
				Given		MDW_Designer
				When		Selecting_Process
				And			Loading_Request
				And			DB_Validation
				And			Updating_XML
				And			Updating_Date
				And			TrackingNumberUpdate
				And			TMS_Buslistener_FFWF_RESP
				And			Validate_XML_Response
				And			CLOSE_Browser
				And			TMS_Buslistener_FFWF_SHIPMENT_RESP
				And			Validate_XML_Response
				And			CLOSE_Browser
				Then		Validate_MDW_Complete