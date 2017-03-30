

Sub Item_PropertyChange(ByVal Name)

' 	Set objPage = Item.GetInspector.ModifiedFormPages("Automator") 
'  	Set objApDate = objPage.Controls("dcApDate")
' 	Set objStartTime = objPage.Controls("tcStartTime")
' 	Set objEndTime = objPage.Controls("tcEndTime")

    REM Select Case Name
        REM Case "Start"
			REM objStartTime.Value = objApDate.Value + #10:00:00#
			REM objEndTime.Value = objApDate.Value + #10:00:00#
    REM End Select
    
'     Set objPage = Item.GetInspector.ModifiedFormPages("Automator")
' 	Set objchkExternal = objPage.Controls("chkExternal")
' 	Set objchkTest = objPage.Controls("chkTest")
' 	Set objchkVested = objPage.Controls("chkVested")
' 	Set objReporter = objPage.Controls("txtReporter")
'     Set MyListBox = Item.GetInspector.ModifiedFormPages("Automator").Controls("OptionButton1") 
'     Select Case Name
'         Case "extChkBox"
            
'             ' objReporter.Visible = False
'     End Select
            
    
End Sub



Sub cmdCheckAvail_Click()

' 	Set myNameSpace = Application.GetNameSpace("MAPI")
' 	Set thisFolder = _
' 	myNameSpace.GetDefaultFolder(9)
	

    Set thisFolder = Application.ActiveExplorer.CurrentFolder
	
	
	'Get dates and times from User Form
		Set objPage = Item.GetInspector.ModifiedFormPages("Automator") 
		Set objApDate = objPage.Controls("dcApDate")
		Set objStartTime = objPage.Controls("tcStartTime")
		Set objEndTime = objPage.Controls("tcEndTime")
		Set objConflict = objPage.Controls("txtConflict")
		Set objConflictName = objPage.Controls("txtConflictName")
		Set objConflictStart = objPage.Controls("txtConflictStart")
		Set objConflictEnd = objPage.Controls("txtConflictEnd")
		Set objConfStartLabel = objPage.Controls("txtConfStartLabel")
		Set objConfEndLabel = objPage.Controls("txtConfEndLabel")
		Set objRoom = objPage.Controls("cmbRoom")
		conflict = False
	
	
	'Conflict Checker
		For Each olApt In thisFolder.Items
			If 	((objStartTime.Value >= olApt.Start And objStartTime.Value <= olApt.End) _
				Or (objEndTime.Value <= olApt.End And objEndTime.Value >= olApt.Start)) _
				And (objRoom.Value = olApt.Location) Then
					conflict = True
					conflictName = olApt.Subject
					conflictStart = TimeValue(olApt.Start)
					conflictEnd = TimeValue(olApt.End)
					Exit For
				Else
					conflict = False
				End If
		 Next
		 
		
		 If conflict = True Then
			objConflict.Caption = "CONFLICT!"
			objConflict.Visible = True
			
			objConflictName.Visible = True
			objConflictName.Caption = conflictName
			
			objConfStartLabel.Visible = True
			objConflictStart.Visible = True
			objConflictStart.Caption = conflictStart
			
			objConfEndLabel.Visible = True
			objConflictEnd.Visible = True
			objConflictEnd.Caption = conflictEnd
		 Else
			objConflict.Caption = "Slot Clear"
			objConflict.Visible = True
			objConflictName.Visible = False
			objConflictStart.Visible = False
			objConflictEnd.Visible = False
			objConfStartLabel.Visible = False
			objConfEndLabel.Visible = False
		 End If
	
	
End Sub


' Private Sub chkExternal_Change()
' 	'Definitions
'     ' 	Set objReporter = objPage.Controls("txtReporter")
'     ' 	objReporter.Visible = False
    
'     msgbox "Hello World"
    
' End Sub

' Private Sub chkVested_Click()
' 	'Definitions
'     ' 	Set objReporter = objPage.Controls("txtReporter")
'     ' 	objReporter.Visible = False
'     msgbox "Hello World"
' End Sub



Sub cmdSubmit_Click()
	Dim equipArray()
	'Definitions
		Set objPage = Item.GetInspector.ModifiedFormPages("Automator")
		Set objEventName = objPage.Controls("txtEventName")
		Set objFirstName = objPage.Controls("txtFirstName")
		Set objLastName = objPage.Controls("txtLastName")
		Set objPrimaryEmail = objPage.Controls("txtPrimaryEmail")
		Set objPhoneNumber = objPage.Controls("txtPhoneNumber")
		Set objchkLaptop = objPage.Controls("chkLaptop")
		Set objchkMonitor = objPage.Controls("chkMonitor")
		Set objchkSpeakers = objPage.Controls("chkSpeakers")
		Set objchkPodium = objPage.Controls("chkPodium")
		Set objchkProjector = objPage.Controls("chkProjector")
		Set objChairQuant = objPage.Controls("txtChairQuant")
		Set objSeatingStyle = objPage.Controls("cmbSeatingStyle")
		Set objRoom = objPage.Controls("cmbRoom")
		Set objApDate = objPage.Controls("dcApDate")
		Set objStartTime = objPage.Controls("tcStartTime")
		Set objEndTime = objPage.Controls("tcEndTime")
		Set objProgramDescription = objPage.Controls("txtProgDes")
		Set objcmbMicrophones = objPage.Controls("cmbMic")
		Set objTableQuant = objPage.Controls("txtTableQuant")
		Set objReserveInitials = objPage.Controls("txtReserver")
		Set objNotes = objPage.Controls("txtNotes")
		Set objPeopleQuant = objPage.Controls("txtPeopleQuant")
		Set objchkExternal = objPage.Controls("chkExternal")
		Set objchkTest = objPage.Controls("chkTest")
		Set objchkVested = objPage.Controls("chkVested")
		Set objReporter = objPage.Controls("txtReporter")
		Set objchkPopup = objPage.Controls("chkPopup")

	'Create Dropdown Lists
	'Work on this later

	'Generate Date Times
		REM startDateTime = CDbl(objApDate.Value) + CDbl(objStartTime.Value)
		REM endDateTime = CDbl(objApDate.Value) + CDbl(objEndTime.Value)
	
	'Generate Program Description Section
		strProgDesSection = "Program Description: " + VbNewLine + objProgramDescription.value
		If objProgramDescription.Value = "" Then
			strProgDesSection = ""
		End If

    'Generate Notes Section
        strNotesSection = "Notes: " + VbNewLine + objNotes.value
        If objNotes.value = "" Then
            strNotesSection = ""
        End If

	'Generate Equipment Section
		strEquipmentSection = "Equipment Needed: "
		If objchkLaptop.Value = True Then
			strEquipmentSection = strEquipmentSection + VbNewLine + "*  Laptop"
		End If
		If objcmbMicrophones.Value <> "0 Mics" Then
		    strEquipmentSection = strEquipmentSection + VbNewLine + "*  " + objcmbMicrophones.value
		End If
		If objchkSpeakers.Value = True Then
			strEquipmentSection = strEquipmentSection + VbNewLine + "*  Speakers"
		End If
		If objchkPodium.Value = True Then
			strEquipmentSection = strEquipmentSection + VbNewLine + "*  Podium"
		End If
		If objchkMonitor.Value = True Then
			strEquipmentSection = strEquipmentSection + VbNewLine + "*  Monitor"
		End If
		If objchkProjector.Value = True Then
			strEquipmentSection = strEquipmentSection + VbNewLine + "*  Projector"
		End If
		
		If strEquipmentSection = "Equipment Needed: " Then
			strEquipmentSection = strEquipmentSection + VbNewLine + "None"
		End If
	
	'Generate Chair Arrangement
		strChairArrangement = "Chairs: "		
		If objChairQuant.Value <> "" And objChairQuant.Value <> "0" Then
			strChairArrangement = strChairArrangement + objChairQuant.Value + VbNewLine + _
			"Arrangement: " + objSeatingStyle.Value
		Else
			strChairArrangement = strChairArrangement + "No chairs requested"
		End If
	
	'Generate Table Quantity
	    strTableArrangement = "Tables: "
	    If objTableQuant.value <> "0" And objTableQuant.value <> "" Then
	        strTableArrangement = strTableArrangement + objTableQuant.Value
        Else
            strTableArrangement = strTableArrangement + "No tables requested"
        End If
				
	'Appointment Form Input
	
	    If objchkExternal = True Then
	        Item.Subject = objEventName.Value + " " + "- E"
	        strReporter = ""
	        strDisclaimer = "Please donate to Richmond Public Library - https://npo.justgive.org/rpl" + VbNewLine + "DISCLAIMER: All chairs and tables will be made available to you." + VbNewLine + VbNewLine + "1. Please configure chairs and tables as you see fit." + VbNewLine + "2. Please return chairs and tables to their original configuration."
	    ElseIf objchkTest = True Then
	        Item.Subject = objEventName.Value + " " + "- T"
	        strReporter = "Rolodex Reporter: " + objReporter.Value + VbNewLine
	        strDisclaimer = ""
	    ElseIf objchkVested = True Then
	        Item.Subject = objEventName.Value + " " + "- V"
	        strReporter = "Rolodex Reporter: " + objReporter.Value + VbNewLine
	        strDisclaimer = ""
	    ElseIf objchkPopup = True Then
	        Item.Subject = objEventName.Value + " " + "- P"
	        strReporter = "Rolodex Reporter: " + objReporter.Value + VbNewLine
	        strDisclaimer = ""
	        
	    End If
	
	
		With Item
		.Location = objRoom.Value
		.Body = strReporter + "Reserved By: " + objReserveInitials.Value + VbNewLine + _
		    "Aniticipated # of people:  " + objPeopleQuant.Value + VbNewLine + VbNewLine + _
			"Name: " + objFirstName.Value + " " + objLastName.Value + VbNewLine + _
			"Email: " + objPrimaryEmail.Value + VbNewLine + _
			"Phone: " + objPhoneNumber.Value + VbNewLine + VbNewLine + _
			strEquipmentSection + VbNewLine + VbNewLine + _
			strChairArrangement + VbNewLine + VbNewLine + _
			strTableArrangement + VbNewLine + VbNewLine + _
			strProgDesSection + VbNewLine + VbNewLine + _
			strNotesSection + VbNewLine + VbNewLine + _
			strDisclaimer
		End With

        If Item.Recipients.Count = 0 Then
            With Item
        		.Recipients.Add("patricia.parks@richmondgov.com")
        		.Recipients.Add("adam.zimmerli@richmondgov.com")
        		.Recipients.Add("enrique.longton@richmondgov.com")
        	End With
        	
            If objPrimaryEmail.Value <> "" Then
    		    Item.Recipients.Add(objPrimaryEmail.Value)
            End If
        	
        End If

	Set myInspector = Item.GetInspector
 	myInspector.SetCurrentFormPage("Appointment")
 	
End Sub


Function Item_Open()
    If Item.EntryID = "" And Item.Start <> "" Then
    	Set objPage = Item.GetInspector.ModifiedFormPages("Automator") 
     	Set objApDate = objPage.Controls("dcApDate")
    	Set objStartTime = objPage.Controls("tcStartTime")
    	Set objEndTime = objPage.Controls("tcEndTime")
    
    	Set myInspector = Item.GetInspector
     	myInspector.SetCurrentFormPage("Automator")
    	cmdCheckAvail_Click()
    	
    	Item.ReminderSet = False 'removes the reminder function'
    	

'        objApDate.Value = Now()
    	objStartTime.Value = objApDate.Value + #8:00:00#
    	objEndTime.Value = objApDate.Value + #8:30:00#
    End If
	
	
End Function