;When calling this script, aquser and aqpw must be passed as paramaters, otherwise func_Open_Aquarius(aq_user,aq_password) will fail
;receipt data must also be passed - AQNumber, Amount, Date, Import or Reverse Code

#Include %A_ScriptDir%\ahk_functions.ahk
SetBatchLines, 100
SetWinDelay, 5
SetKeyDelay, 3
SetMouseDelay, 0
SetTitleMatchMode, 3
CoordMode, Mouse, Window
SendMode, Input
OnExit("func_Exit_Import_Debtor_Receipts")
;Script Variables
	aq_user := A_Args[1] ;passed from HHH python app
	aq_password := A_Args[2] 
	debtor_receipts_client_number := A_Args[3]
	debtor_receipts_amount := A_Args[4]
	debtor_receipts_date := A_Args[5]
	debtor_receipts_note_type := A_Args[6] ;Code for Money Banked D/C Note 
	; FileAppend, % "This is a test for " . aq_user . " " . aq_password . " " . debtor_receipts_client_number . " " . debtor_receipts_amount . " " . debtor_receipts_date . " " . debtor_receipts_code, *
	debtor_receipts_entry_screen := "shlne" ; shortcut in AQ to access Debit/Credit Note Entry Screen
	debtor_receipts_entry_screen_name := "Shadow Ledger Debit/Credit Note Entry Screen | Create (SC024-13-05-02)"
	debtor_receipts_agreement_number := "001" ;May need to make dynamic in future
	debtor_receipts_BSB := "032000"
	debtor_receipts_ACC := "000533794"

	;for the images below, top-left pixel should always be a less common colour on screen
	imagesearch_Main_Screen_Filter := A_WorkingDir . "\icons\imagesearch_Main_Screen_Filter.png"
	imagesearch_Service_Provider := A_WorkingDir . "\icons\imagesearch_Service_Provider.png"
	imagesearch_Client_Number := A_WorkingDir . "\icons\imagesearch_Client_Number.png"
	imagesearch_Note_Type := A_WorkingDir . "\icons\imagesearch_Note_Type.png"
	imagesearch_BSB := A_WorkingDir . "\icons\imagesearch_BSB.png"
	imagesearch_ACC := A_WorkingDir . "\icons\imagesearch_ACC.png"
	imagesearch_Transaction_Amount := A_WorkingDir . "\icons\imagesearch_Transaction_Amount.png"
	imagesearch_Comments := A_WorkingDir . "\icons\imagesearch_Comments.png"
	imagesearch_Value_Date := A_WorkingDir . "\icons\imagesearch_Value_Date.png"
	imagesearch_error := A_WorkingDir . "\icons\imagesearch_error.gif"

func_Exit_Import_Debtor_Receipts() {
	global debtor_receipts_entry_screen_name
	global aq_id
	;closes residual debtor receipt entry screens
	SetTitleMatchMode, RegEx
	While WinExist(debtor_receipts_entry_screen_name)
		WinClose, % debtor_receipts_entry_screen_name
	SetTitleMatchMode, 3
	func_ControlSend("Select Screen" ahk_id %aq_id%,"key","{Delete}")
}

Import_Debtor_Receipts:
	Loop
	{
		func_Open_Aquarius(aq_user,aq_password) ;Returns aq_id (Current AQ process) NEED TO SCRIPT ERROR HANDLING IF AQ DOESN"T OPEN
		SetTitleMatchMode, RegEx
		While WinExist(debtor_receipts_entry_screen_name)
			WinClose, % debtor_receipts_entry_screen_name
		SetTitleMatchMode, 3
		if func_Image_Search_Click("Select Screen" ahk_id %aq_id%,imagesearch_Main_Screen_Filter,180,0,3,2) 
		{
			func_ControlSendInputBox("Select Screen" ahk_id %aq_id%,debtor_receipts_entry_screen)
			func_ControlSend("Select Screen" ahk_id %aq_id%,"key","{Enter}")
		}
		func_Load_Wait(500)
		;Shadow Ledger D/C Note Entry Screen
		WinWait, % debtor_receipts_entry_screen_name,,1
		WinGet, winid, ID, % debtor_receipts_entry_screen_name
		debtor_receipts_screen := % debtor_receipts_entry_screen_name . " ahk_id " . winid
		WinMove, % debtor_receipts_screen,,mon1Left,mon1Top,mon1Right-mon1Left,mon1Bottom-mon1Top

		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Service_Provider,180,0,3)
			func_ControlSend(debtor_receipts_screen,"key","{down}|{down}|{down}|{Tab}|{Tab}")
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Client_Number,180,0,3) 
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_client_number)
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Client_Number,180,20,3) ;note - does not use it's own image as a suitable snip does not exist for optimised imagesearch
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_agreement_number)	
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Note_Type,180,0,3)
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_note_type)
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_BSB,180,0,3)
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_BSB)
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_ACC,180,0,3)
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_ACC)
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Transaction_Amount,180,0,3)
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_amount)
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Comments,180,0,3)
			func_ControlSendInputBox(debtor_receipts_screen,"AHK script via HHH: Debtor Receipts for " . debtor_receipts_client_number . " at " . debtor_receipts_amount)
		else
			continue
		if func_Image_Search_Click(debtor_receipts_screen,imagesearch_Value_Date,180,0,3)
			func_ControlSendInputBox(debtor_receipts_screen,debtor_receipts_date)
		else
			continue
		escape_disabled := 1 ; Prevents users from escaping the script during this section (whilst the debtor receipt is being submitted to AQ)
		func_ControlSend(debtor_receipts_screen,"key","{Enter}")
		func_Load_Wait(500)
		WinWaitActive, Select Screen ahk_id %aq_id%,,1
		escape_disabled := 0
		; The code will try the receipt again if after 'Enter' is hit of the shlne screen, but after waiting for it 
	}
	Until !WinActive(debtor_receipts_screen) ;LOOP terminates if the entry was successful. If screen persists, then entry failed - try again
	ExitApp


