#Include %A_ScriptDir%\gdip.ahk
;AUTO-RUN FUNCTIONS MADE AND REQUIRED FOR HHH AHK SCRIPTS
func_monitor_get("mon") ;RETRIEVES ALL MONITORWORKAREA BOUNDS TO mon1, mon2 etc
func_create_exit_hotkey("Enter") ;Allows user to exit script - super important. tilde added so ESC can be used in multiple simultaneous scripts
; func_kill_onedrive() ;KILLS BACKGROUND ONE DRIVE AS IT MAY INTERFERE AS A POPUP - OLD
OnExit("func_EnableKeys") ;TURNS ALL KEYS BACK ON IF SCRIPT EXITS AT ANY POINT
; OnError("func_error_handling") Doesn't work on older versions of AHK
AQ_dir := "C:\aq-client\bin\aquarius64.exe" ; DEV "C:\devaq-client\bin\aquarius64.exe"
LocalSettings_dir := "C:\Users\Hugh.huang\AppData\Local\HHH"
settingsfile_dir := LocalSettings_dir . "\HHH_plugin_settings_local.ini"
While !FileExist(LocalSettings_dir)
	FileCreateDir, % LocalSettings_dir
While !FileExist(settingsfile_dir)
	FileAppend,, % settingsfile_dir
FileSetAttrib, +H, % settingsfile_dir, 1

;FUNCTION DEFINITIONS
func_kill_onedrive() { ; OLD
	kill_one_drive := A_WorkingDir . "\kill_one_drive.bat"
	if !FileExist(kill_one_drive)
		FileAppend, % "taskkill /IM onedrive.exe", % kill_one_drive
	Run, % kill_one_drive,,Hide
}

func_create_exit_hotkey(user_defined_hotkey) {
	global escape_disabled
	Hotkey, % "~" . user_defined_hotkey, Exit_Script, On
	Return

	#UseHook, On
	Exit_Script:
	if (escape_disabled != 1) ;allows code to set when user cannot interrupt the code with escape
		FileAppend, User-initiated exit, *
		ExitApp
	Return
	block_input:
	return
	#UseHook, Off
}

func_EnableKeys(ExitReason,ExitCode) {
	func_ToggleInput(1)
	; BlockInput, MouseMoveOff
}

; func_error_handling(exception) {
; 	func_EnableKeys(ExitReason,ExitCode)
; 	msgbox % "Error on line " exception.Line ": " exception.Message . "`nWhat: " . exception.What "`nExtra: " . exception.Extra . "`nFile: " . exception.File
; }

func_Open_Aquarius(aq_user,aq_password) {
	global
	IniRead, aq_id_check, % settingsfile_dir, IDs, aq_id
	current_matchmode := A_TitleMatchMode
	SetTitleMatchMode, 3
	imagesearch_User_Name := A_WorkingDir . "\icons\imagesearch_User_Name.png"
	imagesearch_Password := A_WorkingDir . "\icons\imagesearch_Password.png"

	If WinExist("ahk_id" aq_id_check)
	{
		WinMove, ahk_id %aq_id_check%,,mon2Left,mon2Top,mon2Right-mon2Left,mon2Bottom-mon2Top ;mon is defined in ahk_functions.ahk
		Loop
		{
			WinActivate, ahk_id %aq_id_check%
			WinWaitActive, ahk_id %aq_id_check%,,1
		}
		Until (ErrorLevel = 0)
		aq_id := aq_id_check
	}
	Else Try ;Login to new AQ session and remember the ID
	{
		Run, % AQ_dir,,,aq_pid
		WinWait, ahk_pid %aq_pid%
		While !WinExist("Sign On (SC010-01-08)")
			Sleep, 10
		WinGet, signon_id, ID, Sign On (SC010-01-08) ahk_pid %aq_pid%
		sign_on_screen :=  "ahk_id" signon_id
		if func_Image_Search_Click(sign_on_screen,imagesearch_User_Name,200,0,3,2)
			func_ControlSendInputBox(sign_on_screen,aq_user)
		if func_Image_Search_Click(sign_on_screen,imagesearch_Password,200,0,3,2)
			func_ControlSend(sign_on_screen,"string",aq_password)
		func_ControlSend(sign_on_screen,"key","{Enter}")
		func_Load_Wait(500)
		starttime := A_Now
		sign_on_attempt := ""
		While WinExist(sign_on_screen)
		{
			sleep, 10
			if (A_Now - starttime > 1)
			{
				sign_on_attempt := "failed"
				break
			}
		}
		if (sign_on_attempt = "failed")
		{
			func_ToggleInput(1) ;Enable Input
			Msgbox, AQ login failed - please check your credentials and try again.
			ExitApp
		}
		else
		{
			; WinGet, aq_id, ID, ahk_pid %aq_pid%
			login_title := ""
			While !((login_title = "Select Screen") || (login_title = "Change Password | Change (SC010-01-09)"))
				WinGetTitle, login_title, ahk_pid %aq_pid%
			if (login_title = "Select Screen") 
				WinGet, aq_id, ID, ahk_pid %aq_pid%
			else if (login_title = "Change Password | Change (SC010-01-09)")
			{
				func_ToggleInput(1) ;Enable Input
				Msgbox Please change your AQ password and then try again.
				ExitApp
			}
		}		
	}
	Catch
	{
		aq_id := "AQ_RUN_APPLICATION_ERROR"
	}
	IniWrite, %aq_id%, %settingsfile_dir%, IDs, aq_id
	SetTitleMatchMode, %current_matchmode%
	return aq_id
}

func_ToggleInput(state,key_list:="a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z|0|1|2|3|4|5|6|7|8|9|`-|`=|`[|`]|`\|`;|`'|``|`,|`.|`/|CapsLock|Space|Tab|Enter|Escape|Backspace|ScrollLock|Delete|Insert|Home|End|PgUp|PgDn|Up|Down|Left|Right|Numpad0|NumpadIns|Numpad1|NumpadEnd|Numpad2|NumpadDown|Numpad3|NumpadPgDn|Numpad4|NumpadLeft|Numpad5|NumpadClear|Numpad6|NumpadRight|Numpad7|NumpadHome|Numpad8|NumpadUp|Numpad9|NumpadPgUp|NumpadDot|NumpadDel|NumLock|NumpadDiv|NumpadMult|NumpadAdd|NumpadSub|NumpadEnter|F1|F2|F3|F4|F5|F6|F7|F8|F9|F10|F11|F12|F13|F14|F15|F16|F17|F18|F19|F20|F21|F22|F23|F24|LWin|RWin|Control|Shift|RShift|LShift|Alt|Browser_Back|Browser_Forward|Browser_Refresh|Browser_Stop|Browser_Search|Browser_Favorites|Browser_Home|Volume_Mute|Volume_Down|Volume_Up|Media_Next|Media_Prev|Media_Stop|Media_Play_Pause|Launch_Mail|Launch_Media|Launch_App1|Launch_App2|AppsKey|PrintScreen|CtrlBreak|Pause|Break|Help|Sleep|<a|<b|<c|<d|<e|<f|<g|<h|<i|<j|<k|<l|<m|<n|<o|<p|<q|<r|<s|<t|<u|<v|<w|<x|<y|<z|<0|<1|<2|<3|<4|<5|<6|<7|<8|<9|<`-|<`=|<`[|<`]|<`\|<`;|<`'|<``|<`,|<`.|<`/|LButton|RButton|MButton|XButton1|XButton2|WheelDown|WheelUp|WheelLeft|WheelRight") { ;1 is ENABLE, 0 is DISABLE
	global key_state
	if (state != 1) && (state != 0)
	{
		;Create Notice that function has not been passed appropiate values
		msgbox % "func_ToggleInput has not been passed the correct parameters. Select OK to continue without Input Disabled"
		return
	}
	if (state = 1) 
	{
		Gosub, toggle_keys_on
		BlockInput, MousemoveOff
		func_destroy_notification()
		return
	}
	else if (state = 0) ; TODO mouse input block does not work when remote desktoping in (user can still move mouse)
	{
		func_create_notification("Keys have been disabled during this script. Press ENTER to enable keys and stop the script")
		BlockInput, MouseMove		
		Gosub, toggle_keys_off
		return
	}

	toggle_keys_on:
	Loop, Parse, key_list, |
		Hotkey, % "*" . A_loopfield, do_nothing, Off	
	key_state := "ON"
	Return

	toggle_keys_off:
	Loop, Parse, key_list, |
		Hotkey, % "*" . A_loopfield, do_nothing, On	
	key_state := "OFF"
	return

	#UseHook, On
	do_nothing:
	Return
	#UseHook, Off
}

func_ControlSend(win_title,key_or_string,controls) { ;USES | AS A DELIMTER
	Process,Priority,,High
	if (key_or_string = "key")
		Loop,Parse,controls,|
		{
			WinActivate, %win_title%
			WinWaitActive, %win_title%,,1
			If (Errorlevel = 0) && WinActive(win_title)
				SendInput, %A_LoopField%
		}
	Else if (key_or_string = "string")
		{
			WinActivate, %win_title%
			WinWaitActive, %win_title%,,1
			If (Errorlevel = 0) && WinActive(win_title)
				SendInput, {Raw}%controls%
		}	
	Process,Priority,,Normal
}

func_ControlSendInputBox(win_title,string) { ;USEFUL FOR CONFIRMING THE STRING HAS BEEN ENTERED
	Clipboard := ""
	While (Clipboard != string)
	{
		func_ControlSend(win_title,"key","{Ctrl Down}|a|{Ctrl Up}")
		func_ControlSend(win_title,"string",string)
		func_ControlSend(win_title,"key","{Ctrl Down}|a|c|{Ctrl Up}")
	}
}

func_create_notification(notif_text) {
	global notification_text
	Gui notification: New, +toolwindow +hwndnotification, HHH Notification
	Gui notification: font, s14 q4 cwhite, Roboto Condensed
	Gui notification: Color, 58595b
	Gui notification: -caption
	Gui notification: Margin, 0, 0
	Gui notification: Add, Text, +BackgroundTrans +Wrap +Center x20 y20 w300 vnotification_text, % notif_text
	Gui notification: Add, Text, +BackgroundTrans +Wrap h20
	Gui notification: Show, w320 center
	WinSet, ExStyle, +0x20, HHH Notification
	Winset, Transparent, 200, HHH Notification
	WinSet, AlwaysOnTop, On,  HHH Notification
}

func_destroy_notification() {
	Gui notification: Destroy
}

func_monitor_get(name_prefix) {
	global
	Sysget, monitor_count, MonitorCount
	Loop, % monitor_count
	{
		mon_name = %name_prefix%%A_Index%
		Sysget, %mon_name%, MonitorWorkArea, A_Index
	}
}

func_Image_Search(win_title,image_path, ByRef x_img_found, ByRef y_img_found) {
	global scale_factor
	Process,Priority,,High
	WinGetPos,winpos_x,winpos_y,win_width,win_height, % win_title
	WinActivate, % win_title
	WinWaitActive, % win_title,,1
	If (Errorlevel != 0) ; If problem with image search
	{
		func_Image_Search_status := "window not found"
		func_Image_Search_success := 0
	}
	ImageSearch, x_img_found, y_img_found, 0, 0, win_width, win_height, % "*50 *TransE4E8E8 " . image_path
	If (Errorlevel != 0) ; If problem with image search
	{
		func_Image_Search_status := "image_not_found:" . image_path
		func_Image_Search_success := 0
	}
	Else
	{
		func_Image_Search_status := "image_found"
		func_Image_Search_success := 1
	}
	Process,Priority,,Normal
	return func_Image_Search_success
}

func_Image_Search_Click(win_title,image_path,x_offset:=0,y_offset:=0,click_count:=1,wait_seconds:=3) {
	if !FileExist(image_path)
	{
		func_Image_Search_Click_status := "file not found" . image_path
		return 0
	}
	starttime := A_Now
	While !func_Image_Search(win_title,image_path,x_img_found,y_img_found)
		if (A_Now - starttime > wait_seconds) 
		{
			func_Image_Search_Click_status := "image_search_timeout: " . image_path
			return 0
		}	
	click_x := x_img_found + x_offset
	click_y := y_img_found + y_offset
	WinActivate, % win_title
	WinWaitActive, % win_title,,1
	if (ErrorLevel = 0) && WinActive(win_title)
	{
		Click, %click_x%, %click_y%, 3
		func_Image_Search_Click_status := "success"
		return 1
	}
}

func_Load_Wait(timeout_ms) {
	timeout := A_TickCount + timeout_ms
	While (A_Cursor != "Wait")
		If (A_TickCount >= timeout)
			{
				break
			}	
	While (A_Cursor = "Wait")
		Sleep, 1
return
}

