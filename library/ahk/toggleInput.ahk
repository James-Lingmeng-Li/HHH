;When calling this script from python, input ON/OFF state must be passed as a variable.
#Persistent
#SingleInstance, force
#Include %A_ScriptDir%\ahk_functions.ahk

SetBatchLines, 100
SetWinDelay, 5
SetKeyDelay, 3
SetMouseDelay, 0
SetTitleMatchMode, 3
CoordMode, Mouse, Window
SendMode, Input

;Script Variables
	toggleInputState := A_Args[1] ;passed from HHH python app, should be 1 for 'on' or 0 for 'off'

if (toggleInputState = "OFF") ; turn all key input off
{
    func_ToggleInput(0)
}
else ; destroys script which runs the included onExit function that toggles input back on, and enables all keys 
    ExitApp
