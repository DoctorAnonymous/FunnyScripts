#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force


drag_enabled := 0

+^#F22::
	if(drag_enabled)
	{
		Click up
		drag_enabled := 0
	}
	else
	{
		drag_enabled := 1
		Click down
	}
	return

LButton::
	if(drag_enabled)
	{
		Click up
		drag_enabled := 0
	}
	else
		click down
	return

LButton up::
	if(drag_enabled)
	{
		click up
	}
	else
	{
		click up
	}
	return
