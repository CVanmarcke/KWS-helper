#SingleInstance Force
#InstallKeybdHook
#NoEnv  					; Recommended for performance and compatibility with future AutoHotkey releases.

#Include KWSHandler.ahk
#Include SpeechDetector.ahk
#Include toExcel.ahk
#Include AHKHID.ahk

init_this_file() { 				; called automatically when the script starts
	static _ := init_this_file()
	SendMode Input  			; Recommended for new scripts due to its superior speed and reliability.
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 1 		; evt met regex (https://www.autohotkey.com/docs/commands/SetTitleMatchMode.htm)
	initKWSHandler()
}

#If
F11::Findbutton()  ; testscriptje
F12::Run, %A_AHKPath% "%A_ScriptDir%\aview.ahk" ; eigen testscriptje
; :X:getsettingsfile::FileCopy, C:\Users\cvmarc2\AppData\Local\Philips Device Control Center\AppControlConfig.7.0.xml, \\mixer\home50\cvmarc2\uzlsystem\Bureaublad\Autohotkey\, 1
; :X:applysettingsfile::FileCopy, AppControlConfig.7.0.xml, C:\Users\cvmarc2\AppData\Local\Philips Device Control Center\, 1
:X:openlastpt::openLastPtInLog_KWS()
; :X:hidephilips::hidePhilipsDeviceControl()
:X:hoogteverlies::heightLossGUI()
^Down::MoveLineDown()
^Up::MoveLineUp()


#If WinActive("Pt. ")
^b::^c				; maakt van ctrl-b gewoon ctrl-c; de originele functie van ctrl-b was uitloggen, en was te dicht tegen ctrl-c
^k::deleteLine()		; verwijderd de hele lijn (VIM style).
^d::deleteLine()
F7::copyLastReport_KWS()
F9::cleanReport_KWS()				; verslag cleaner
~^c Up::clipboardcleaner() ; Haalt automatisch "uit ongevalideerd verslag" weg.
::tiradsnodule::Run, %A_AHKPath% "%A_ScriptDir%\TIRADS-GUI.ahk" ; Work in progress, beta versie werkt wel al
:X:hoogteverlies::heightLossGUI()
:X:pedabdomen::pedAbdomenTemplate()

#If WinActive("Pt. ") or (WinExist("Pt. ") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")))
F7::copyLastReport_KWS()
F9::cleanReport_KWS() 

; Autoscroller (in Enterprise en IMPAX)
#if WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")
$�::auto_scroll(-1, "&", "�", "Space")
$&::auto_scroll(1, "&", "�", "Space")

; Allows windowing in IMPAX with the numpad keys
#If WinActive("ahk_exe impax-client-main.exe")
Numpad1::F2
Numpad2::F3
Numpad3::F4
Numpad4::F5
Numpad5::F6
XButton2::F4
XButton1::F7

; personal settings
#If WinActive("Diagnostic Desktop") or WinActive("Pt. ")
XButton2::Numpad3
XButton1::Numpad6

#If WinActive("ahk_exe EXCEL.EXE") or WinActive("Notepad++")
^o::openEAD_KWS()

#If WinActive("GVIM")
Capslock::Esc

#If Winactive("ahk_exe emacs.exe")
Capslock::Ctrl
RCtrl::Capslock


Findbutton() {
	global
	findImage("images\aview\AVIEW_SaveResult.png", 1000)
	MouseMove, FoundX+5, FoundY+5
}
findEitherImage(image, image2, sleeptime) {
	global
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %image%
	while (ErrorLevel = 1) {
		sleep, %sleeptime%
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %image2%
		if (ErrorLevel = 0) {
			return 1
		}
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %image%
	}
	return 1
}
findImage(image, sleeptime, ErrorLevelLoop := 1) {
	global
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %image%
	while (ErrorLevel = ErrorLevelLoop) {
		sleep, %sleeptime%
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %image%
	}
	return 1
}


PassHotkey(keypressed) { 
	; Filter last 4 numbers of Keypressed to account for variants (however picked up and release will not be able to be differentiated.
	keypressed := SubStr(keypressed, StrLen(keypressed)-3)
	Switch Keypressed {
		Case "0010": 	; back
		MoveLineUp()
		Case "0008":	; forw
		MoveLineDown() 
		Case "0004":		; Play/Pause
		Send {BackSpace}
		Case "0000":
		Send {Ctrl Up}
	}
	If WinActive("Pt. ") {
		Switch keypressed {
			Case "0020":		; EOL
			saveAndClose_KWS()
			Case "0080":		; -i-
			copyLastReport_KWS()
			Case "0040":		; INS/OVR
			validateAndClose_KWS()
			;	Case "0004":		; Play/Pause
			;		Send {BackSpace}
			Case "0200":		; F1
			copyLastReport_KWS()
			Case "0400":		; F2
			; KWStoExcel()
			Case "0800":		; F3
			cleanReport_KWS()
			Case "1000":		; F4
			Send ^{F8}	
			;	Case "2000":		; Back button
			;		MsgBox, back button 
			;	Case "0000":		; picked up
			;		MsgBox, picked up
			;	Case "0001":		; Put down
			;		MsgBox, put down
			;	Case "0000":		; Released
			;		;do nothing, release
			;	Case keypressed:
			;		MsgBox, unknown key: %keypressed%
		}
	} else if (WinExist("Pt. ") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe"))){
		Switch keypressed {
			Case "0020":		; EOL
			saveAndClose_KWS()
			Case "0080":		; -i-
				copyLastReport_KWS()
			Case "0040":		; INS/OVR
				validateAndClose_KWS()
			Case "0200":		; F1
				copyLastReport_KWS()
			Case "0400":		; F2
				return
			Case "0800":		; F3
				cleanReport_KWS()
			Case "1000":		; F4
				return
			Case "2000":		; Back button
				Send {Ctrl Down}
		}
	}
}

