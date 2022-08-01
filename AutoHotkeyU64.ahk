#SingleInstance Force
#InstallKeybdHook
#NoEnv  					; Recommended for performance and compatibility with future AutoHotkey releases.

#Include KWSHandler.ahk
#Include SpeechDetector.ahk
#Include AHKHID.ahk

init_this_file() { 				; called automatically when the script starts
	static _ := init_this_file()
	SendMode Input  			; Recommended for new scripts due to its superior speed and reliability.
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 1 		; evt met regex (https://www.autohotkey.com/docs/commands/SetTitleMatchMode.htm)
	initKWSHandler()
}

;; Deze hotkeys zijn globale hotkeys, niet gelimiteerd tot KWS/PACS
#If
F11::test()  ; testscriptje
F12::Run, %A_AHKPath% "%A_ScriptDir%\aview.ahk" ; eigen testscriptje
:X:openlastpt::openLastPtInLog_KWS()
^Down::MoveLineDown()
^Up::MoveLineUp()


;; Hotkeys gelimiteerd tot KWS patientscherm
#If WinActive("KWS ahk_exe javaw.exe")
^b::^c				; maakt van ctrl-b gewoon ctrl-c; de originele functie van ctrl-b was uitloggen, en was te dicht tegen ctrl-c
^k::deleteLine()		; verwijderd de hele lijn (VIM style).
^d::deleteLine()
^o::Send {Down} {Home} {Enter} {Up}
F7::copyLastReport_KWS()
F9::cleanReport_KWS()		; verslag cleaner
~^c Up::clipboardcleaner() ; Haalt automatisch "uit ongevalideerd verslag" weg.
^Enter::pressOKButton()
::tiradsnodule::Run, %A_AHKPath% "%A_ScriptDir%\TIRADS-GUI.ahk" ; Work in progress, beta versie werkt wel al
:X:wervelfx::heightLossGui()
:X:hoogteverlies::heightLossGUI()
:X:pedabdomen::pedAbdomenTemplate()
^m::aanvaarderMode()

;; Hotkeys gelimiteerd tot PACS of patientscherm KWS.
#If WinActive("KWS ahk_exe javaw.exe") or (WinExist("KWS ahk_exe javaw.exe") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")))
F7::copyLastReport_KWS()
F9::cleanReport_KWS() 

; Autoscroller (in Enterprise en IMPAX)
#if WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")
$é::auto_scroll(-1, "&", "é", "Space")
$&::auto_scroll(1, "&", "é", "Space")
NumpadDiv::scrollDown() 
NumpadMult::scrollUp()

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
#If WinActive("Diagnostic Desktop") or WinActive("KWS ahk_exe javaw.exe")
XButton2::Numpad3
XButton1::Numpad6

#If WinActive("ahk_exe EXCEL.EXE") or WinActive("Notepad++")
^o::openEAD_KWS()

#If WinActive("GVIM")
Capslock::Esc

#If Winactive("ahk_exe emacs.exe")
Capslock::Ctrl
RCtrl::Capslock

test() {
	;;Controlclick, x500 y506, KWS ;; werkt
	WinGetPos, vWinX, vWinY,,, A
	SetControlDelay -1
	CoordMode, Pixel, Window
	CoordMode, Mouse, Window

	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\eadnrLabel.png
	;; Mouseclick, left, FoundX+50, FoundY+10

	;; MouseClick, left, mouseX, mouseY
	;; Controlclick, % "x" -vWinX+mouseX " y" -vWinY+mouseY+23, A ,,,1, NA Pos
	Controlclick, % "x" -vWinX+FoundX+70 " y" -vWinY+FoundY+8+23 , A ,,,, NA Pos
	;; ead := _getEAD(returnMouse := false) 
	;; MsgBox, %ead%
	;; _makeSplashText(title := "ead", ead , time := -2000)
}

scrollUp() {
	global sleepTime
	if (sleepTime = "") {
		sleepTime := 30
	}
	While GetKeyState("NumpadMult","P") {
		Send {wheelup 1}
		Sleep sleepTime
		if GetKeyState("NumpadAdd","P") {
			sleepTime := sleepTime * 0.9
		} else if GetKeyState("NumpadSub","P") {
			sleepTime := sleepTime / 0.9
		}
	}
	return
}
scrollDown() {
	global sleepTime
	if (sleepTime = "") {
		sleepTime := 30
	}
	While GetKeyState("NumpadDiv","P") {
		Send {wheeldown 1}
		Sleep sleepTime
		if GetKeyState("NumpadAdd","P") {
			sleepTime := sleepTime * 0.9
		} else if GetKeyState("NumpadSub","P") {
			sleepTime := sleepTime / 0.9
		}
	}
	return
}

PassHotkey(keypressed) { 
	; Filter last 4 numbers of Keypressed to account for variants (however picked up and release will not be able to be differentiated.
	keypressed := SubStr(keypressed, StrLen(keypressed)-3)
	Switch Keypressed {
		Case "0010": 	; back button
		MoveLineUp()
		Case "0008":	; forward button
		MoveLineDown() 
		Case "0004":		; Play/Pause
		Send {BackSpace}
		Case "0000":
		Send {Ctrl Up}
	}
	If WinActive("KWS ahk_exe javaw.exe") {
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
			Send {F3}
			Case "0400":		; F2
			closeWithoutSaving()
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
	} else if (WinExist("KWS ahk_exe javaw.exe") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe"))){
		Switch keypressed {
			Case "0020":		; EOL
			saveAndClose_KWS()
			Case "0080":		; -i-
				copyLastReport_KWS()
			Case "0040":		; INS/OVR
				validateAndClose_KWS()
			Case "0200":		; F1
				Send {F3}
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



/*
; Key codes (last 4 digits) of the Philips Speech device
0020 	; EOL
0080	; -i-
0040	; INS
0010	; Back
0001	; REC
0008	; Fwd
0004	; Play/pause
0200	; F1
0000    ; Release button
0000 	; Probably pick up signal (not sure)
0400 	; F2
0800	; F3
1000	; F4
2000	; BackButton
0001	; Probably put down signal (not sure)
*/
