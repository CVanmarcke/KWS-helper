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

;; Deze hotkeys zijn globale hotkeys, niet gelimiteerd tot KWS/PACS
#If
F11::Findbutton()  ; testscriptje
F12::Run, %A_AHKPath% "%A_ScriptDir%\aview.ahk" ; eigen testscriptje
:X:openlastpt::openLastPtInLog_KWS()
^Down::MoveLineDown()
^Up::MoveLineUp()


;; Hotkeys gelimiteerd tot KWS patientscherm
#If WinActive("Pt. ")
^b::^c				; maakt van ctrl-b gewoon ctrl-c; de originele functie van ctrl-b was uitloggen, en was te dicht tegen ctrl-c
^d::deleteLine()
F7::copyLastReport_KWS()
F9::cleanReport_KWS()		; verslag cleaner
~^c Up::clipboardcleaner() ; Haalt automatisch "uit ongevalideerd verslag" weg.
::tiradsnodule::Run, %A_AHKPath% "%A_ScriptDir%\TIRADS-GUI.ahk" ; Work in progress, beta versie werkt wel al
:X:wervelfx::heightLossGui()
:X:hoogteverlies::heightLossGUI()
:X:pedabdomen::pedAbdomenTemplate()

;; Hotkeys gelimiteerd tot PACS of patientscherm KWS.
#If WinActive("Pt. ") or (WinExist("Pt. ") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")))
F7::copyLastReport_KWS()
F9::cleanReport_KWS() 

; Autoscroller (in Enterprise en IMPAX)
#if WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")
$é::auto_scroll(-1, "&", "é", "Space")
$&::auto_scroll(1, "&", "é", "Space")

; Allows windowing in IMPAX with the numpad keys
#If WinActive("ahk_exe impax-client-main.exe")
Numpad1::F2
Numpad2::F3
Numpad3::F4
Numpad4::F5
Numpad5::F6
XButton2::F4
XButton1::F7


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
	If WinActive("Pt. ") {
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
			; KWStoExcel()
			Case "0800":		; F3
			cleanReport_KWS()
			Case "1000":		; F4
			Send ^{F8}	
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
