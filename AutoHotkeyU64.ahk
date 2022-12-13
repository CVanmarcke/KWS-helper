#SingleInstance Force
#InstallKeybdHook
#NoEnv		; Recommended for performance and compatibility with future AutoHotkey releases.

#Include KWSHandler.ahk
#Include SpeechDetector.ahk
#Include AHKHID.ahk

init_this_file() {				; called automatically when the script starts
	static _ := init_this_file()
	SendMode Input			; Recommended for new scripts due to its superior speed and reliability.
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 1		; evt met regex (https://www.autohotkey.com/docs/commands/SetTitleMatchMode.htm)
	global excelSavePath
	excelSavePath := A_ScriptDir . "\cases.xlsx"

	initKWSHandler()
}

;==================================================================
; Hier start de code voor de hotkeys. Gebruik de handleiding op https://github.com/CVanmarcke/KWS-helper en op https://autohotkey.com/docs/Hotkeys.htm om deze aan te passen.
; =================================================================


;; Deze hotkeys zijn globale hotkeys, niet gelimiteerd tot KWS/PACS
#If
CapsLock::F8      ; remaps capslock naar F8 (voor de speech)
^CapsLock::CapsLock   ; om toch nog capslock te gebruiken moet je Ctrl + capslock induwen
^Down::MoveLineDown() ; Ctrl+pijltje omhoog
^Up::MoveLineUp() ; Ctrl+pijltje omlaag
:X:openlastpt::openLastPtInLog_KWS()

;; Hotkeys die enkel werken als het KWS patientscherm in focus is
#If WinActive("KWS ahk_exe javaw.exe")
^b::^c				; maakt van ctrl-b gewoon ctrl-c; de originele functie van ctrl-b was uitloggen, en was te dicht tegen ctrl-c waardoor ik er soms perongeluk op duwde. Dit zet dit uit.
^e::KWStoExcel(excelSavePath)
!d::deleteLine()
!x::Send {Backspace}       ; Alt-x
!&::MoveLineUp()
!SC003::MoveLineDown() ;; SC003 = é
F7::copyLastReport_KWS()
F9::cleanReport_KWS()		; verslag cleaner
:X:tiradsnodule::Run, %A_AHKPath% "%A_ScriptDir%\TIRADS-v2.ahk" ; WIP, beta versie werkt wel
:X:tirads2::Run, %A_AHKPath% "%A_ScriptDir%\TIRADS-v2.ahk"
:X:pedabdomen::pedAbdomenTemplate()
:X:wervelfx::heightLossGui()
:X:hoogteverlies::heightLossGUI()
:X:calcRI::RIcalculatorGUI()
:X:RIcalc::RIcalculatorGUI()
:X:volcalc::VolumeCalculator()
:X:calcvol::VolumeCalculator()
:X:vdtcalc::VDTCalculator()
:X:calcvdt::VDTCalculator()
^NumpadAdd::aanvaardOnderzoek(1, "IV veneus {+} 3 PO") ; Met contrast en die tekst
^NumpadSub::aanvaardOnderzoek(0, "") ; zonder contrast, geen tekst

;; Hotkeys die enkel werken als PACS of patientscherm KWS in focus is.
#If WinActive("KWS ahk_exe javaw.exe") or (WinExist("KWS ahk_exe javaw.exe") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")))
F7::copyLastReport_KWS()
F9::cleanReport_KWS()
!v::validateAndClose_KWS() ; Alt-v
!s::saveAndClose_KWS()     ; Alt-s
^s::saveAndClose_KWS()     ; Ctrl-s
!c::cleanReport_KWS()      ; Alt-c
!r::copyLastReport_KWS()   ; Alt-r
!g::onveranderdMetVorigVerslag()   ; Alt-g

#If WinActive("ahk_class SunAwtDialog ahk_exe javaw.exe") or WinActive("KWS ahk_exe javaw.exe")
;; Wanneer "vdg", "vmg" of "vwk" worden getypt zal een datumvork worden ingevoerd
;; Handig bij bijvoorbeeld het aanvaardingen scherm of querry's
^Enter::pressOKButton()
^NumpadEnter::pressOKButton()
:X*b0:vdg::insertDatePeriod(0)
:X*b0:vmg::insertDatePeriod(1)
:X*b0:vwk::insertDatePeriod(7)
;; :X*b0:gst::insertPastDatePeriod(1) ;; Gisteren: nog niet actief om verwarring te voorkomen want gst kan accidenteel wel veel geschreven worden

; Autoscroller (in Enterprise en IMPAX) is gemaakt door johannes. Cfr handleiding.
#if WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")
;; $SC003::auto_scroll(-1, "&", "é", "Space") ;; SC003 = é. From https://www.autohotkey.com/boards/viewtopic.php?t=17547
;; $^&::auto_scroll(1, "&", "é", "Space")

#If WinActive("ahk_exe EXCEL.EXE") or WinActive("Google Spreadsheets - Google Chrome")
^o::openEAD_KWS()

; Allows windowing in IMPAX with the numpad keys
#If WinActive("ahk_exe impax-client-main.exe")
Numpad1::F2
Numpad2::F3
Numpad3::F4
Numpad4::F5
Numpad5::F6
Numpad6::F7

;============================================
; Keybindings voor de philips speechdevice
; Kan wat complex zijn, je mag altijd zelf wat experimenteren. Bottom line is dat codes voor de knoppen vanonder staan, en je gewoon bij de switch-statement een lijn bij moet voegen of veranderderen afhankdelijk wat je wil.
; ===========================================
PassHotkey(keypressed) {
	;; only use last 4 numbers of keypressed
	;; 2 versies van philips speech devices, waarbij de laatste 4 digits van de code hetzelfde zijn.
	;; de "picked up" en "key release" kunnen zo wel niet onderscheiden worden
	keypressed := SubStr(keypressed, StrLen(keypressed)-3)
	Switch Keypressed {
		Case "0010":	; back button
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
			Case "0200":		; F1
				copyLastReport_KWS()
			Case "0400":		; F2
				closeWithoutSaving()
			Case "0800":		; F3
				cleanReport_KWS()
			Case "1000":		; F4
				Send ^{F8}
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
0020	; EOL
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