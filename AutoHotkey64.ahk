#SingleInstance Force
InstallKeybdHook()
SetWorkingDir(A_ScriptDir)
SetTitleMatchMode(1)
#Include "KWSHandler.ahk"
#Include "SpeechDetector.ahk"

initKWSHandler()
excelSavePath := A_ScriptDir . "\cases.xlsx"

;==================================================================
; Hier start de code voor de hotkeys. Gebruik de handleiding op https://github.com/CVanmarcke/KWS-helper en op https://autohotkey.com/docs/Hotkeys.htm om deze aan te passen.
; =================================================================


;; Deze hotkeys zijn globale hotkeys, niet gelimiteerd tot KWS/PACS
#HotIf
CapsLock::F8      ; remaps capslock naar F8 (voor de speech)
^CapsLock::CapsLock   ; om toch nog capslock te gebruiken moet je Ctrl + capslock induwen
^Down::MoveLineDown() ; Ctrl+pijltje omhoog
^Up::MoveLineUp() ; Ctrl+pijltje omlaag
:X:openlastpt::openLastPtInLog_KWS()

#HotIf (GetKeyState("LButton") and WinActive("Diagnostic Desktop - Images ("))
RButton::Ctrl ;; linker + rechtermuis samen: pan
;; linkermuis + scroll: zoom.
WheelUp::Send("{LButton Up}{Ctrl down}{WheelUp}{Ctrl Up}{LButton Down}")
WheelDown::Send("{LButton Up}{Ctrl down}{WheelDown}{Ctrl Up}{LButton Down}")

;; Hotkeys die enkel werken als het KWS patientscherm in focus is
#HotIf WinActive("KWS ahk_exe javaw.exe")
^b::^c			; maakt van ctrl-b gewoon ctrl-c; de originele functie van ctrl-b was uitloggen, en was te dicht tegen ctrl-c
^l::return    ; disables the hotkey to prevent accidental shutdown
^e::KWStoExcel(excelSavePath)
!d::deleteLine()
!x::Send("{Backspace}")       ; Alt-x
!&::MoveLineUp()
!SC003::MoveLineDown() ;; SC003 = Ã©
F7::copyLastReport_KWS()
F9::cleanReport_KWS()		; verslag cleaner
F12::Run(A_AHKPath " `"" A_ScriptDir "\aanvaardingen.ahk`"")
:X:tiradsnodule::Run(A_AHKPath " `"" A_ScriptDir "\TIRADSv2.ahk`"")
:X:tirads2::Run(A_AHKPath " `"" A_ScriptDir "\TIRADSv2.ahk`"")
:X:aanvaarder::Run(A_AHKPath " `"" A_ScriptDir "\aanvaardingen.ahk`"")
:X:startaanv::Run(A_AHKPath " `"" A_ScriptDir "\aanvaardingen.ahk`"")
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
#HotIf WinActive("KWS ahk_exe javaw.exe") or (WinExist("KWS ahk_exe javaw.exe") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe")))
F7::copyLastReport_KWS()
F9::cleanReport_KWS()
!v::validateAndClose_KWS() ; Alt-v
!s::saveAndClose_KWS()     ; Alt-s
^s::saveAndClose_KWS()     ; Ctrl-s
!c::cleanReport_KWS()      ; Alt-c
!r::copyLastReport_KWS()   ; Alt-r
!g::onveranderdMetVorigVerslag()   ; Alt-g

#HotIf (GetKeyState("LButton") and WinActive("Diagnostic Desktop - Images ("))
RButton::Ctrl ;; Linker en rechtermuisknop samen indrukken om te pannen
WheelUp::Send("{LButton Up}{Ctrl down}{WheelUp}{Ctrl Up}{LButton Down}") ;; Linkermuis + scrollen om te zoomen
WheelDown::Send("{LButton Up}{Ctrl down}{WheelDown}{Ctrl Up}{LButton Down}")

#HotIf WinActive("ahk_class SunAwtDialog ahk_exe javaw.exe") or WinActive("KWS ahk_exe javaw.exe")
;; Wanneer "vdg", "vmg" of "vwk" worden getypt zal een datumvork worden ingevoerd
;; Handig bij bijvoorbeeld het aanvaardingen scherm of querry's
^Enter::pressOKButton()
^NumpadEnter::pressOKButton()
:X*b0:vdg::insertDatePeriod(0)
:X*b0:vmg::insertDatePeriod(1)
:X*b0:vwk::insertDatePeriod(7)
;; :X*b0:gst::insertPastDatePeriod(1) ;; Gisteren: nog niet actief om verwarring te voorkomen want gst kan accidenteel wel veel geschreven worden


#HotIf WinActive("ahk_exe EXCEL.EXE") or WinActive("Google Spreadsheets - Google Chrome")
^o::openEAD_KWS()

; Allows windowing in IMPAX with the numpad keys
#HotIf WinActive("ahk_exe impax-client-main.exe")
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
		Case "5672": 	; back button
			MoveLineUp()
		Case "5664":	; forward button
			MoveLineDown() 
		Case "5660":		; Play/Pause
			Send("{BackSpace}")
		Case "5656":
			Send("{Ctrl Up}")
	}
	Switch keypressed {
			Case "8857":		; rec + back button
				Send("^{F8}")
			Case "8884":		; Ins + back button
				onveranderdMetVorigVerslag()
			Case "8872":		; backwards + back button
				Send("{F3}")
			Case "8864":		; forwards
				Send("+{F3}")
			Case "8860":		; F1
				Send("{Ctrl Up}")
				deleteLine()
				Send("{Ctrl Down}")
	}
	If WinActive("KWS ahk_exe javaw.exe") {
		Switch keypressed {
			Case "5688":		; EOL
				saveAndClose_KWS()
			Case "5684":		; -i-
				copyLastReport_KWS()
			Case "5620":		; INS/OVR
				validateAndClose_KWS()
			Case "5856":		; F1
				Send("{F3}")
			Case "6056":		; F2
				closeWithoutSaving()
			Case "6456":		; F3
				cleanReport_KWS()
			Case "7256":		; F4
				Send("^{F8}")
		}
	} else if (WinExist("KWS ahk_exe javaw.exe") and (WinActive("Diagnostic Desktop") or WinActive("ahk_exe impax-client-main.exe"))){
		Switch keypressed {
			Case "5688":		; EOL
				saveAndClose_KWS()
			Case "5684":		; -i-
				copyLastReport_KWS()
			Case "5620":		; INS/OVR
				validateAndClose_KWS()
			Case "5856":		; F1
				Send("{F3}")
			Case "6056":		; F2
				return
			Case "6456":		; F3
				cleanReport_KWS()
			Case "7256":		; F4
				return
			Case "8856":		; Back button
				Send("{Ctrl Down}")
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

/* keycodes in combination with back button (ctrl):
8888 EOL
8884 i
8820 Ins
8872 Backwards
8857 Rec
8864 Fw
8860 play
9056 F1
9256 F2
9656 F3
0456 F4
*/