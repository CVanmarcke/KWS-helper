;****************************************************************
;																*
;     		Kuist het verslag op, zet de puntjes op de i 		*
;																*
;**************************************************************** 
;																*
; 	Auteur: Cedric Vanmarcke									*
; 																*
; 	Handleiding: zie readme.txt bestand							*
;																*
; 	Vrij te gebruiken door iedereen								*
;																*
;	Voor vragen: cedric.vanmarcke@uzleuven.be					*
; 	Bij fouten, graag het vorige verslag, huidige verslag 		*
;		en resultaat doorsturen naar mijn email.				*
;																*
;****************************************************************

initKWSHandler() {
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 1 		;evt met regex (https://www.autohotkey.com/docs/commands/SetTitleMatchMode.htm)
	global logfile
	logfile := "logfile.csv"
}

kwsReportCleaner() {
	; Actieve querry zou aberrante medicatieschemas moeten vinden en fixen"
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg)[\s\S]+(?:ONDERZOEKE?N?:\R{1,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<content>[\s\S]+?)(?:\R*$|[\n\r]{2,}\*\* Eind)"
	; Werkende querry--> RegexQuerry := "(?<header>(?:Leuven|Pellenberg)[\s\S]+(?:ONDERZOEKE?N?:\R{1,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE[\s\S]+toegediend:\s[\d.,]{0,4}\s(?:ml|mg|zakje|spuit)\R)?)\R*(?<content>[\s\S]+?)(?:$|[\n\r]{2,}\*\* Eind)"
	; (?<header>(?:Leuven|Pellenberg)[\s\S]+(?:ONDERZOEKE?N?:\R)(?<type>(?:.+\R?)+)(?:\R{2,}TOEGEDIENDE MEDICATIE[\s\S]+toegediend:\s[\d.,]{0,4}\s(?:ml|mg|zakje|spuit))*)\R{2,}(?<content>[\s\S]+?)(?:$|[\n\r]{2,}\*\* Eind)"
	; bigquerry "(?:Leuven|Pellenberg)[\s\S]+(\d+-\d+-\d+)[\s\S]+(?:KLINISCHE INLICHTINGEN:[\n\r])([\s\S]+)[\n\r]{2}(?:DIAGNOSTISCHE VRAAGSTELLING:[\n\r])([\s\S]+)[\n\r]{2}(?:ONDERZOEKE?N?:[\n\r])((?:.+[\n\r]?)+)[\n\r]{2,}(?:TOEGEDIENDE MEDICATIE[\s\S]+toegediend:\s[\d.,]{0,4}\s(?:ml|mg|zakje|spuit)[\n\r]+)*([\s\S]+?)(?=[\r\n]*\*\* einde|$)"
	copyKWSreporttoclip()
	RegExMatch(clipboard, RegexQuerry, report)	
	_copytexttoKWS(reportheader . "`n" . cleanreport(reportcontent))
}

copyKWSreporttoclip(selectReportBox := True) {
	if WinExist("Pt. ")
		WinActivate 
	else
		throw Exception("KWS has no patient open!", -1)
	if (selectReportBox) {
		CoordMode "Pixel"
		MouseGetPos, mouseX, mouseY
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bevindingenLabel.png
		if (ErrorLevel = 2)
			_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000, updateSplashExists := False)
		else if (ErrorLevel = 1)
			_makeSplashText(title := "Error", text := "Error finding the report field: did the action succeed nonetheless?", time := -2000, updateSplashExists := False)
		else {
			MouseClick, left, FoundX+100, FoundY+200
			MouseMove, mouseX, mouseY
		}
	}
	clipboard := ""             			; maakt het clipboard leeg 
	Send, ^a                    			; select all
	Send, ^c                    			; copy
	ClipWait, 1                 			; wacht tot er data in het clipboard is
	if (ErrorLevel)             			; als NOT, is er data in clipboard
		throw Exception("Could not copy data to clipboard!", -1)                 				; STOPT als geen data in clipboard
}

_copytexttoKWS(text, overwrite := true) {
	If WinActive("Pt. ") {
		tempclip := clipboard
		clipboard := ""  
		clipboard := text           	; maakt het clipboard leeg 
		ClipWait, 1			; wacht tot er data in het clipboard is
		if (overwrite)
			Send, ^a
		Send, ^v 
		clipboard := tempclip
		return
	}
	if WinExist("Pt. ")  {
		WinActivate
	} else {
		_makeSplashText("No patient window exists", 2000)
		return
	}
	CoordMode "Pixel"
	MouseGetPos, mouseX, mouseY
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bevindingenLabel.png
	if (ErrorLevel = 2) {
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000, updateSplashExists := False)
		return
	}
	else if (ErrorLevel = 1) {
		_makeSplashText(title := "Error", text := "Could not find the report field!", time := -2000, updateSplashExists := False)
		return
	}
	MouseClick, left, FoundX+100, FoundY+200
	_copytexttoKWS(text, overwrite)
	return
}

cleanreport(inputtext) {
	; TODO
	; dd 01-01-2020 vervangen door "van datum"
	; \R-woord --> - woord
	; 1) Hoofdletter
	; T1/Th2

	inputtext := RegExReplace(inputtext, "im)^(besluit|conclusie)", "CONCLUSIE")						; replaces case insensitive besluit/conclusie door upper
	if (RegExMatch(inputtext, "m)^\. .+\R") OR RegExMatch(inputtext, "m)^.+# ?.?\R")) { 			; only executes if there is ". " or "#" in the script
		inputtext := _sorttext(inputtext) 
	}
	inputtext := StrReplace(inputtext, "bekend", "gekend", CaseSensitive := false)
	inputtext := StrReplace(inputtext, "foraminaal spinaal stenose", "foraminaal- of spinaalstenose") 		
	inputtext := StrReplace(inputtext, "iffuse restrictie", "iffusie restrictie") 		
	inputtext := StrReplace(inputtext, "ormale doorgankelijkheid van de", "ormaal doorgankelijke") 		
	inputtext := StrReplace(inputtext, "pig katheter", "PIC katheter", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "flair", "FLAIR", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "fascikels graad", "Fazekas graad", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "segment VIII", "segment 8") 
	inputtext := StrReplace(inputtext, "segment VII", "segment 7") 		
	inputtext := StrReplace(inputtext, "segment VI", "segment 6") 		
	inputtext := StrReplace(inputtext, "segment IV", "segment 4") 
	inputtext := StrReplace(inputtext, "segment V", "segment 5") 	
	inputtext := StrReplace(inputtext, "segment III", "segment 3") 		
	inputtext := StrReplace(inputtext, "segment II", "segment 2") 	
	; inputtext := RegExReplace(inputtext, "[\n\r]GECOMMUNICEERDE DRINGENDE BEVINDINGEN:[\n\r]$", "") 	; removes if no CONCLUSIE
		
	; inputtext := RegExReplace(inputtext, "i)gekende?", "\#\#\#")				
	; inputtext := RegExReplace(inputtext, "i)op niveau van", "in")

	inputtext := RegExReplace(inputtext, "(?<=^|[\n\r])\*\s?(.+?):? ?(?=\R)", "* $U1:")				; adds : at end of string with * and makes uppercase. Not done with m) because of strange bug where it woudl only capture the first
	inputtext := RegExReplace(inputtext, "m)(?<=[\w\d\)])\s?$", ".")								; adds . to end of string, word, digit or )
	inputtext := RegExReplace(inputtext, "m)(?<=\. |^- |^)(\w)", "$U1") 							; converts to uppercase after ., newline or newline -
	inputtext := RegExReplace(inputtext, "(?<=:)\ ?(\w)", " $L1")									; converts uppercase after : to lowercase
	inputtext := RegExReplace(inputtext, "([CThD]\d{1,2}[\/-])[TD](?=\d{1,2})", "$1Th") 			; corrects T1/X to Th1 TODO: werkt neit T11-L3
	inputtext := RegExReplace(inputtext, "[TD](?=\d{1,2}[\/-][ThDL]{1,2}\d{1,2})", "Th") 			; corrects X/T1 to Th1
	inputtext := RegExReplace(inputtext, "((?:C|Th|L|S)\d{1,2})\/((?:C|Th|L|S)\d{1,2})", "$1-$2") 	; corrects L1/L2 to L1-L2
	inputtext := RegExReplace(inputtext, "(\d{1,2})\/(\d{1,2})\/(\d{2,4})", "$1-$2-$3") 			; corrects d/m/y tot d-m-y
	inputtext := RegExReplace(inputtext, "\R{3,}", "`n`n") 											; replaces triple+ newline with double
	inputtext := RegExReplace(inputtext, "im)^(?=\w|\()(?!CONCLUSIE|Vergel[ei]j?k[ie]n|Mede in|In vergel|NB|Nota|Storende|Suboptim|Reserv|Naar [lr])", "- ")	; adds - to all words and (, excluding BESLUIT, vergeleken...
	inputtext := RegExReplace(inputtext, "m)^- (.+:[\n\r])(?![ .\n\r\t])", "$1")							; removes - if :, except if after the newline followed by whitespace or . or... 
	inputtext := RegExReplace(inputtext, "(?<=CONCLUSIE:[\n\r])- (.+\s?)$", "$1") 					; removes - if als maar 1 lijn conclusie, zal het het streepje weg doen.
	inputtext := RegExReplace(inputtext, "(\d )a( \d)", "$1à$2") 									; maakt à als a tussen 2 getallen.
	inputtext := RegExReplace(inputtext, "- {2,}", "- ")
	inputtext := RegExReplace(inputtext, "i)supervis.*", "") ; verwijder supervisie.

	return inputtext
}

_sorttext(inputtext) {
	;TODO: zet geen enter tussen inputtext en de dotlines
	if (RegExMatch(inputtext, "([\s\S]+)(\RCONCLUSIE[\s\S]+)", split)) {
		outputtext := _sorttext(split1) . split2
		return outputtext
	} else if (RegExMatch(inputtext, "([\s\S]+)(\*.+)([\s\S]+)", split)) {
		outputtext := _sorttext(split1) . split2 . _sorttext(split3)
		return outputtext
	} else {
		dotlines := ""
		while (RegExMatch(inputtext, "m)(?:\R|^)\.?\ ?(.+#.?)$" , line)) { 	;gets all lines with XXXXX #
			dotlines := dotlines . "`n. " . line1
			inputtext := StrReplace(inputtext, line , "")
		}
		while (RegExMatch(inputtext, "m)(?:\R|^)\.\ ?(.+)$" , line)) { 		;gets all lines with . XXXXX
			dotlines := dotlines . "`n. " . line1
			inputtext := StrReplace(inputtext, line , "")
		}
		outputtext := inputtext . "`n" . StrReplace(StrReplace(dotlines, "#" , ""), ". - ", ". ") . "`n`n"
		return outputtext
	}
}

vorigVerslagCopy() {
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg)[\s\S]+(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{1,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl |tov ).{5,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren).+)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"
	; (?<header>(?:Leuven|Pellenberg)[\s\S]+(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R)(?<type>(?:.+\R?)+)\R{2,}(?:TOEGEDIENDE MEDICATIE[\s\S]+toegediend:\s[\d.,]{0,4}\s(?:ml|mg|zakje|spuit)\R+)?)(.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl |tov ).{5,45}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren).+)?\R+(?<content>[\s\S]+?)(?:$|\*\* Eind)"
	; bigquerry "(?:Leuven|Pellenberg)[\s\S]+(\d+-\d+-\d+)[\s\S]+(?:KLINISCHE INLICHTINGEN:[\n\r])([\s\S]+)[\n\r]{2}(?:DIAGNOSTISCHE VRAAGSTELLING:[\n\r])([\s\S]+)[\n\r]{2}(?:ONDERZOEKE?N?:[\n\r])((?:.+[\n\r]?)+)[\n\r]{2,}(?:TOEGEDIENDE MEDICATIE[\s\S]+toegediend:\s[\d.,]{0,4}\s(?:ml|mg|zakje|spuit)[\n\r]+)*([\s\S]+?)(?=[\r\n]*\*\* einde|$)"

	if WinExist("Pt. ") 
		WinActivate 
	else
		throw Exception("KWS has no patient open!", -1)
	
	CoordMode "Pixel"
	MouseGetPos, mouseX, mouseY
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bevindingenLabel.png
	if (ErrorLevel = 2) {
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000, updateSplashExists := False)
		return
	}
	else if (ErrorLevel = 1) {
		_makeSplashText(title := "Error", text := "Could not find the report field!", time := -2000, updateSplashExists := False)
		return
	}
	MouseClick, left, FoundX+100, FoundY+200
	copyKWSreporttoclip(selectReportBox := False)
	
	currentreportunclean := clipboard
	clipboard := ""
	
	MouseClick, right, FoundX+100, FoundY+200 
	Send {Down}		    					; klik pijltje naar beneden
	Send {Down}
	Send {Enter}		    				; selecteerd "neem laatste verslag over"

	Sleep, 100		    					; geeft tijd om vorig verslag te laden, kan evt verhoogd of verlaagd worden (200 werkt sowieso)
 
 	try	copyKWSreporttoclip(selectReportBox := False)
	catch e {
		Send {Enter}
		_makeSplashText(title := "ERROR", text := "Geen laatste gelijkaardig verslag gevonden!", time := -3000, updateSplashExists := False)
		return
	}

	if (RegExMatch(clipboard, "m)^ingevoerde beelden$" )) {
		Send, ^z
		Send, ^z
		Send ^{F8}										; Initieer dictee (ctrl F8)
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000, updateSplashExists := False)
		return											; Klaar
	} 

	oldreportunclean := clipboard				; zet variabele gelijk aan clipboard
	clipboard := ""             			; maakt het clipboard leeg

	RegExMatch(oldreportunclean, RegexQuerry, oldreport)
	RegExMatch(currentreportunclean, RegexQuerry, currentreport)
	compared := ""

	if (oldreportcomparedwith && oldreportcompdate) {
		compared := StrReplace(oldreportcomparedwith, oldreportcompdate, oldreportdate)
		oldreportcontent :=  StrReplace(oldreportcontent, oldreportcompdate, oldreportdate)
	} else {
		compared := "In vergelijking met het voorgaande onderzoek van " . oldreportdate . ":" 
	}

	; zoekt naar het besluit, en als het het vind voegt het "vergelijking met" toe of veranderd eht de datum van "in vergelijking met"
	conclusielocatie := RegExMatch(oldreportcontent, "(BESLUIT|CONCLUSIE).*[\n\r]", conclusieText)
	if (conclusielocatie) {
		RegExMatch(oldreportcontent, "(?:ergel(?:ij|e)k|opzichte|vgl |tov ).{5,55}?(?<date>\d+[-\/.]\d+[-\/.]\d+)",conclcompare, conclusielocatie - 5)
		if (conclcomparedate) {
			oldreportcontent := StrReplace(oldreportcontent, conclcomparedate, oldreportdate)
		} else {
			oldreportcontent := RegExReplace(oldreportcontent, "((?:BESLUIT|CONCLUSIE).*)\R", "$1`nIn vergelijking met het voorgaande onderzoek van " . oldreportdate . ":`n", ,1, conclusielocatie-5)
		}
	}

	if InStr(currentreporttype, "rx") {
		_copytexttoKWS(currentreportheader . "`n" . compared . "`n`n" . cleanreport(oldreportcontent))
	} else {
		_copytexttoKWS(currentreportheader . "`n" . compared . "`n`n" . oldreportcontent)
	}
	Send ^{F8}										; Initieer dictee (ctrl F8)
	MouseMove, mouseX, mouseY
	return											; Klaar
}

validateAndClose(){
	if WinExist("Pt. ") 
		WinActivate
	else
		return
	global	splashExists
	if (splashExists == "" or splashExists == False) {
		_makeSplashText(title := "Validate function", text := "Press the button again to validate and close.", time := -3000)
		Send, ^s
	} else {
		_destroySplash()
		CoordMode "Pixel"
		MouseGetPos, mouseX, mouseY
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\valideersluit.png
		if (ErrorLevel = 2)
			_makeSplashText(title := "Validate report", text := "Something went wrong when looking for the button", time := -2000, updateSplashExists := False)
		else if (ErrorLevel = 1)
			_makeSplashText(title := "Validate report", text := "ERROR: validate button not found!", time := -2000)
		else {
			WinGetTitle, title, A
			ead := SubStr(title, InStr(title, "(")+1, 8)
			MouseClick, left, FoundX+5, FoundY+5
			MouseMove, mouseX, mouseY
			_log(ead, "Gevalideerd en gesloten")
			_makeSplashText(title := "Gevalideerd", text := "Gevalideerd.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000, updateSplashExists := False)
		}
	}
	return
}

SaveAndCloseReportKWS() {
	if WinExist("Pt. ") 
		WinActivate
	else
		return
	global	splashExists
	if (splashExists == "" or splashExists == False) {
		_makeSplashText(title := "Save function", text := "Press save button again to close.", time := -3000)
		Send, ^s
	} else {
		_destroySplash()
		CoordMode "Pixel"
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bewarenGreyedButton.png
		if (ErrorLevel = 2) {
			_makeSplashText(title := "Save function", text := "Something went wrong when checking if it was alreade saved", time := -2000, updateSplashExists := False)
		} else if (ErrorLevel = 1) {
			_makeSplashText(title := "Save function", text := "Try again, report was not saved", time := -2000)
			Send, ^s
		} else {
			WinGetTitle, title, A
			ead := SubStr(title, InStr(title, "(")+1, 8)
			_log(ead, "Opgeslagen en gesloten")
			WinClose, Pt. 
			_makeSplashText(title := "Opgeslagen", text := "Opgeslagen.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000, updateSplashExists := False)
		}
	}
	return
}

_makeSplashText(title := "Splash title", text := "Splash text", time := -3000, updateSplashExists := True) {
	global	splashExists
	_destroySplash()
	splashExists := updateSplashExists
	Gui, splashGui:+AlwaysOnTop +Disabled -SysMenu +Owner  	; +Owner avoids a taskbar button.
	Gui, splashGui:Add, Text,, %text%
	Gui, splashGui:Show, NoActivate, %title%		; NoActivate prevents taking the focus
	SetTimer, RemoveSplash, %time%
	return
	RemoveSplash:
	_destroySplash()   		; label is used because only a label can be called by settimer
	return					; dit zorgt dat er een hoop fout loopt, opletten en zien of dit niet verwijderd kan worden. 
							; Ik heb het nu in de functie geplaatst: eens zien of het nog allemaal fout loopt...
}

_destroySplash() {
	global	splashExists
	splashExists := False
	Gui, splashGui:Destroy
}

heightLossGUI() {
	Global Height1
	Global Height2
	Gui, heightLoss:+LastFound
	GuiHWND := WinExist()
	
	Gui, heightLoss:Add , Text  ,        , Height collapsed and normal vertebra
	Gui, heightLoss:Add , Edit  , vHeight1,
	Gui, heightLoss:Add , Edit  , vHeight2,
	Gui, heightLoss:Add , Button, Default, OK
	Gui, heightLoss:Show, , Vertebral height loss calculator
	WinWaitClose, ahk_id %GuiHWND%  		;--waiting for gui to close
	WinActivate, Pt. 
	return _copytexttoKWS(result, false)               	;--returning value
	;-------
	heightLossButtonOK:
	  GuiControlGet, h1, , Height1
	  GuiControlGet, h2, , Height2
	  hl := _calcHeightLoss(h1, h2)
	  result := "hoogteverlies van " . hl[1] . " mm of " . hl[2] . "%"
	  Gui, heightLoss:Destroy
	return
	;-------
	heightLossGuiEscape:
	heightLossGuiClose:
	  result := ""
	  Gui, heightLoss:Destroy
	return
}

_calcHeightLoss(h1, h2) {
	absolute := Max(h1,h2) - Min(h1,h2)
	percentage := Round((1 - (Min(h1,h2) / Max(h1,h2))) * 100)
	return [absolute, percentage]
}

_pasteClip(str) {
	clipboard := ""
	Clipboard := str
	ClipWait,1
	Send, ^v
	return str
}

openEAD(input := ""){
	; if no EAD number provided, tries to get it from clipboard or get it from selected text
	if (RegExMatch(input, "\D(\d{8})\D", ead) == 0 and RegExMatch(clipboard, "\D(\d{8})\D", ead) == 0) { 
		Clipboard := ""
		Send, ^c
		ClipWait, 1
		if (RegExMatch(clipboard, "\D(\d{8})\D", ead))
			openEAD(ead1)
		return
	} else {
		clipboard := ead1
		if WinExist("Startscherm (") {
			WinActivate
			Send, ^+z
			Sleep, 300
			if WinExist("Zoek pat") {
				WinActivate
				Send, ^v
				Send, {Enter}
			}
		return
		} else {
		return
		}
	}
}

_log(str, extrastr*) {
	global logfile
	FormatTime, timestring, A_Now, yyyy-MM-dd HH:mm:ss
	output := timestring  . "|" . str
	for index, s in extrastr
		output .= "|" . s
	if FileExist(logfile)
		output := "`n" . output
	FileAppend, %output%, %logfile%
}

openLastPtInLog() {
	global logfile
	Loop, read, %logfile%
		lastline := A_loopreadline
	openEAD(lastline)
	return
}

auto_scroll(richting := 1, decreaseKey := "&", increaseKey := "é", directionKey := "Space", pauseKey := """"){ ; Automatisch scrollen. Versnel & vertraag. Gemaakt door Johannes Devos
	Suspend, On
	; Hotkey, If
	; try Hotkey, %decreaseKey%, Off
	; try Hotkey, %increaseKey%, Off
	; try Hotkey, %directionKey%, Off
  MouseGetPos,,,windowUnderCursor
  WinGet, temp, ProcessName, ahk_id %windowUnderCursor%
  If (temp = "impax-client-main.exe" OR temp = "syngo.Common.Container.exe" OR temp = "javaw.exe" OR temp = "javawClinapps.exe"){
	keys := "{" . decreaseKey  . "}{" . increaseKey . "}{" . directionKey . "}{" . pauseKey . "}" 
    hook := InputHook("L0", keys)
    hook.VisibleNonText := false
    hook.Start()
    static sleep_delay := 300
    endloop := 0
    hook.OnChar := Func("_auto_scroll_down_helper")
    Loop{
		if not GetKeyState(pauseKey) {
			if (richting = 1)
				send {wheeldown 1}
			if (richting = -1)
				send {wheelup 1}
		}
      
		Sleep sleep_delay
	  
		if (!hook.InProgress){
			if (hook.EndReason = "Stopped"){
				break
			} else {
				if (hook.EndKey = decreaseKey){
					sleep_delay := sleep_delay / (0.85)
				}
				if (hook.EndKey = increaseKey){
					sleep_delay := sleep_delay * (0.85)
				}
				if (hook.Endkey = directionKey){
					richting := richting * (-1)
				}
				hook.Start()
			}
		}
    }
    hook.Stop()
    Suspend, Off
    return
  }
  ; Hotkey, If
 ;  Hotkey, %increaseKey%, On
  ; Hotkey, %decreaseKey%, On
 ;  Hotkey, %directionKey%, On
  ;Send ^&
}

_auto_scroll_down_helper(ih, char){
  ih.Stop()
}

