; ****************************************************** 
;	Dit bestand bevat de eigenlijke functies, die worden opgeroepen door autohotkey.
;
; 	Handleiding: zie readme bestand	of https://github.com/CVanmarcke/KWS-helper
;								
; 	Auteur: Cedric Vanmarcke
;
;	Voor vragen: cedric.vanmarcke@uzleuven.be	
; 	Bij fouten, graag het vorige verslag, huidige verslag 		
;		en resultaat doorsturen naar mijn email.	
;		
;****************************************************************

; CALLABLE FUNCTIONS
; --------------------------------------

initKWSHandler() {
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 1
	Coordmode, Pixel
	Coordmode, Mouse

	global logfile
	logfile := "logfile.csv"

	;;clean teams: tmp folder, cache folder, old logs file
	;; heel veel: Service Worker/CacheStorage/.... (momenteel 66 mb)
	;; Cache: 19.6 mb
	;; staan allemaal in P:\uzlsystem\AppData\Microsoft\Teams

	;; indien blockinput true is, zullen de volgende knoppen geblokkeerd worden:
	global blockinput
	blockinput := false
	#If blockinput = true
	Hotkey, If, blockinput = true
	Hotkey, Enter, _blockInputHelper
	Hotkey, RButton, _blockInputHelper
	Hotkey, LButton, _blockInputHelper
}

;=================================================
; Verslag opkuiser:
; Als er een bepaalde aanpassing je niet aanstaat, kan je die lijn verwijderen.
;=================================================
cleanreport(inputtext) {
	inputtext := RegExReplace(inputtext, "im)^(besluit|conclusie)", "CONCLUSIE")				; replaces case insensitive besluit/conclusie door upper
	if (RegExMatch(inputtext, "m)^\. ?.+(?:\R|$)") OR RegExMatch(inputtext, "m)^.+# ?.?(?:\R|$)")) { 			; only executes if there is ". " or "#" in the script
		inputtext := _sorttext(inputtext) ; zet alle zinnen met een punt vooraan, onder het verslag.
	}
	sleep, 50
	inputtext := StrReplace(inputtext, "bekend", "gekend", CaseSensitive := false)
	inputtext := StrReplace(inputtext, ": in het kader van de gekende", ": gekende") 		
	inputtext := StrReplace(inputtext, "in het kader van", "door") 		
	inputtext := StrReplace(inputtext, "foraminaal spinaal stenose", "foraminaal- of spinaalstenose") 		
	inputtext := StrReplace(inputtext, "diffuse restrictie", "diffusie restrictie") 		
	inputtext := StrReplace(inputtext, "normale doorgankelijkheid van de", "normaal doorgankelijke") 		
	inputtext := StrReplace(inputtext, "pig katheter", "PIC katheter") 	
	inputtext := StrReplace(inputtext, "flair", "FLAIR", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "fascikels graad", "Fazekas graad", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "tbc", "TBC")
	inputtext := StrReplace(inputtext, "EKG", "ECG", CaseSensitive := false)		
	inputtext := StrReplace(inputtext, " hili", " hila") 	
	inputtext := StrReplace(inputtext, "longtrauma", "longtrama") 	
	inputtext := RegExReplace(inputtext, "[KkCc]a[mn] configuratie", """cam"" configuratie")
	;; inputtext := StrReplace(inputtext, "bewaarde", "intacte") 
	inputtext := StrReplace(inputtext, "partiële beeld", "partiëel in beeld") 
	inputtext := StrReplace(inputtext, "plaatsen schroef", "plaat en schroef") 
	inputtext := StrReplace(inputtext, "segment VIII", "segment 8")
	inputtext := StrReplace(inputtext, "segment VII", "segment 7")
	inputtext := StrReplace(inputtext, "segment VI", "segment 6", CaseSensitive := true) 		
	inputtext := StrReplace(inputtext, "segment IV", "segment 4") 
	inputtext := RegExReplace(inputtext, "segment V(?=[ ,\.])", "segment 5") 	
	inputtext := StrReplace(inputtext, "segment III", "segment 3") 		
	inputtext := StrReplace(inputtext, "segment II", "segment 2") 	
	inputtext := RegExReplace(inputtext, "segment I(?=[ ,\.])", "segment 1") 	
	inputtext := RegExReplace(inputtext, "\RGECOMMUNICEERDE DRINGENDE BEVINDINGEN:(?=\R|$)", "") 	; verwijderd die zin
	inputtext := RegExReplace(inputtext, "im)^Vergelijking met ", "Vergeleken met ")		
	; inputtext := RegExReplace(inputtext, "i)gekende?", "\#\#\#")				
	inputtext := RegExReplace(inputtext, "im)[\ \t]*supervis.*$", "") ; verwijderd supervisie.

	inputtext := RegExReplace(inputtext, "m)^ *\.?-? *(.+)\/ ?(?=\R|$)", "  . $1") ;; Alle zinnen met / op einde krijgen " . " er voor
	inputtext := RegExReplace(inputtext, "(?<=^|[\n\r])\*\s?(.+?):? ?(?=\R)", "* $U1:")			; adds : at end of string with * and makes uppercase. Not done with m) because of strange bug where it would only capture the first
	inputtext := RegExReplace(inputtext, "m)([\w\d\)\%\°])\ ?(?=\R|$)", "$1.")				; adds . to end of string, word, digit or )
	inputtext := RegExReplace(inputtext, "m)(?<=\. |^- |^)(\w)", "$U1") 					; converts to uppercase after ., newline or newline -
	inputtext := RegExReplace(inputtext, "(?<=:)\ ?([A-Z][^A-Z])", " $L1")					; converts after : to lowercase (escept if 2x capital letter) for eg. DD, FLAIR, ...
	inputtext := RegExReplace(inputtext, "(?<=[\-\/\ ])[D](?=[1-9](?:[0-2]|[\ \:\ ]))", "T") ; Corrects -T10 of /D10 naar -Th10
	;;inputtext := RegExReplace(inputtext, "(?<=[\-\/])[DT](?=[1-9](?:[0-2]|[\ \:]))", "Th") ; Corrects -T10 of /D10 naar -Th10
	;;inputtext := RegExReplace(inputtext, "(?<=\ )[DT](?=[1-9][0-2]?[\-\/])", "Th") ; Corrects T10- naar Th10
	;;inputtext := RegExReplace(inputtext, "([CThD]\d{1,2}[\/-])[TD](?=\d{1,2})", "$1Th") 			; corrects T1/X to Th1 TODO: werkt niet T11-L3
	;;inputtext := RegExReplace(inputtext, "[TD](?=\d{1,2}[\/-][ThDL]{1,2}\d{1,2})", "Th") 			; corrects X/T1 to Th1
	inputtext := RegExReplace(inputtext, "((?:C|Th|L|S)\d{1,2})\/((?:C|Th|L|S)\d{1,2})", "$1-$2") 	; corrects L1/L2 to L1-L2
	inputtext := RegExReplace(inputtext, "(\d{1,2})\/(\d{1,2})\/(\d{2,4})", "$1-$2-$3") 			; corrects d/m/y tot d-m-y
	inputtext := RegExReplace(inputtext, "\R{3,}", "`n`n") 											; replaces triple+ newline with double
	inputtext := RegExReplace(inputtext, "im)^\-?(?<=\-)?(?=\w|\(|\"")(?!CONCLUSIE|Vergeleken|Mede in|In (?:vergel|vgl)|NB|Nota|Storende|Suboptim|Opname in|Reserve|Naar [lr])", "- ")	; adds - to all words and (, excluding BESLUIT, vergeleken...
	inputtext := RegExReplace(inputtext, "(\d )a( \d)", "$1à$2") 						; maakt à als a tussen 2 getallen.
	inputtext := RegExReplace(inputtext, "(\-\.) {2,}", "$1 ")  ; zorgt dat er niet meer dan 1 spatie na een streepje komt
	return inputtext
}

cleanReport_KWS() {
	; Actieve querry zou aberrante medicatieschemas moeten vinden en fixen"
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg)[\s\S]+(?:ONDERZOEKE?N?:\R{1,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<content>[\s\S]+?)(?:\R*$|[\n\r]{2,}\*\* Eind)"
	_KWS_CopyReportToClipboard(True)
	RegExMatch(clipboard, RegexQuerry, report)	
	_KWS_PasteToReport(reportheader . "`n" . cleanreport(reportcontent))
}

copyLastReport_KWS() {
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).+)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"
	_KWS_CopyReportToClipboard(selectReportBox := True)
	_BlockUserInput(True)
	Send {Down} ;; deselecteert de inhoud
	currentreportunclean := clipboard
	clipboard := ""
	;; TODO: evt via "toon laatst gelijktijdig verslag". De titel veranderd dan naar "Dialoog"
	_KWS_SelectReportBox("right")
	Send {Down}					; klik pijltje naar beneden
	Send {Down}
	Send {Enter}	    				; selecteerd "neem laatste verslag over"
	Sleep, 400					; geeft tijd om vorig verslag te laden, kan evt verhoogd of verlaagd worden (400 werkt sowieso)
	_BlockUserInput(false)
 	try _KWS_CopyReportToClipboard(selectReportBox := False)
	catch e {
		_makeSplashText(title := "ERROR", text := "Probleem met het voorgaande verslag te kopiëren!", time := -3000)
		_log("Catch clause regel 113, poging tot _KWS_copyreporttoclipboard")
		_log("clipboard: `n" . clipboard)
		Send {Enter}
		return
	}

	if (RegExMatch(clipboard, "m)^ingevoerde beelden$" )) { ; undoes the whole operation
		Send, ^z
		Send, ^z
		Send ^{F8}					; Initieer dictee (ctrl F8)
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\samenvattingLabel.png
		if (ErrorLevel = 0) { ;; verwijderd "ingevoerde beelden" onderaan
			MouseClick, left, FoundX+20, FoundY+30
			Send ^a
			Send {Backspace}
		}
		return
	} 
	 
	oldreportunclean := clipboard			; zet variabele gelijk aan clipboard
	clipboard := ""             			; maakt het clipboard leeg
	RegExMatch(oldreportunclean, RegexQuerry, oldreport)
	RegExMatch(currentreportunclean, RegexQuerry, currentreport)
	compared := ""
	;; Checkt of de regex van het vorige verslag gelukt is, en zo niet verwijderd het hele gedoe.
	if (oldreportcontent = "") {
		_log("Catch clause regel 143, error bij regexmatch")
		_log("old report: `n" . oldreportunclean)
		_makeSplashText(title := "ERROR", text := "Probleem met REGEX van het clipboard van het vorige verslag: sluit de patient en probeer opnieuw.", time := -2000)
		Send, ^z
		Send, ^z
		Send ^{F8}					; Initieer dictee (ctrl F8)
	}
	;; Veranderd de datum in een reeds bestaande vergelijking, of voegt de vergelijktekst toe indien die nog niet aanwezig was.
	if (oldreportcomparedwith && oldreportcompdate) {
		compared := StrReplace(oldreportcomparedwith, oldreportcompdate, oldreportdate)
		oldreportcontent :=  StrReplace(oldreportcontent, oldreportcompdate, oldreportdate)
	} else {
		compared := "In vergelijking met het voorgaande onderzoek van " . oldreportdate . ":" 
	}

	; zoekt naar het besluit, en als het het vindt voegt het "vergelijking met" toe of veranderdt het de datum van "in vergelijking met"
	conclusielocatie := RegExMatch(oldreportcontent, "(BESLUIT|CONCLUSIE).*[\n\r]", conclusieText)
	if (conclusielocatie) {
		RegExMatch(oldreportcontent, "(?:ergel(?:ij|e)k|opzichte|vgl |tov ).{5,55}?(?<date>\d+[-\/.]\d+[-\/.]\d+)",conclcompare, conclusielocatie - 5)
		if (conclcomparedate) {
			oldreportcontent := StrReplace(oldreportcontent, conclcomparedate, oldreportdate)
		} else {
			oldreportcontent := RegExReplace(oldreportcontent, "((?:BESLUIT|CONCLUSIE).*)\R", "$1`nIn vergelijking met het voorgaande onderzoek van " . oldreportdate . ":`n", ,1, conclusielocatie-5)
		}
	}
	if (SubStr(currentreportheader, StrLen(currentreportheader)) != "`n") { ; Fix dat soms de "compare" tegen de onderzoeken wordt gezet. eventueel te fixen in de REGEX of gewoon extra newlines hier vanonder en dan de overschot newlines wegdoen
		currentreportheader .= "`n"
	}
	if (InStr(currentreporttype, "RX thorax")) {
		oldreportcontent := cleanreport(oldreportcontent) ;; automatisch maakt het verslag proper als het een RX thorax is.
	}
	oldreportcontent := RegExReplace(oldreportcontent, "im)[\ \t]*supervis.*$", "") ; verwijder supervisie.
	_KWS_PasteToReport(currentreportheader . "`n" . compared . "`n`n" . oldreportcontent)
	Send ^{F8}							; Initieer dictee (ctrl F8)
	MouseMove, mouseX, mouseY
	return								; Klaar
}

validateAndClose_KWS() {
	global splashExists
	if (splashExists == "" or splashExists == False) {
		_makeSplashText("Validate function", "Press the button again to validate and close.", time := -3000, doublePressMode := True)
		Send, ^s
		return
	} 
	_destroySplash()
	WinActivate, KWS ahk_exe javaw.exe 
	MouseGetPos, mouseX, mouseY
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\valideersluit.png
	if (ErrorLevel >= 1) {
		_makeSplashText("ERROR valideerfunctie", "ERROR: valideerknop niet gevonden, of er is iets mis gegaan met het zoeken.", -2000)
		return
	} 
	ead := _getEAD()
	MouseClick, left, FoundX+5, FoundY+5
	MouseMove, mouseX, mouseY
	_log(ead, "Gevalideerd en gesloten")
	_makeSplashText(title := "Gevalideerd", text := "Gevalideerd.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000)
}

saveAndClose_KWS() {
	global splashExists
	WinActivate, KWS ahk_exe javaw.exe 
	if (splashExists == "" or splashExists == False) {
		_makeSplashText("Save function", "Press save button again to close.", -3000, doublePressMode := True)
		Send, ^s
		return
	}
	_destroySplash()
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bewarenGreyedButton.png
	if (ErrorLevel = 2) {
		_makeSplashText(title := "Save function", text := "Something went wrong when checking if it was alreade saved", time := -2000)
	} else if (ErrorLevel = 1) {
		_makeSplashText(title := "Save function", text := "Try again, report was not saved", time := -2000)
		Send, ^s
	} else {
		MouseGetPos, mouseX, mouseY
		ead := _getEAD(False)
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\crossActive.png
		if (ErrorLevel = 2) {
			_makeSplashText(title := "Save function", text := "Something went wrong looking for the close button", time := -2000)
		} else if (ErrorLevel = 1) {
			_makeSplashText(title := "Save function", text := "Try again, close button not found", time := -2000)
			Send, ^s
		}
		MouseClick, left, FoundX + 3, FoundY + 3
		MouseMove, mouseX, mouseY
		_log(ead, "Opgeslagen en gesloten")
		_makeSplashText(title := "Opgeslagen", text := "Opgeslagen.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000)
	}
}

closeWithoutSaving() {
	MouseGetPos, mouseX, mouseY
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\crossActive.png
	if (ErrorLevel = 2) {
		_makeSplashText(title := "Save function", text := "Something went wrong looking for the close button", time := -2000)
	} else if (ErrorLevel = 1) {
		_makeSplashText(title := "Save function", text := "Try again, report was not saved close button not found", time := -2000)
		Send, ^s
	}
	_BlockUserInput(True)
	MouseClick, left, FoundX + 3, FoundY + 3
	sleep, 100
	Send {Enter}
	MouseMove, mouseX, mouseY
	_BlockUserInput(false)
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
	WinActivate, KWS ahk_exe javaw.exe 
	sleep, 50
	return _KWS_PasteToReport(result, false)               	;--returning value
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

openEAD_KWS(input := "") {
	; if no EAD number provided, tries to get it from clipboard or get it from selected text
	if (RegExMatch(input, "\D(\d{8})\D", ead) == 0 and RegExMatch(clipboard, "\D(\d{8})\D", ead) == 0) { 
		Clipboard := ""
		Send, ^c
		ClipWait, 1
		if (RegExMatch(clipboard, "\D(\d{8})\D", ead))
			openEAD_KWS(ead1)
		return
	}
	clipboard := ead1
	if WinExist("KWS ahk_exe javaw.exe") {
		WinActivate
		Send, ^+z ; Hotkey voor zoek patient
		Sleep, 400
		Send, ^v
		Send, {Enter}
	}
	return
}

openLastPtInLog_KWS() {
	global logfile
	Loop, read, %logfile%
		lastline := A_loopreadline
	openEAD_KWS(lastline)
	return
}

pedAbdomenTemplate() { ; Gemaakt door Johannes Devos, aangepast en opgekuist door CV.
	;; ptdata := _KWS_GetDemographicDataPatient()
	result := ""
	birthdate := _getBirthDate(returnMouse := true)
	birthdate := RegExReplace(birthdate, "(\d{2}).(\d{2}).(\d{4})", "$3$2$1") 											; replaces triple+ newline with double
	FormatTime, now, ,yyyyMMddHHmmss 
	;; age := CalcAge(birthdate . "000000", now)
	age := _CalcAge(birthdate, now)
	year := age[1]
	months := age[2]
	days := age[3]
	global milt, linkerNier, lever, rechterNier

	Gui, pedAbdGui:+LastFound
	GuiHWND := WinExist()
	Gui, pedAbdGui:Add, Text, x2 y-1 w140 h20 , Patiëntennaam: 
	Gui, pedAbdGui:Add, Text, x2 y19 w140 h20 , Leeftijd:
	Gui, pedAbdGui:Add, Text, x42 y39 w100 h20 , Jaar:
	Gui, pedAbdGui:Add, Text, x42 y59 w100 h20 , Maand:
	Gui, pedAbdGui:Add, Text, x42 y79 w100 h20 , Dag:
	Gui, pedAbdGui:Add, Text, x142 y-1 w230 h20 , 
	Gui, pedAbdGui:Add, Text, x142 y39 w230 h20 , % age[1]
	Gui, pedAbdGui:Add, Text, x142 y59 w230 h20 , % age[2]
	Gui, pedAbdGui:Add, Text, x142 y79 w230 h20 , % age[3]
	Gui, pedAbdGui:Add, Text, x2 y109 w130 h20 , Miltspan (mm):
	Gui, pedAbdGui:Add, Text, x2 y129 w130 h20 , Linkernier (mm):
	Gui, pedAbdGui:Add, Text, x2 y149 w130 h20 , Leverspan (mm):
	Gui, pedAbdGui:Add, Text, x2 y169 w130 h20 , Rechternier (mm):
	Gui, pedAbdGui:Add, Edit, x132 y109 w100 h20 vmilt, 0
	Gui, pedAbdGui:Add, Edit, x132 y129 w100 h20 vlinkerNier, 0
	Gui, pedAbdGui:Add, Edit, x132 y149 w100 h20 vlever, 0
	Gui, pedAbdGui:Add, Edit, x132 y169 w100 h20 vrechterNier, 0
	Gui, pedAbdGui:Add, Button, x2 y199 w370 h30 Default, OK
	Gui, pedAbdGui:Show, x759 y391 h236 w379, Echografie Pediatrie Afmetingen
	WinWaitClose, ahk_id %GuiHWND%  		; waiting for gui to close
	WinActivate, KWS ahk_exe javaw.exe 
	if (result != "")
		_KWS_PasteToReport(result, false)       	; returning value
	return 
; --------
pedAbdGuiButtonOK:
	Gui, Submit
	result := _makePedReport(age, milt, linkernier, lever, rechternier)
	Gui, pedAbdGui:Destroy
	return
; -------
pedAbdGuiEscape:
pedAbdGuiClose:
	result := ""
	Gui, pedAbdGui:Destroy
	return
}

pressOKButton() { ;; not really used, might be used in the future.
	suspend permit
	MouseGetPos, mouseX, mouseY
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\okButton.png
	if (ErrorLevel = 2)
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000)
	else if (ErrorLevel = 1)
		return
	else {
		MouseClick, %mousebutton%, FoundX+5, FoundY+5
		MouseMove, mouseX, mouseY
	}
}

;; WORK IN PROGRESS
aanvaarderMode() {
	Gui, aanvaardGUI:+LastFound +AlwaysOnTop +Owner
	GuiHWND := WinExist()
	Gui, aanvaardGUI:Add , Text  ,        , Aanvaardmodus
	Gui, aanvaardGui:Add , Text , , h - j - l - p
	Gui, aanvaardGUI:Add , Button, Default, OK
	Gui, aanvaardGUI:Show, , Aanvaardmodus
	;; suspend on
	#ifWinExist Aanvaardmodus
	Hotkey, IfWinExist, Aanvaardmodus
	Hotkey, ^Enter, pressOKButton
	Hotkey, ^h, _pressAanvaardOption
	Hotkey, ^j, _pressAanvaardOption
	Hotkey, ^l, _pressAanvaardOption
	Hotkey, ^p, _pressAanvaardOption
	WinWaitClose, ahk_id %GuiHWND%  		;--waiting for gui to close
	WinActivate, KWS ahk_exe javaw.exe 
	return
	;-------
	aanvaardGUIButtonOK:
	aanvaardGUIEscape:
	aanvaardGUIClose:
		suspend off
		Gui, aanvaardGUI:Destroy
	return
}



MoveLineUp() {
	clipboard := ""
	Send, {End}
	Send, +{Home}
	Send, +{Left}
	Send, ^x
	ClipWait, 1
	Send, {Up}
	Send, {End}
	Send, ^v
	Return
}

MoveLineDown() {
	clipboard := ""
	Send, {End}
	Send, +{Home}
	Send, +{Left}
	Send, ^x
	ClipWait, 1
	Send, {Down}
	Send, {End}
	Send, ^v
	Return
}

deleteLine() {
	Send, {Home}
	Send, +{Down}
	Send, {Backspace}
}

; HELPER FUNCTIONS
; --------------------------------------

_KWS_CopyReportToClipboard(selectReportBox := True) {
	_BlockUserInput(true)
	If (not WinActive("KWS ahk_exe javaw.exe")) {
		if WinExist("KWS ahk_exe javaw.exe")
			WinActivate 
		else
			throw Exception("KWS is not open!", -1)
	}
	if (selectReportBox) {
		_KWS_SelectReportBox()
	}
	clipboard := ""             			; maakt het clipboard leeg 
	Send, ^a                    			; select all
	Send, ^c                    			; copy
	_BlockUserInput(false)
	ClipWait, 1                 			; wacht tot er data in het clipboard is
	if (ErrorLevel)             			; als NOT, is er data in clipboard
		throw Exception("Could not copy data to clipboard!", -1)                 				; STOPT als geen data in clipboard
}

_KWS_SelectReportBox(mousebutton := "left") {
	;; assumes KWS already active
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bevindingenLabel.png
	if (ErrorLevel = 2)
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000)
	else if (ErrorLevel = 1)
		_makeSplashText(title := "Error", text := "Error finding the report field: did the action succeed nonetheless?", time := -2000)
	_BlockUserInput(true)
	MouseGetPos, mouseX, mouseY
	MouseClick, %mousebutton%, FoundX+100, FoundY+200
	MouseMove, mouseX, mouseY
	_BlockUserInput(false)
}


_KWS_PasteToReport(text, overwrite := true) {
	If WinActive("KWS ahk_exe javaw.exe") {
		_BlockUserInput(True)
		tempclip := clipboard
		clipboard := ""  
		clipboard := text           	; maakt het clipboard leeg 
		ClipWait, 1			; wacht tot er data in het clipboard is
		if (overwrite) {
			Send, ^a
		}
		Send, ^v 
		Sleep 50
		;; TODO !!! Werkt niet meer.... evt via findwindow??
		if WinExist("Foutboodschap JavaKWS") { ; Fixes "could not access clipboard"
			WinActivate
			SendInput, {Enter}
			Sleep, 50
			WinActivate, KWS ahk_exe javaw.exe 
			_KWS_PasteToReport(text, overwrite)
		}
		clipboard := tempclip
		_BlockUserInput(false)
		return
	}
	if WinExist("KWS ahk_exe javaw.exe")
		WinActivate 
	else
		throw Exception("KWS is not open!", -1)
	_KWS_SelectReportBox()
	_KWS_PasteToReport(text, overwrite)
	return
}

_sorttext(inputtext) {
	if (RegExMatch(inputtext, "([\s\S]+)(\RCONCLUSIE[\s\S]+)", split)) {
		outputtext := _sorttext(split1) . split2
		return outputtext
	} else if (RegExMatch(inputtext, "([\s\S]+)(\*.+)([\s\S]+)", split)) {
		outputtext := _sorttext(split1) . split2 . _sorttext(split3)
		return outputtext
	} else {
		dotlines := ""
		;; TODO: evt via Loop, Parse, inputtext, "`n`r" en dan A_LoopField
		while (RegExMatch(inputtext, "m)(?:\R|^)\.?\ ?(.+#.?)$" , line)) { 	;gets all lines with XXXXX #
			dotlines := dotlines . "`n. " . line1
			inputtext := StrReplace(inputtext, line , "")
		}
		while (RegExMatch(inputtext, "m)(?:\R|^)\.\ ?(.+[^\/])?$" , line)) { 		;gets all lines with . XXXXX (as long as not ending with /)
		; while (RegExMatch(inputtext, "m)(?:\R|^)\.\ ?(.+)$" , line)) { 		;gets all lines with . XXXXX
			dotlines := dotlines . "`n. " . line1
			inputtext := StrReplace(inputtext, line , "")
		}
		outputtext := inputtext . "`n" . StrReplace(StrReplace(dotlines, "#" , ""), ". - ", ". ") . "`n`n"
		return outputtext
	}
}

_pressAanvaardOption() {
	suspend Permit
	MouseGetPos, mouseX, mouseY
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\contrastLabel.png
	if (ErrorLevel = 2)
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000)
	else if (ErrorLevel = 1)
		return
	else {
		FoundY := FoundY+5
		if (GetKeyState("h","P")) { ;; zonder
			FoundX := FoundX + 80
		} else if (GetKeyState("j","P")) { ;; Met
			FoundX := FoundX + 165
		} else if (GetKeyState("l","P")) { ;; Zonder/met
			FoundX := FoundX + 250
		} else if (GetKeyState("p","P")) { ;; textvak
			FoundX := FoundX + 250
			FoundY := FoundY + 200
		}
		MouseClick, left, FoundX, FoundY
		MouseMove, mouseX, mouseY
	}
}

_getEAD(returnMouse := false) {

	if (returnMouse) {
		MouseGetPos, mouseX, mouseY
	}
	clipboard := temp
	clipboard := ""
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\eadnrLabel.png
	_BlockUserInput(true)
	Mouseclick, left, FoundX+70, FoundY+10
	_BlockUserInput(false)
	Clipwait, 1
	ead := clipboard
	clipboard := temp
	if (returnMouse) {
		MouseMove, mouseX, mouseY
	}
	return ead
}

_getBirthDate(returnMouse := false) {
	if (returnMouse) {
		MouseGetPos, mouseX, mouseY
	}
	clipboard := temp
	clipboard := ""
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\eadnrLabel.png
	_BlockUserInput(True)
	Mouseclick, left, FoundX-25, FoundY+10
	Clipwait, 1
	date := SubStr(clipboard, 2)
	clipboard := temp
	if (returnMouse) {
		MouseMove, mouseX, mouseY
	}
	_BlockUserInput(false)
	return date
}

_makeSplashText(title := "Splash title", text := "Splash text", time := -3000, doublePressMode := false) {
	time := Abs(time) * (-1) ;; Zorgt dat time altijd negatief is (dat removesplash dus maar 1 keer wordt uitgevoerd ipv in loop)
	global splashExists
	_destroySplash()
	splashExists := doublePressMode
	Gui, splashGui:+AlwaysOnTop +Disabled -SysMenu +Owner  	; +Owner avoids a taskbar button.
	Gui, splashGui:Add, Text,, %text%
	Gui, splashGui:Show, NoActivate, %title%		; NoActivate prevents taking the focus
	destroySplash_fn := Func("_destroySplash").bind() ; Settimer aanvaard enkel een label, iets wat ik probeer te vermijden. Op deze manier kan ik toch een functie aan setTimer geven.
	SetTimer, % destroySplash_fn, %time%
	return
}

_destroySplash() {
	global	splashExists
	splashExists := False
	Gui, splashGui:Destroy
}

_calcHeightLoss(h1, h2) {
	absolute := Max(h1,h2) - Min(h1,h2)
	percentage := Round((1 - (Min(h1,h2) / Max(h1,h2))) * 100)
	return [absolute, percentage]
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

_KWS_closeMededelingen() {
	if Winexist("Mededelingen ahk_exe javaw.exe") {
		Winclose
	}
}

clipboardcleaner() {
	ClipWait, 1                 			; wacht tot er data in het clipboard is
	sleep 50
	Foundpos := InStr(clipboard, "uit ongevalideerd verslag")
	if (Foundpos) {
		FoundPos := RegExMatch(clipboard, "[\n\r]",,FoundPos)
		clipboard := SubStr(clipboard, FoundPos + 4)
		FoundPos2 := Instr(clipboard, "** Einde tekst uit ongevalideerd verslag **")
		if (FoundPos2)
			clipboard := SubStr(clipboard, 1, FoundPos2 - 5)
		Return
	} else {
		Return
	}
}

_BlockUserInput(block := true) {
	global blockinput
	blockinput := block
	if (blockinput) {
		MouseClick, left, 0, 0,,, U, R ;; zorgt dat als de muis ingeduwd was dat er geen error komt
		BlockInput, Mousemove   ; fragile sequence of events, therefore block mouse movement to prevent user interference.
		BlockInput, Send
	}
	if (not blockinput) {
		BlockInput, Default
		BlockInput, MouseMoveOff
	}
	;; Settimer die de blokkage opheft na 1 seconde, voor moest het programma crashen of vastlopen
	blockInput_fn := Func("_BlockUserInput").bind(false) ; Settimer aanvaard enkel een label, iets wat ik probeer te vermijden. Op deze manier kan ik toch een functie aan setTimer geven.
	SetTimer, % blockInput_fn, -1000
}

_blockInputHelper() {
	;; lege functie, wordt opgeroepen elke keer als een geblokte input wordt opgeroepen.
	return
}

_MouseIsOver(vWinTitle:="", vWinText:="", vExcludeTitle:="", vExcludeText:="") {
	MouseGetPos,,, hWnd
	return WinExist(vWinTitle (vWinTitle=""?"":" ") "ahk_id " hWnd, vWinText, vExcludeTitle, vExcludeText)
}

_KWS_GetDemographicDataPatient() {
	WinGetTitle, VensterTitel, A
	RegExMatch(VensterTitel, "(\d{6})([M,V,G,B])\d{3}, (\d{1,3})([jmd])", EMD)
	if (EMD = "") {
		_makeSplashText("Error", "Geboortedatum werd niet gevonden.", 2000, false)
		return EMD
	}
	;Extract de leeftijd
	if ((EMD1 < 500000) and (EMD3 < 50)) { ; tegen 2050 zal dit moeten worden aangepast, niet meer mijn probleem dan.
		leeftijd := "20" . EMD1
	} else {
		leeftijd := "19" . EMD1
	}
	if (EMD2 = "M" or EMD3 = "B") { ; eigenlijk niet nodig
		geslacht := "m"
	} else {
		geslacht := "f"
	}

	FormatTime, now, ,yyyyMMdd 
	dataArray := _CalcAge(leeftijd, now) ; [Years, months, days]
	; Extract de naam
	RegExMatch(VensterTitel, "- (.*) \(", naam)
	dataArray.push(geslacht)
	dataArray.push(naam1)
	return dataArray ; [years, months, days, geslacht, naam]
}

_makePedReport(age, Milt, LinkerNier, Lever, RechterNier) { ; Gemaakt door Johannes Devos, aangepast
	Gemiddelde_Nieren := [4.48, 5.28, 6.15, 6.23, 6.65, 7.36, 7.36, 7.87, 8.09, 7.83, 8.33, 8.9, 9.2, 9.17, 9.6, 10.42, 9.79, 10.05, 10.93, 10.04, 10.53, 10.81]
	SD_Nieren := [0.31, 0.66, 0.67, 0.63, 0.54, 0.54, 0.64, 0.5, 0.54, 0.72, 0.51, 0.88, 0.9, 0.82, 0.64, 0.87, 0.75, 0.62, 0.76, 0.86, 0.29, 1.13]
	Gemiddelde_Milt := [53, 59, 63, 70, 75, 84, 85, 86, 97, 101, 101]
	SD_Milt := [7.8, 6.3, 7.6, 9.6, 8.4, 9, 10.5, 10.7, 9.7, 11.7, 10.3] ;Lengte 11
	Gemiddelde_Lever := [64, 73, 79, 85, 86, 100, 105, 105, 115, 118, 121]
	SD_Lever := [10.4, 10.8, 8, 10, 11.8, 13.6, 10.6, 12.5, 14, 14.6, 11.7]
	Result := [0,0,0,0]
	SetFormat, Float, 0.2

	Years := age[1]
	Months := age[2]
	Days := age[3]

	; NIEREN
	if (Years >= 1){
		Index := Years + 4
		if (Years >= 19){
			Index := 22
		}
		Gemiddelde := Gemiddelde_Nieren[Index]
		SD := SD_Nieren[Index]
		Result[1] := (LinkerNier/10 - Gemiddelde)/SD
		Result[2] := (RechterNier/10 - Gemiddelde)/SD
	}
	else if (Months >= 1 or Days > 7){
		Index := Months / 4 + 2
		Index := Floor(Index)
		Gemiddelde := Gemiddelde_Nieren[Index]
		SD := SD_Nieren[Index]
		Result[1] := (LinkerNier/10 - Gemiddelde)/SD
		Result[2] := (RechterNier/10 - Gemiddelde)/SD
	}
	else {
		Gemiddelde := Gemiddelde_Nieren[1]
		SD := Gemiddelde_Nieren[1]
		Result[1] := (LinkerNier/10 - Gemiddelde)/SD
		Result[2] := (RechterNier/10 - Gemiddelde)/SD
	}
	; MILT & LEVER
	if (Years = 0){
		Index := Floor((Months-1)/3) + 1
		if (Months = 0)
			Index := 1
		if (Months > 9)
			Index := 3
		Result[3] := (Milt - Gemiddelde_Milt[Index])/SD_Milt[Index]
		Result[4] := (Lever - Gemiddelde_Lever[Index])/SD_Lever[Index]
	}
	else if (Years < 3){
		Result[3] := (Milt - Gemiddelde_Milt[4])/SD_Milt[4]
		Result[4] := (Lever - Gemiddelde_Lever[4])/SD_Lever[4]
	}
	else if (Years < 6){
		Result[3] := (Milt - Gemiddelde_Milt[5])/SD_Milt[5]
		Result[4] := (Lever - Gemiddelde_Lever[5])/SD_Lever[5]
	}
	else if (Years < 7){
		Result[3] := (Milt - Gemiddelde_Milt[6])/SD_Milt[6]
		Result[4] := (Lever - Gemiddelde_Lever[6])/SD_Lever[6]
	}
	else if (Years < 9){
		Result[3] := (Milt - Gemiddelde_Milt[7])/SD_Milt[7]
		Result[4] := (Lever - Gemiddelde_Lever[7])/SD_Lever[7]
	}
	else if (Years < 11){
		Result[3] := (Milt - Gemiddelde_Milt[8])/SD_Milt[8]
		Result[4] := (Lever - Gemiddelde_Lever[8])/SD_Lever[8]
	}
	else if (Years < 13){
		Result[3] := (Milt - Gemiddelde_Milt[9])/SD_Milt[9]
		Result[4] := (Lever - Gemiddelde_Lever[9])/SD_Lever[9]
	}
	else if (Years < 15){
		Result[3] := (Milt - Gemiddelde_Milt[10])/SD_Milt[10]
		Result[4] := (Lever - Gemiddelde_Lever[10])/SD_Lever[10]
	}
	else {
		Result[3] := (Milt - Gemiddelde_Milt[11])/SD_Milt[11]
		Result[4] := (Lever - Gemiddelde_Lever[11])/SD_Lever[11]
	}
	SetFormat, Float, 0.1
	Verslag := "Normale ligging van de retroperitoneale grote vaten.`nNormale ligging van de organen.`n`nLeverspan: " . Lever/10 . " cm (SD: " . Result[4] . ").`nHomogeen leverparenchym met normale reflectiviteit.`nNormale portahoofdstam en intrahepatische portatakken.`nNormale hepatische venen met normale hepatofugale flow.`nNormale hepatopetale portale flow.`nNormale flow in de a. hepatica.`nGeen gedilateerde intrahepatische of extrahepatische galwegen aangetoond.`nNormale galblaas.`nNormale pancreas. Geen visualisatie van de ductus van Wirsung.`nMilt: " . Milt/10 . " cm (SD: " . Result[3] . ").`nNormale milt.`n`nNormale bijnieren en bijnierloges.`nLinkernier: " . Linkernier/10 . " cm (SD: " . Result[1] . ").`nRechternier: " . Rechternier/10 . " cm (SD: " . Result[2] . ").`nNormale reflectiviteit van het nierparenchym met corticomedullaire differentiatie.`nGeen hydro-ureteronefrose.`nNormale blaasvulling.`nNormale aflijning en dikte van de blaaswand.`n`nNormale ligging van de A. en V. Mesenterica Superior.`nGeen adenopathieën aangetoond.`nNormale darmwanden.`n###Normaal terminale ileum.`n###Normale appendix.`n`nCONCLUSIE:`n###`n`nGECOMMUNICEERDE DRINGENDE BEVINDINGEN:`n"
	return Verslag
}

_CalcAge(FromDay,ToDay) {   ;Age calculation function
	FromDay := substr(FromDay,1,8)
	ToDay := Substr(ToDay,1,8)
	Years := 0
	Months := 0
	Days := 0
	;; Global Years,Months,Days
	; If born on February 29
	If SubStr(FromDay,5,4) = 0229 and Mod(SubStr(ToDay,1,4), 4) != 0 and SubStr(ToDay,5,4) = 0228
		PlusOne = 1
	ThisMonth := SubStr(ToDay,1,6)
	; Set ThisMonthLength equal to next month
	ThisMonthLength := % SubStr(ToDay,5,2) = "12" ? SubStr(ToDay,1,4)+1 . "01"
	: SubStr(ToDay,1,4) . Substr("0" . SubStr(ToDay,5,2)+1,-1)
	; Days in this month saved in ThisMonthLength
	EnvSub, ThisMonthLength, %ThisMonth%, d
	; Set ThisMonthday to FromDay or  (if FromDay higher) last day of this month
	If SubStr(FromDay,7,2) > ThisMonthLength
		ThisMonthDay :=  ThisMonth . ThisMonthLength
	Else
		ThisMonthDay :=  ThisMonth . SubStr(FromDay,7,2)
	; Calculate last month's length
	LastMonthLength := % SubStr(ToDay,5,2) = "01" ? SubStr(ToDay,1,4)-1 . "12"
	: SubStr(ToDay,1,4) . Substr("0" . SubStr(ToDay,5,2)-1,-1)
	LastMonth := LastMonthLength
	; Days in last month saved in LastMonthLength
	EnvSub, LastMonthLength, %ThisMonth% ,d
	LastMonthLength := LastMonthLength*(-1)
	; Set LastMonthday to FromDay or (if FromDay higher) last day of last month
	If SubStr(FromDay,7,2) > LastMonthLength
		LastMonthDay :=  LastMonth . LastMonthLength
	Else
		LastMonthDay :=  LastMonth . SubStr(FromDay,7,2)
	; Calculate years
	Years  := % SubStr(ToDay,5,4) - SubStr(FromDay,5,4) < 0 ? SubStr(ToDay,1,4)-SubStr(FromDay,1,4)-1
	: SubStr(ToDay,1,4)-SubStr(FromDay,1,4)
	; Calculate months
	Months := % SubStr(ToDay,5,2)-SubStr(FromDay,5,2) < 0 ? SubStr(ToDay,5,2)-SubStr(FromDay,5,2)+12
	: SubStr(ToDay,5,2)-SubStr(FromDay,5,2)
	Months := % SubStr(ToDay,7,2) - SubStr(ThisMonthDay,7,2) < 0 ? Months -1 : Months
	Months := % Months = -1 ? 11 : Months
	; Calculate days
	TodayDate := SubStr(ToDay,1,8)          ; Remove any time portion of stamp
	EnvSub, ThisMonthDay,ToDayDate , d
	EnvSub, LastMonthDay,ToDayDate , d
	Days  := % ThisMonthDay <= 0 ? -1*ThisMonthDay : -1*LastMonthDay
	; If February 28
	Years := % plusone = 1 ? Years +1 : Years
	days := % plusone = 1 ? 0 : days
	If (TodayDate <= FromDay)
		Years := 0, Months := 0,Days := 0
	age := [Years, Months, days]
	return age
}
