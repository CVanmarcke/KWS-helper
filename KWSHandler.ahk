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

; TODO: pedabdomen nog eens testen of de stddev kloppen.
; TODO: functie om automatisch alle schermen voor spoed te openen.


initKWSHandler() {
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 1
	Coordmode, Pixel
	Coordmode, Mouse

	global logfile
	logfile := "logfile.csv"

	;; Verwijderd de Teams cache folder: die neemt soms meer dan een GB aan data in zonder reden.
	FileRemoveDir, P:\uzlsystem\AppData\Microsoft\Teams\Service Worker\CacheStorage, 1

	_makeSplashText(title := "KWS-helper", text := "Started KWS-helper", time := -3000)

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
	inputtext := RegExReplace(inputtext, "m)^\: ", ". ")				; replaces : if at the front of the sentence with (speechfout).
	inputtext := RegExReplace(inputtext, "im)^(besluit|conclusie)", "CONCLUSIE")				; replaces case insensitive besluit/conclusie door upper
	inputtext := RegExReplace(inputtext, "im)^punt ", ". ")	; corrigeert speech fout dat het punt typt ipv punt (enkel in het begin van de zin)
	inputtext := RegExReplace(inputtext, "m)^ *\.?-? *(.+)\/ ?(?=\R|$)", "  . $1")                          ;; Alle zinnen met / op einde krijgen " . " er voor
	if (RegExMatch(inputtext, "m)^\. ?.+(?:\R|$)") OR RegExMatch(inputtext, "m)^.+# ?.?(?:\R|$)")) { 	; only executes if there is ". " or "#" in the script
		inputtext := _sorttext(inputtext) ; zet alle zinnen met een punt vooraan, onder het verslag.
	}
	sleep, 50
	inputtext := StrReplace(inputtext, "bekend", "gekend", CaseSensitive := false)
	inputtext := StrReplace(inputtext, "formaliteit", "voor maligniteit")
	inputtext := StrReplace(inputtext, ": in het kader van de gekende", ": gekende") 		
	inputtext := StrReplace(inputtext, "in het kader van", "door") 		
	inputtext := StrReplace(inputtext, "ongewijzigd", "onveranderd", CaseSensitive := false)
	inputtext := StrReplace(inputtext, "foraminaal spinaal stenose", "foraminaal- of spinaalstenose") 		 ;; frequente speech fout
	inputtext := StrReplace(inputtext, "diffuse restrictie", "diffusie restrictie") 		;; frequente speech fout
	inputtext := StrReplace(inputtext, "normale doorgankelijkheid van de", "normaal doorgankelijke") 		
	inputtext := StrReplace(inputtext, "pig katheter", "PIC katheter") 	
	inputtext := StrReplace(inputtext, "flair ", "FLAIR ", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "fascikels graad", "Fazekas graad", CaseSensitive := false) 	
	inputtext := StrReplace(inputtext, "tbc", "TBC")
	inputtext := StrReplace(inputtext, " EKG ", " ECG ", CaseSensitive := false)
	inputtext := StrReplace(inputtext, " ecg ", " ECG ", CaseSensitive := false)
	inputtext := StrReplace(inputtext, " hili", " hila") 	
	inputtext := StrReplace(inputtext, "longtrauma", "longtrama") 	
	inputtext := StrReplace(inputtext, "op niveau van", "aan") 	
	inputtext := RegExReplace(inputtext, "[KkCc]a[mn] configuratie", """cam"" configuratie")
	;; inputtext := StrReplace(inputtext, "bewaarde", "intacte") 
	inputtext := StrReplace(inputtext, "partiële beeld", "partiëel in beeld") 
	inputtext := StrReplace(inputtext, "plaatsen schroef", "plaat en schroef") 
	inputtext := StrReplace(inputtext, "rx", "RX", CaseSensitive := false) 		
	inputtext := StrReplace(inputtext, "segment VIII", "segment 8")
	inputtext := StrReplace(inputtext, "segment VII", "segment 7")
	inputtext := StrReplace(inputtext, "segment VI", "segment 6", CaseSensitive := true) 		
	inputtext := StrReplace(inputtext, "segment IV", "segment 4") 
	inputtext := RegExReplace(inputtext, "segment V(?=[ ,\.])", "segment 5") 	
	inputtext := StrReplace(inputtext, "segment III", "segment 3") 		
	inputtext := StrReplace(inputtext, "segment II", "segment 2") 	
	inputtext := RegExReplace(inputtext, "segment I(?=[ ,\.])", "segment 1") 	
	inputtext := RegExReplace(inputtext, "(\d )n?o?r?maal( \d)", "$1x$2") 	; Corrigeert een veel gemaakte fout van de speech

       ;; mammo:
	inputtext := RegExReplace(inputtext, "(eefseltype(?:ring)?)\:? ([a-d])", "$1 $U2") 	; zet het weefseltype in hoofdletters

	inputtext := RegExReplace(inputtext, "[\r\n]GECOMMUNICEERDE DRINGENDE BEVINDINGEN:[\n\r]?$", "") 	; verwijderd die zin
	inputtext := RegExReplace(inputtext, "im)^Vergelijking met ", "Vergeleken met ")		
	; inputtext := RegExReplace(inputtext, "i)gekende?", "\#\#\#")				
	inputtext := RegExReplace(inputtext, "im)[\ \t]*supervis.*$", "") ; verwijderd supervisie.

	inputtext := RegExReplace(inputtext, "m)[\ \t]+$", "")  ; zorgt dat er geen nutteloze spaties op het inde van de zin komen
	;;; inputtext := RegExReplace(inputtext, "([\n\r\.]) +(?=[\n\r])", "$1")  ; zorgt dat er geen spatie achter . of op nieuwe lijn komt
	inputtext := RegExReplace(inputtext, "([A-Z])([A-Z][a-z]{3,})", "$U1$L2") 				; corrigeert WOord naar Woord
	inputtext := RegExReplace(inputtext, "(?<=^|[\n\r])\*\s?(.+?):? ?(?=\R)", "* $U1:")			; adds : at end of string with * and makes uppercase. Not done with m) because of strange bug where it would only capture the first
	inputtext := RegExReplace(inputtext, "m)([\w\d\)\%\°])\ ?(?=\R|$)", "$1.")				; adds . to end of string, word, digit or )
	inputtext := RegExReplace(inputtext, "m)(?<=\. |^- |^)(\w)", "$U1") 					; converts to uppercase after ., newline or newline -
	inputtext := RegExReplace(inputtext, "([a-z])([\:\.])([a-zA-Z])", "$1$2 $3")					; makes sure there is a space after a colon or point (if not number)...
	inputtext := RegExReplace(inputtext, "(?<=:)\ ?([A-Z][^A-Z])", " $L1")					; converts after : to lowercase (escept if 2x capital letter) for eg. DD, FLAIR, ...
;;	inputtext := RegExReplace(inputtext, "(?<=[\-\/\ ])[D](?=[1-9](?:[0-2]|[\ \:\ ]))", "T") ; Corrects -T10 of /D10 naar -Th10
	;;inputtext := RegExReplace(inputtext, "(?<=[\-\/])[DT](?=[1-9](?:[0-2]|[\ \:]))", "Th") ; Corrects -T10 of /D10 naar -Th10
	;;inputtext := RegExReplace(inputtext, "(?<=\ )[DT](?=[1-9][0-2]?[\-\/])", "Th") ; Corrects T10- naar Th10
	;;inputtext := RegExReplace(inputtext, "([CThD]\d{1,2}[\/-])[TD](?=\d{1,2})", "$1Th") 			; corrects T1/X to Th1 TODO: werkt niet T11-L3
	;;inputtext := RegExReplace(inputtext, "[TD](?=\d{1,2}[\/-][ThDL]{1,2}\d{1,2})", "Th") 			; corrects X/T1 to Th1
	inputtext := RegExReplace(inputtext, "((?:C|Th|L|S)\d{1,2})\/((?:C|Th|L|S)\d{1,2})", "$1-$2") 	; corrects L1/L2 to L1-L2
	inputtext := RegExReplace(inputtext, "(\d{1,2})\/(\d{1,2})\/(\d{2,4})", "$1-$2-$3") 			; corrects d/m/y tot d-m-y
	inputtext := RegExReplace(inputtext, "\R{3,}", "`n`n") 											; replaces triple+ newline with double
	inputtext := RegExReplace(inputtext, "im)^\-?(?<=\-)?(?=\w|\(|\"")(?!CONCLUSIE|Vergeleken|Mede in|In (?:vergel|vgl)|NB|Nota|Storende|Suboptim|Opname in|Reserve|Naar [lr]|[PBT]IRADS|\d[\/\)\.])", "- ")	; adds - to all words and (, excluding BESLUIT, vergeleken...
	inputtext := RegExReplace(inputtext, "(CONCLUSIE:[\n\r\R])\-\ (.+)(?:[\n\r\R]|$)(?!-)", "$1$2")	; Als maar 1 lijn conclusie, zal het het streepje weglaten. WERKT NOG NIET
	;;inputtext := RegExReplace(inputtext, "(\d )a( \d)", "$1Ã$2") 						; maakt Ã  als a tussen 2 getallen.
	inputtext := RegExReplace(inputtext, "([\-\.]) {2,}(?=[\R\n\r\w])", "$1 ")  ; zorgt dat er niet meer dan 1 spatie na een streepje komt
	inputtext := RegExReplace(inputtext, "\ +, \ +", ", ")  ; verwijdert te veel spaties rond een komma 
	inputtext := RegExReplace(inputtext, "\x{2013}", "-")  ; veranderd het unicode streepje (– aka \x{2013}) naar een ASCII streepje. Nog niet getest.
	return inputtext
}

cleanReport_KWS() {
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:ONDERZOEKE?N?:\R{1,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<content>[\s\S]+?)(?:\R*$|[\n\r]{2,}\*\* Eind)"
 	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch e {
		return
	}
	RegExMatch(clipboard, RegexQuerry, report)	
	if (not reportheader = "")
		_KWS_PasteToReport(reportheader . "`n" . cleanreport(reportcontent))
}


mergeReport(currentreportunclean, oldreportunclean) {
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).*)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"

	RegExMatch(currentreportunclean, RegexQuerry, currentreport)
	RegExMatch(oldreportunclean, RegexQuerry, oldreport)
	compared := ""
	;; Checkt of de regex van het vorige verslag gelukt is, en zo niet verwijderd het hele gedoe.
	if (oldreportcontent = "") {
		_log("Catch clause regel 143, error bij regexmatch")
		_log("old report: `n" . oldreportunclean)
		_makeSplashText(title := "ERROR", text := "Probleem met REGEX van het clipboard van het vorige verslag: sluit de patient en probeer opnieuw.", time := -2000)
		return ""
	}
	;; Veranderd de datum in een reeds bestaande vergelijking, of voegt de vergelijktekst toe indien die nog niet aanwezig was.
	;; if (oldreportcomparedwith && oldreportcompdate) {
	if (oldreportcomparedwith) {
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
	report := currentreportheader . "`n" . compared . "`n`n" . oldreportcontent
	return report
}


copyLastReport_KWS() {
	tempClip := clipboard
	clipboard := ""
 	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch e {
		return
	}
	currentreportunclean := clipboard
	clipboard := ""

	MouseGetPos, mouseX, mouseY
	_BlockUserInput(True)
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\NieuweMededelingHeader.png ;; klikt eerst de mededeling weg als die er is
	if (FoundX) {
		MouseClick, left, FoundX+400, FoundY+15
		sleep 50
	}
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\toonLaatstVerslagKnop.png
	if (FoundX = "") {
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\toonLaatstVerslagKnopSelected.png
		if (FoundX = "") {
			_makeSplashText(title := "Error", text := "Geen vorig verslag aanwezig!", time := -2000)
			MouseMove, mouseX, mouseY
			return
		}
	}
	MouseClick, left, FoundX+10, FoundY+10 
	sleep, 300

	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\laatstVerslagHeader.png
	if (FoundX = "" and FoundX2 = "") {
		_makeSplashText(title := "Error", text := "Verslag popup niet gevonden!", time := -2000)
		MouseMove, mouseX, mouseY
		return
	}
	MouseClick, left, FoundX+100, FoundY+200
	sleep, 100
	Send, ^a
	Send, {Ctrl down}c{Ctrl up} ;; zou ook reliablity verhogen
	Clipwait, 1
	oldreportunclean := clipboard			; zet variabele gelijk aan clipboard

	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\sluitLaatstVerslagKnop.png
	MouseClick, left, FoundX+5, FoundY+5
	sleep, 100
	MouseClick, left, FoundX+5, FoundY+650 ;; Selecteerd Textbox van KWS
	sleep, 100 ;; even tijd geven
	clipboard := ""
	if (RegExMatch(oldreportunclean, "m)^ingevoerde beelden$" )) { ; undoes the whole operation
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		_BlockUserInput(false)
		return
	} 
	mergedReport := mergeReport(currentreportunclean, oldreportunclean)
	If (not mergedReport = "") ;; Checkt dat het effectief gelukt is om samen te voegen
		_KWS_PasteToReport(mergedReport, true)
	Send ^{F8}							; Initieer dictee (ctrl F8)
	MouseMove, mouseX, mouseY
	_BlockUserInput(false)
	clipboard := tempClip
	return								; Klaar
}

selectLastReport_KWS() {
	RegexQuerry := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).*)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"
 	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch e {
		return
	}
	_BlockUserInput(True)
	Send {Down} ;; deselecteert de inhoud
	currentreportunclean := clipboard
	clipboard := ""
	_KWS_SelectReportBox("right")
	_BlockUserInput(True)
	Send {Down}					; klik pijltje naar beneden
	Send {Down}
	Send {Enter}	    				; selecteerd "neem laatste verslag over"
	if (A_UserName = "cvmarc2") {
		;; Mag enkel uitgevoerd worden als de technische dienst de toegangsactie "laatsteGelijkaardigVerslagDetails" heeft opengezet! Indien niet, is dit niet nodig.
		WinWait, Neem een gelijkaardig verslag over ahk_class SunAwtDialog ahk_exe javaw.exe,, 3
		WinActivate
		_BlockUserInput(false)
		WinWaitClose, Neem een gelijkaardig verslag over ahk_class SunAwtDialog ahk_exe javaw.exe
		Sleep, 200
		WinActivate, KWS ahk_exe javaw.exe
	}
	Sleep, 400					; geeft tijd om vorig verslag te laden, kan evt verhoogd of verlaagd worden (400 werkt sowieso)
 	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch e {
		_makeSplashText(title := "ERROR", text := "Probleem met het voorgaande verslag te kopieren!", time := -3000)
		_log("Catch clause regel 113, poging tot _KWS_copyreporttoclipboard")
		_log("clipboard: `n" . clipboard)
		Send {Enter}
		return
	} finally {
		_BlockUserInput(false)
	}

	if (RegExMatch(clipboard, "m)^ingevoerde beelden$" )) { ; undoes the whole operation
		Send, ^z
		Send, ^z
		Send ^{F8}					; Initieer dictee (ctrl F8)
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\samenvattingLabel.png
		if (ErrorLevel = 0) { ;; verwijderd "ingevoerde beelden" onderaan
			MouseGetPos, mouseX, mouseY
			MouseClick, left, FoundX+20, FoundY+30
			Send ^a
			Send {Backspace}
			MouseMove, mouseX, mouseY
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

onveranderdMetVorigVerslag() {
	tempClip := clipboard
	clipboard := ""
 	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch e {
		return
	}
	currentreportunclean := clipboard
	clipboard := ""

	MouseGetPos, mouseX, mouseY
	_BlockUserInput(True)
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\NieuweMededelingHeader.png ;; klikt eerst de mededeling weg als die er is
	if (FoundX) {
		MouseClick, left, FoundX+400, FoundY+15
		sleep 50
	}
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\toonLaatstVerslagKnop.png
	if (FoundX = "") {
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\toonLaatstVerslagKnopSelected.png
		if (FoundX = "") {
			_makeSplashText(title := "Error", text := "Geen vorig verslag aanwezig!", time := -2000)
			MouseMove, mouseX, mouseY
			return
		}
	}
	MouseClick, left, FoundX+10, FoundY+10 
	sleep, 300

	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\laatstVerslagHeader.png
	if (FoundX = "" and FoundX2 = "") {
		_makeSplashText(title := "Error", text := "Verslag popup niet gevonden!", time := -2000)
		MouseMove, mouseX, mouseY
		return
	}
	MouseClick, left, FoundX+100, FoundY+200
	Send, ^a
	Send, {Ctrl down}c{Ctrl up} ;; zou ook reliablity verhogen
	Clipwait, 1
	oldreportunclean := clipboard			; zet variabele gelijk aan clipboard
	clipboard := ""

	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\sluitLaatstVerslagKnop.png
	MouseClick, left, FoundX+5, FoundY+5
	MouseClick, left, FoundX+5, FoundY+650 ;; Selecteerd Textbox van KWS
	sleep, 100 ;; even tijd geven
	if (RegExMatch(oldreportunclean, "m)^ingevoerde beelden$" )) { ; undoes the whole operation
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		_BlockUserInput(false)
		return
	} 

	RegexQuerry := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).*)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"

	RegExMatch(currentreportunclean, RegexQuerry, currentreport)
	RegExMatch(oldreportunclean, RegexQuerry, oldreport)

	compared := ""
	if (oldreportcomparedwith && oldreportcompdate) {
		compared := StrReplace(oldreportcomparedwith, oldreportcompdate, oldreportdate)
		oldreportcontent :=  StrReplace(oldreportcontent, oldreportcompdate, oldreportdate)
	} else {
		compared := "In vergelijking met het voorgaande onderzoek van " . oldreportdate . ":" 
	}
	if (SubStr(currentreportheader, StrLen(currentreportheader)) != "`n") { ; Fix dat soms de "compare" tegen de onderzoeken wordt gezet. eventueel te fixen in de REGEX of gewoon extra newlines hier vanonder en dan de overschot newlines wegdoen
		currentreportheader .= "`n"
	}

	mergedReport := currentreportheader . "`n" . compared . "`n`n- Globaal ongewijzigde positie van het supportmateriaal.`n- Globaal ongewijzigd cardiopulmonaal beeld."
	sleep, 200
	_KWS_PasteToReport(mergedReport, true)
	Send ^{F8}							; Initieer dictee (ctrl F8)
	MouseMove, mouseX, mouseY
	_BlockUserInput(false)
	clipboard := tempClip
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
	ead := _getEAD(true)
	Send, {Ctrl down}{Shift down}v{Ctrl up}{Shift up} ;; KWS knop om te valideren
	_log(ead, "Gevalideerd en gesloten")
	_makeSplashText(title := "Gevalideerd", text := "Gevalideerd.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000)
}

saveAndClose_KWS() {
	global splashExists
	WinActivate, KWS ahk_exe javaw.exe 
	if (splashExists == "" or splashExists == False) {
		_makeSplashText("Save function", "Press save button again to close.", -3000, doublePressMode := True)
		Send, {Ctrl down}s{Ctrl up}
		return
	}
	_destroySplash()
	ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\bewarenGreyedButton.png
	if (ErrorLevel = 2) {
		_makeSplashText(title := "Save function", text := "Something went wrong when checking if it was alreade saved", time := -2000)
	} else if (ErrorLevel = 1) {
		_makeSplashText(title := "Save function", text := "Try again, report was not saved", time := -2000)
		Send, ^s
		return
	} 
	ead := _getEAD(true)

	Send, {Ctrl down}{F4}{Ctrl up} ;; KWS knop om huidig formulier te sluiten
	_log(ead, "Opgeslagen en gesloten")
	_makeSplashText(title := "Opgeslagen", text := "Opgeslagen.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000)
}

closeWithoutSaving() {
	Send, {Ctrl down}{F4}{Ctrl up} ;; KWS knop om huidig formulier te sluiten
	_BlockUserInput(True)
	sleep, 100
	Send {Enter}
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


VolumeCalculator() {
	Global volX
	Global volY
	Global volZ
	Gui, volCalc:+LastFound
	GuiHWND := WinExist()
	Gui, volCalc:Add , Text  ,        , X, Y and Z
	Gui, volCalc:Add , Edit  , vvolX,
	Gui, volCalc:Add , Edit  , vvolY,
	Gui, volCalc:Add , Edit  , vvolZ,
	Gui, volCalc:Add , Button, Default, OK
	Gui, volCalc:Show, , Volume calculator
	WinWaitClose, ahk_id %GuiHWND%  		;--waiting for gui to close
	WinActivate, KWS ahk_exe javaw.exe
	sleep, 50
	return _KWS_PasteToReport(volume, false)               	;--returning value
	;-------
	volCalcButtonOK:
	  GuiControlGet, x, , volX
	  GuiControlGet, y, , volY
	  GuiControlGet, z, , volZ
	  volume := Round(x * y * z * 0.52, 1)
	  ;MsgBox, %x% %y% %z% en volume: %volume%
	  Gui, volCalc:Destroy
	return
	;-------
	volCalcGuiEscape:
	volCalcGuiClose:
	  result := ""
	  Gui, volCalc:Destroy
	return
}

VDTCalculator() {
	Global diameter1, diameter2, date1, date2
	Global VDTDays
	Global VDTresult
	FormatTime, currentdate, A_Now, yyy-MM-dd
	Gui, VDTCalc:+LastFound
	GuiHWND := WinExist()
	Gui, VDTCalc:Add , Text, x60, Date (yyyy-MM-dd)
	Gui, VDTCalc:Add , Text, xp+110 yp+0, Size (mm)
	Gui, VDTCalc:Add , Text, x10 yp+20, Previous
	Gui, VDTCalc:Add , Edit, x60 yp+0 w100 R1 vdate1 gVDTCalcGuiRefresh,
	Gui, VDTCalc:Add , Edit, xp+110 yp+0 w70 R1 vdiameter1 gVDTCalcGuiRefresh number,
	Gui, VDTCalc:Add , Text, x10 yp+25, Current
	Gui, VDTCalc:Add , Edit, x60 yp+0 w100 R1 vdate2 gVDTCalcGuiRefresh, %currentdate%
	Gui, VDTCalc:Add , Edit, xp+110 yp+0 w70 R1 vdiameter2 gVDTCalcGuiRefresh number,
	Gui, VDTCalc:Add , Text, x60 yp+25 w100 vVDTDays hwndVDTDays, Days
	Gui, VDTCalc:Add , Text, xp+110 yp+0 w70 vVDTresult hwndVDTresult, VDT
	Gui, VDTCalc:Add , Button, x60 yp+25 w180 Default, OK
	Gui, VDTCalc:Show, , Volume Doubling Time calculator
	WinWaitClose, ahk_id %GuiHWND%  		;--waiting for gui to close
	WinActivate, KWS ahk_exe javaw.exe
	sleep, 50
	return _KWS_PasteToReport(VDT, false)               	;--returning value
	;-------
	VDTCalcButtonOK:
	  Gui, VDTCalc:Destroy
	return
	;-------
	VDTCalcGuiEscape:
	VDTCalcGuiClose:
	  VDT := ""
	  Gui, VDTCalc:Destroy
	return
VDTCalcGuiRefresh:
	GuiControlGet, diameter1
	GuiControlGet, diameter2
	GuiControlGet, date1
	GuiControlGet, date2
	date1 := RegExReplace(date1, "(\d{4}).?(\d{2}).?(\d{2})", "$1$2$3")
	date2 := RegExReplace(date2, "(\d{4}).?(\d{2}).?(\d{2})", "$1$2$3")
	EnvSub, date2, % date1, Days
	volume1 := Round(diameter1**3 * 0.52, 1)
	volume2 := Round(diameter2**3 * 0.52, 1)
	VDT := Round((ln(2) * date2)/(ln(volume2/volume1)), 0)
	GuiControl, Text, %VDTDays%, % "Interval: " date2 " days"
	GuiControl, Text, %VDTresult%, % "VDT: " VDT " days"
	return
}


RIcalculatorGUI() {
	Global vel1
	Global vel2
	Global RIresult
	Gui, RIcalc:+LastFound
	GuiHWND := WinExist()

	Gui, RIcalc:Add , Text  ,        , Calculate RI from PSV and EDV
	Gui, RIcalc:Add , Edit  , vvel1 gRIcalcRefresh,
	Gui, RIcalc:Add , Edit  , vvel2 gRIcalcRefresh,
	Gui, RIcalc:Add , Text, vRIresult hwndRIresult w100, RI:
	Gui, RIcalc:Add , Button, Default, OK
	Gui, RIcalc:Show, , RI calculator
	WinWaitClose, ahk_id %GuiHWND%  		;--waiting for gui to close
	WinActivate, KWS ahk_exe javaw.exe
	sleep, 50
	return _KWS_PasteToReport(result, false)               	;--returning value
	;-------
	RIcalcButtonOK:
	  result := "PSV: " . Max(v1, v2) . " cm/s; RI " . RI[1]
	  Gui, RIcalc:Destroy
	return
	;-------
	RIcalcGuiEscape:
	RIcalcGuiClose:
	  result := ""
	  Gui, RIcalc:Destroy
	return
RIcalcRefresh:
	  GuiControlGet, v1, , vel1
	  GuiControlGet, v2, , vel2
	  RI := _calcRI(v1, v2)
	  GuiControl, Text, %RIresult%, % "RI: " RI[1] ""
	  return
}

openEAD_KWS(ead := "") {
	if RegExMatch(ead, "[^1-9]?(\d{8})[^1-9]?", matchEAD) {
		clipboard := ""
		sleep 100 ;; 100 werkt dacht ik
		clipboard := matchEAD1
		clipWait, 1
		WinActivate, KWS ahk_exe javaw.exe
		Send, ^+z ; Hotkey voor zoek patient
		Sleep, 400
		Send, {Ctrl down}v{Ctrl up}
		Send, {Enter}
	} else if (ead = "") {
		clipboard := ""
		sleep 100 ;; 150 werkt, te zien of lager werk
		Send, {Ctrl down}c{Ctrl up} ;; zou ook reliablity verhogen
		ClipWait, 1
		openEAD_KWS(clipboard)
	}
}

openLastPtInLog_KWS() {
	global logfile
	Loop, read, %logfile%
		lastline := A_loopreadline
	if (RegExMatch(lastline, "[^1-9]?(\d{8})[^1-9]?", matchEAD)) {
		;; MsgBox, %lastline% en %matchEAD1%
		sleep, 50 ;; needed for some reason ....
		openEAD_KWS(matchEAD1)
	}
}

pedAbdomenTemplate() {
	;; TODO: nieren SD werden niet gedaan bij eentje van 7d
	;; Gemaakt door Johannes Devos, aangepast door CV.
	;; ptdata := _KWS_GetDemographicDataPatient()
	global age, milt, linkerNier, lever, rechterNier
	global SDmilt, SDliNier, SDlever, SDreNier
	result := ""
	birthdate := _getBirthDate(returnMouse := true)
	birthdate := RegExReplace(birthdate, "(\d{2}).(\d{2}).(\d{4})", "$3$2$1")
	FormatTime, now, ,yyyyMMddHHmmss 
	;; age := CalcAge(birthdate . "000000", now)
	age := _CalcAge(birthdate, now)
	year := age[1]
	months := age[2]
	days := age[3]

	Gui, pedAbdGui:+LastFound
	GuiHWND := WinExist()
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
	Gui, pedAbdGui:Add, Edit, x132 y109 w100 h20 vmilt gpedAbdGuiRefresh, 0
	Gui, pedAbdGui:Add, Edit, x132 y129 w100 h20 vlinkerNier gpedAbdGuiRefresh, 0
	Gui, pedAbdGui:Add, Edit, x132 y149 w100 h20 vlever gpedAbdGuiRefresh, 0
	Gui, pedAbdGui:Add, Edit, x132 y169 w100 h20 vrechterNier gpedAbdGuiRefresh, 0
	Gui, pedAbdGui:Add, Text, x242 y109 w30 h20 vSDmilt hwndSDmilt, 0
	Gui, pedAbdGui:Add, Text, x242 y129 w30 h20 vSDliNier hwndSDliNier, 0
	Gui, pedAbdGui:Add, Text, x242 y149 w30 h20 vSDlever hwndSDlever, 0
	Gui, pedAbdGui:Add, Text, x242 y169 w30 h20 vSDreNier hwndSDreNier, 0
	Gui, pedAbdGui:Add, Button, x2 y199 w300 h30 Default, OK
	Gui, pedAbdGui:Show, x759 y391 h236 w305, Echografie Pediatrie Afmetingen
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
pedAbdGuiRefresh:
	GuiControlGet, milt
	GuiControlGet, linkerNier
	GuiControlGet, lever
	GuiControlGet, rechterNier
	SDs := _getStandardDevsPedAbd(age, milt, linkerNier, lever, rechterNier)
	GuiControl, Text, %SDmilt%, % SDs[3]
	GuiControl, Text, %SDliNier%, % SDs[1]
	GuiControl, Text, %SDlever%, % SDs[4]
	GuiControl, Text, %SDreNier%, % SDs[2]
	return
}

pressOKButton() {
	;; not really used, might be used in the future.
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

;; TODO WORK IN PROGRESS
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

KWStoExcel(excelSavePath) {
	global ExcComment
	global ExcKlinInl
	global ExcDiagnVraag
	global ExcTags
	global ExcCategorySelect
	global ExcOnderzoek
	global OpTeVolgen
	global XL

	_BlockUserInput(True)
	_KWS_CopyReportToClipboard()
	RegexQuery := "(?:Leuven|Pellenberg)[\s\S]+(?<datum>\d{2}-\d{2}-\d{4})[\s\S]+(?:KLINISCHE INLICHTINGEN:[\n\r]+)(?<klinlicht>[\s\S]+)[\n\r]{2,}(?:DIAGNOSTISCHE VRAAGSTELLING:[\n\r]+)(?<diagvraag>[\s\S]+)[\n\r]{2,}(?:ONDERZOEKE?N?:[\n\r]+)(?<onderzoek>(?:.+\R{0,1}?)+)\R{2,}(?:[\s\S]+)"
	RegExMatch(clipboard, RegexQuery, report)
	ead := _getEAD()
	_BlockUserInput(false)
	if FileExist(excelSavePath) {
		if not WinExist("ahk_exe EXCEL.EXE") {
			Run, %excelSavePath%
			_makeSplashText(title := "Opening excel", text := "Opening " . excelSavePath, time := 3000)
			WinWait, ahk_exe EXCEL.EXE
		}
	} else {
		WinWait, ahk_exe EXCEL.EXE
		MsgBox, Excel file niet gevonden! Maak een nieuw excel bestand in dezelfde folder als dit script met de naam %excelSavePath%`nEr zal nu geprobeerd een te maken (niet getest)...
		XL := ComObjCreate("Excel.Application")
		XL.Workbooks.Add
		XL.ActiveWorkbook.SaveAs(excelSavePath)
	}
	;; XL := ComObjGet(excelSavePath) ;; looks for excel
	categoryList := "|neuro|thorax|abdomen|spine|MSK|NKO|mammo|urogen|vasc|cardio"
	indexCategory := 1
	category := ""
	Switch
	{
		case RegExMatch(reportonderzoek, "i)hersen|schedel|hypophyse"):
		indexCategory := 2
		case RegExMatch(reportonderzoek, "i)abdomen|lever|pancreas"):
		indexCategory := 4
		case RegExMatch(reportonderzoek, "i)thorax|longen|embol"):
		indexCategory := 3
		case RegExMatch(reportonderzoek, "i)wervel"):
		indexCategory := 5
		case RegExMatch(reportonderzoek, "i)knie|schouder|heup|arm|been|pols"):
		indexCategory := 6
		case RegExMatch(reportonderzoek, "i)oor|hals|rots|schild|speeksel"):
		indexCategory := 7
		case RegExMatch(reportonderzoek, "i)mammo|borst"):
		indexCategory := 8
		case RegExMatch(reportonderzoek, "i)nier|blaas|prostaat|gyna|vrouw|bekken"):
		indexCategory := 9
		case RegExMatch(reportonderzoek, "i)vascul"):
		indexCategory := 10
		case RegExMatch(reportonderzoek, "i)hart|coronair"):
		indexCategory := 11

	}
	Gui, ExcelGUI:+LastFound
	ExcelGuiHWND := WinExist()
	Gui, ExcelGUI:Add, Text, x10 y10 w60 h30, % reportdatum
	Gui, ExcelGUI:Add, Text, xp+70 yp+0 w60 h30, % ead
	Gui, ExcelGUI:Add, DropDownList, xp+70 yp+0 w130 R1 vExcCategorySelect R10 Choose%indexCategory%, % categoryList
	Gui, ExcelGUI:Add, Edit, x10    yp+30            R1 vExcOnderzoek, % reportonderzoek
	Gui, ExcelGUI:Add, Text, x10    yp+40 w130 R2 , Comment
	Gui, ExcelGUI:Add, Edit, xp+140 yp+0  w190 R2 vExcComment, 
	Gui, ExcelGUI:Add, Text, xp-140 yp+40 w130 R2 , Klin. inlichtingen
	Gui, ExcelGUI:Add, Edit, xp+140 yp+0  w190 R2 vExcKlinInl, %reportklinlicht%
	Gui, ExcelGUI:Add, Text, xp-140 yp+40 w130 R2 , Diagn. vraagstelling
	Gui, ExcelGUI:Add, Edit, xp+140 yp+0  w190 R2 vExcDiagnVraag, %reportdiagvraag%
	Gui, ExcelGUI:Add, Text, xp-140 yp+40 w130 R2 , Tags
	Gui, ExcelGUI:Add, Edit, xp+140 yp+0  w190 R2 vExcTags ,
	Gui, ExcelGUI:Add, Text, xp-140 yp+40 w130 R2 , Op te volgen?
	Gui, ExcelGUI:Add, DropDownList, xp+140 yp+0 w190 R5 vOpTeVolgen Choose1, |Over 1 dag|Over 1 week|Over 1 maand|Op te volgen
	Gui, ExcelGUI:Add, Button, x10    yp+40 w130 h20 Default, OK
	Gui, ExcelGUI:Add, Button, xp+140 yp+0 w130 h20 gExcCancelButton, Cancel
	Gui, ExcelGUI:Show, x360 y233, Save to excel script
	WinWaitClose, ahk_id %ExcelGuiHWND%  		;--waiting for gui to close
;; todo: add a check to see if excel has been found and XL object initialized
	return
	ExcelGUIButtonOK:
	XL := ComObjGet(excelSavePath) ;; looks for excel
	Gui, ExcelGUI:Submit
	lastCell := excelFindLastCell(XL).row + 1

	if (InStr(OpTeVolgen, "Over")) { ;; This function finds the future date to follow up on
		FutureDate := A_Now
		Switch
		{
			case InStr(OpTeVolgen, "day"):   EnvAdd, FutureDate, 1, days
			case InStr(OpTeVolgen, "week"):  EnvAdd, FutureDate, 7, days
			case InStr(OpTeVolgen, "month"): EnvAdd, FutureDate, 31, days
		}
		FormatTime, opTeVolgen, %FutureDate%, yyy-MM-dd
	}
	XL.Application.ActiveSheet.range("A"lastCell).value := ead
	XL.Application.ActiveSheet.range("B"lastCell).value := reportdatum
	XL.Application.ActiveSheet.range("C"lastCell).value := ExcCategorySelect
	XL.Application.ActiveSheet.range("D"lastCell).value := ExcComment
	XL.Application.ActiveSheet.range("E"lastCell).value := opTeVolgen
	XL.Application.ActiveSheet.range("F"lastCell).value := RegExReplace(ExcOnderzoek, "\.?[\r\n]", ". ")
	XL.Application.ActiveSheet.range("G"lastCell).value := RegExReplace(ExcKlinInl, "\.?[\r\n]", ". ")
	XL.Application.ActiveSheet.range("H"lastCell).value := RegExReplace(ExcDiagnVraag, "\.?[\r\n]", ". ")
	XL.Application.ActiveSheet.range("I"lastCell).value := ExcTags
	_makeSplashText(title := "Saved to excel", text := ead . " is saved to excel", time := -1500)
	Gui, ExcelGUI:Destroy
		return
		
	ExcCancelButton:
	ExcelGUIGuiEscape:
	ExcelGUIGuiClose:
		Gui, ExcelGUI:Destroy
	return
}

excelFindLastCell(objExcel, sheet := 1) {
	static xlByRows    := 1
	     , xlByColumns := 2
	     , xlPrevious  := 2
	lastRow := objExcel.Sheets(sheet).Cells.Find("*", , , , xlByRows   , xlPrevious).Row
	lastCol := objExcel.Sheets(sheet).Cells.Find("*", , , , xlByColumns, xlPrevious).Column
	return {row: lastRow, column: lastCol}
}


ObjIndexOf(obj, item, case_sensitive:=false) {
	for i, val in obj {
		if (case_sensitive ? (val == item) : (val = item))
			return i
	}
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

IndentLine(direction) {
	temp := clipboard
	clipboard := ""
	Send, {Home}
	Send, +{End}
	Send, ^x
	ClipWait, 1
	line := clipboard
	clipboard := ""
	if (direction = "promote") {
		line := RegExReplace(line, "^\s*[\.\-]?\s*(.+)$", "- $1")
	} else if (direction = "demote") {
		line := RegExReplace(line, "^\s*[\.\-]?\s*(.+)$", "  . $1")
	}
	clipboard := line
	ClipWait, 1
	if (GetKeyState("{Alt}","P")) {
		Send, {Alt Up}
	}
	Send, ^v
}


deleteLine() {
	Send, {End}
	Send, +{Home}
	Send, {Delete}{Delete}
}

yankLine() {
	Send, {Home}
	Send, +{Down}
	Send, ^c
}

insertDatePeriod(daysInFuture := 0) {
	Send, {Home} ;; zorgt dat de functie eender waar in het datumvak gestart kan worden
	FormatTime, CurrentDateTime,, ddMMyyyy
	SendInput %CurrentDateTime%0000
	if (daysInFuture >= 0) {
		FutureDate := A_Now
		EnvAdd, FutureDate, %daysInFuture%, days
		FormatTime, morgen, %FutureDate%, ddMMyyyy
		Send, {tab}
		SendInput %morgen%2359
		;; MsgBox, %morgen%2359
	}
}

insertPastDatePeriod(daysInPast := 1) {
	Send, {Home} ;; zorgt dat de functie eender waar in het datumvak gestart kan worden
	daysInPast := -1 * Abs(daysInPast)
	PastDate := A_Now
	; EnvAdd, PastDate, %daysInPast%, days
	Pastdate += daysInPast, Days
	FormatTime, PastDate, %PastDate%, ddMMyyyy
	SendInput %PastDate%0000
	Send, {tab}
	FormatTime, CurrentDateTime,, ddMMyyyy
	SendInput %CurrentDateTime%0000
}


initKWSWindows() {
	;; Uitvoeringswerklijst
	Send, ^{Space}
	sleep, 150
	SendInput Uitvoeringswerklijst
	sleep, 400
	Send, {Enter}
	winwait, Uitvoeringswerklijst parameters ahk_class SunAwtDialog,,3
	if WinExist("Uitvoeringswerklijst parameters ahk_class SunAwtDialog") {
		Winactivate
		sleep, 450
		Send, {Tab}
		Send, e
		sleep, 350
		Send, {Enter}
	}

	;; receptieeenheid
	Send, ^{Space}
	sleep, 150
	SendInput zet receptie eenheid
	sleep, 450
	Send, {Enter}
	winwait, Kies een receptie eenheid ahk_class SunAwtDialog,,3
	if WinExist("Kies een receptie eenheid ahk_class SunAwtDialog") {
		Winactivate
		sleep, 450
		MouseGetPos, mouseX, mouseY
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, images\EenheidLabel.png
		if (ErrorLevel = 2)
			_makeSplashText("Error", "Something went wrong when looking for the field", time := -2000)
		else if (ErrorLevel = 1)
			return
		else {
			MouseClick, %mousebutton%, FoundX+60, FoundY+5
			MouseMove, mouseX, mouseY
		}
		SendInput 555
		Send, {Enter}
		sleep, 150
		Send, {Tab}
		sleep, 450
		Send, {Enter}
		sleep, 600
	}

	;; receptie lijst
	Send, ^{Space}
	sleep, 150
	SendInput receptie werklijst hos
	sleep, 400
	Send, {Enter}
	winwait, parameters voor receptiewerklijst ahk_class SunAwtDialog,,3
	if WinExist("parameters voor receptiewerklijst ahk_class SunAwtDialog") {
		Winactivate
		sleep, 450
		;; Send, {Down}
		Send, {Tab}
		Send, {Down}{Down}
		Send, {Tab}
		sleep, 150
		insertDatePeriod(0)
		sleep, 250
		Send, {Enter}
		sleep, 2200
	}
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
	Send, {Ctrl down}c{Ctrl up} ;; zou ook reliablity verhogen
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
			sleep 50
		}
		Send, {Ctrl down}v{Ctrl up} ;; zou ook reliablity verhogen
		Sleep 50
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

_calcRI(v1, v2) {
	absolute := Round((Max(v1,v2) - Min(v1,v2)) / Max(v1,v2), 2)
	percentage := Round(((Max(v1,v2) - Min(v1,v2)) / Max(v1,v2)) * 100, 1)
	return [absolute, percentage]
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
	if (RegExMatch(clipboard, "(verslag \*\*[\r\n]{2,})([\s\S]+?)([\r\n]{2,}\*\* Einde)", KWSfiltered)) {
		clipboard := KWSfiltered2
		return
	} else {
		Return
	}
}

_BlockUserInput(block := true) {
	global blockinput
	blockinput := block
	if (blockinput) {
		BlockInput, On
		MouseClick, left, 0, 0,,, U, R ;; zorgt dat als de muis ingeduwd was dat er geen error komt
		;; BlockInput, Mousemove
		BlockInput, Send
	}
	if (not blockinput) {
		BlockInput, Off
		BlockInput, Default
		;; BlockInput, MouseMoveOff
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


findAndReplaceGUI() {
	WinGet, active_id, ID, A ;; gets the window where the script was activated
	clipboard := ""
	Send, {Ctrl down}c{Ctrl up} ;; zou ook reliablity verhogen
	ClipWait, 1
	originalText := clipboard
	global repTextBox
	global findText
	global replaceText
	global RegexToggle
	global IgnoreCaseFlag
	global MultilineFlag
	global SinglelineFlag
	global UngreadyFlag

	Gui, repGUI:+LastFound -DPIscale
	repGuiHWND := WinExist()
	Gui, repGUI:Margin, 10, 10
	Gui, repGUI:Add, Text, w130 , Find
	Gui, repGUI:Add, Edit, w500 R2 gupdateRepGUI vfindText, 
	Gui, repGUI:Add, Text, w130 , Replace
	Gui, repGUI:Add, Edit, w500 R2 gupdateRepGUI vreplaceText, 
	Gui, repGUI:Add, Button, gReplaceButton, Replace!
	Gui, repGUI:Add, Checkbox, yp+5 xp+65 vRegexToggle checked gupdateRepGUI, Enable regex
	Gui, repGUI:Add, Checkbox, yp+0	xp+90 vIgnoreCaseFlag checked gupdateRepGUI, IgnoreCase
	Gui, repGUI:Add, Checkbox, yp+0	xp+90 vMultilineFlag  checked gupdateRepGUI, Multiline-mode
	Gui, repGUI:Add, Checkbox, yp+0	xp+90 vSinglelineFlag gupdateRepGUI, Singleline 
	Gui, repGUI:Add, Checkbox, yp+0	xp+70 vUngreadyFlag   gupdateRepGUI, Ungready
	Gui, repGUI:Add, ActiveX, vrepTextBox x10 w500 h400, htmlfile
	repTextBox.write(_getHTMLReplaceBox(originalText, "", "", False))
	Gui, repGUI:Show,,Find and replace
	WinWaitClose, ahk_id %repGuiHWND%  		;--waiting for gui to close
	Gui, repGui:Destroy
	Return
	;-------
	ReplaceButton:
	GuiControlGet, needle, , findText
	GuiControlGet, replacement, , replaceText
	GuiControlGet, regexTogglevar, , RegexToggle
	GuiControlGet, IgnoreCaseFlag
	GuiControlGet, MultilineFlag
	GuiControlGet, SinglelineFlag
	GuiControlGet, UngreadyFlag
	flag := _findreplaceConstructRegexFlags(IgnoreCaseFlag, MultilineFlag, SinglelineFlag, UngreadyFlag)
	temp := clipboard
	clipboard := ""
	if (regexTogglevar) {
		clipboard := RegExReplace(originalText, flag . needle, replacement)
		if (clipboard = "")
			clipboard := haystack
	} else {
		clipboard := StrReplace(originalText, needle, replacement)
	}
	WinActivate, ahk_id %active_id%
	ClipWait, 1
	sleep, 400
	Send, ^v
	sleep, 100
	clipboard := temp
	Gui, repGui:Destroy
	repGUIEscape:
	repGUIClose:
	  Gui, repGui:Destroy
	return
	updateRepGUI:
		GuiControlGet, needle, , findText
		GuiControlGet, replacement, , replaceText
		GuiControlGet, regexTogglevar, , RegexToggle
		GuiControlGet, IgnoreCaseFlag
		GuiControlGet, MultilineFlag
		GuiControlGet, SinglelineFlag
		GuiControlGet, UngreadyFlag
		flag := _findreplaceConstructRegexFlags(IgnoreCaseFlag, MultilineFlag, SinglelineFlag, UngreadyFlag)
		repTextBox.open()
		repTextBox.write(_getHTMLReplaceBox(originalText, needle, replacement, regexTogglevar, flag))
		repTextBox.close()
	return
}

_getHTMLReplaceBox(haystack, needle, replacement, regextoggle := True, flag := "") {
	; https://www.autohotkey.com/boards/viewtopic.php?t=84074
	html =
	(
		<style>
		body {
			font-family: calibri;
			font-size: 13px;
			white-space: pre-line;
			overflow-wrap: normal;
		}
		.red {
			color: red;
			font-weight: bold;
			text-decoration: line-through;
		}
		.replacement {
			color: blue;
			font-weight: bold;
		}
		</style>
		<body>
		<div>INSERTLOCATION</div>
		</body>
	)
	;; Haalt "de zin van "uit ongevalideerd verslag" weg indien aanwezig
	if (RegExMatch(haystack, "(verslag \*\*[\r\n]{2})([\s\S]+)([\r\n]{2}\*\* Einde)", KWSfiltered)) {
		haystack := KWSfiltered2
	}
	if (regextoggle) {
		replacedText := RegExReplace(haystack, flag . needle, "<span class=""red"">$0</span><span class=""replacement"">" . replacement . "</span>")
		if (replacedText = "") {
			replacedText := "<span class=""replacement"">---------------`nError in Regex code`n---------------</span>`n`n" . haystack
		}
	}
	else {
		replacedText := StrReplace(haystack, needle, "<span class=""red"">" . needle . "</span><span class=""replacement"">" . replacement . "</span>")
	}
	return StrReplace(html, "INSERTLOCATION", "<pre>" . replacedText . "</pre>")
}	

_findreplaceConstructRegexFlags(ignoreCase := 1, multilineMode := 1, singleLineMode := 0, ungreedy := 0) {
	flag := ""
	flag .= ignoreCase ? "i" : ""
	flag .= multilineMode ? "m" : ""
	flag .= singleLineMode ? "s" : ""
	flag .= ungreedy ? "U" : ""
	flag .= ")"
	return flag
}

_makePedReport(age, Milt, LinkerNier, Lever, RechterNier) { ; Gemaakt door Johannes Devos, aangepast
	Result := _getStandardDevsPedAbd(age, Milt, LinkerNier, Lever, RechterNier)
	SetFormat, Float, 0.1
	Verslag := "Normale ligging van de retroperitoneale grote vaten.`nNormale ligging van de organen.`n`nLeverspan: " . Lever/10 . " cm (SD: " . Result[4] . ").`nHomogeen leverparenchym met normale reflectiviteit.`nNormale portahoofdstam en intrahepatische portatakken.`nNormale hepatische venen met normale hepatofugale flow.`nNormale hepatopetale portale flow.`nNormale flow in de a. hepatica.`nGeen gedilateerde intrahepatische of extrahepatische galwegen aangetoond.`nNormale galblaas.`nNormale pancreas. Geen visualisatie van de ductus van Wirsung.`nMilt: " . Milt/10 . " cm (SD: " . Result[3] . ").`nNormale milt.`n`nNormale bijnieren en bijnierloges.`nLinkernier: " . Linkernier/10 . " cm (SD: " . Result[1] . ").`nRechternier: " . Rechternier/10 . " cm (SD: " . Result[2] . ").`nNormale reflectiviteit van het nierparenchym met corticomedullaire differentiatie.`nGeen hydro-ureteronefrose.`nNormale blaasvulling.`nNormale aflijning en dikte van de blaaswand.`n`nNormale ligging van de A. en V. Mesenterica Superior.`nGeen adenopathieen aangetoond.`nNormale darmwanden.`n###Normaal terminale ileum.`n###Normale appendix.`n`nCONCLUSIE:`n###`n`nGECOMMUNICEERDE DRINGENDE BEVINDINGEN:`n"
	return Verslag
}

_getStandardDevsPedAbd(age, Milt, LinkerNier, Lever, RechterNier) {
	;; TODO: toch nog iets mis met de standaardeviaties? Controleren.
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
	return Result
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
