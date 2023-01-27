; ******************************************************
;	Dit bestand bevat de eigenlijke functies, die worden opgeroepen door autohotkey.
;
;	Handleiding: zie readme bestand	of https://github.com/CVanmarcke/KWS-helper
;
;	Auteur: Cedric Vanmarcke
;
;	Voor vragen: cedric.vanmarcke@uzleuven.be
;	Bij fouten, graag het vorige verslag, huidige verslag
;		en resultaat doorsturen naar mijn email.
;
;****************************************************************

; CALLABLE FUNCTIONS
; --------------------------------------

; TODO: pedabdomen nog eens testen of de stddev kloppen.
; TODO: functie om automatisch alle schermen voor spoed te openen.


initKWSHandler() {
	global logfile
	logfile := "logfile.csv"
	CoordMode("Pixel")
	CoordMode("Mouse")
	SetMouseDelay(-1) ;remove delays from mouse actions
	SetDefaultMouseSpeed 0

	;; Verwijderd de Teams cache folder: die neemt soms meer dan een GB aan data in zonder reden.
	try DirDelete("P:\uzlsystem\AppData\Microsoft\Teams\Service Worker\CacheStorage", 1)

	;; indien blockinput true is, zullen de volgende knoppen geblokkeerd worden:
	_makeSplashText(title := "KWS-helper", text := "Started KWS-helper", time := -3000)
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
	inputtext := RegExReplace(inputtext, "m)^([\t\ ])+\*", "$1.")                          ;; Als * geindenteerd is wordt het vervangen met .
	inputtext := RegExReplace(inputtext, "m)[\ \t]+$", "")  ; zorgt dat er geen nutteloze spaties op het einde van de zin komen
	if (RegExMatch(inputtext, "m)^\. ?.+(?:\R|$)") OR RegExMatch(inputtext, "m)^.+# ?.?(?:\R|$)")) {	; only executes if there is ". " or "#" in the script
		inputtext := _sorttext(inputtext) ; zet alle zinnen met een punt vooraan, onder het verslag.
	}
	Sleep(50)
	inputtext := StrReplace(inputtext, "bekend", "gekend", 0)
	inputtext := StrReplace(inputtext, "formaliteit", "voor maligniteit")
	inputtext := StrReplace(inputtext, ": in het kader van de gekende", ": gekende")
	inputtext := StrReplace(inputtext, "in het kader van", "door")
	inputtext := StrReplace(inputtext, "ongewijzigd", "onveranderd", 0)
	inputtext := StrReplace(inputtext, "foraminaal spinaal stenose", "foraminaal- of spinaalstenose")		 ;; frequente speech fout
	inputtext := StrReplace(inputtext, "diffuse restrictie", "diffusie restrictie")			;; frequente speech fout
	inputtext := StrReplace(inputtext, "normale doorgankelijkheid van de", "normaal doorgankelijke")
	inputtext := StrReplace(inputtext, "suscebiliteits", "susceptibiliteits")
	inputtext := StrReplace(inputtext, "pig katheter", "PIC katheter")
	inputtext := StrReplace(inputtext, "flair ", "FLAIR ", 0)
	inputtext := StrReplace(inputtext, "fascikels graad", "Fazekas graad", , &CaseSensitive := false)
	inputtext := StrReplace(inputtext, "tbc", "TBC")
	inputtext := StrReplace(inputtext, " EKG ", " ECG ", 0)
	inputtext := StrReplace(inputtext, " ecg ", " ECG ", 0)
	inputtext := StrReplace(inputtext, " hili", " hila")
	inputtext := StrReplace(inputtext, "longtrauma", "longtrama")
	inputtext := StrReplace(inputtext, "aortaal", "aortisch")
	inputtext := StrReplace(inputtext, "op niveau van", "aan")
	inputtext := RegExReplace(inputtext, "[KkCc]a[mn] configuratie", "`"cam`" configuratie")
	;; inputtext := StrReplace(inputtext, "bewaarde", "intacte")
	inputtext := StrReplace(inputtext, "partiële beeld", "partiëel in beeld")
	inputtext := StrReplace(inputtext, "plaatsen schroef", "plaat en schroef")
	inputtext := StrReplace(inputtext, "rx ", "RX ", 0)
	inputtext := RegExReplace(inputtext, "(?<=[a-z\d]),(?=[a-z])", ", ")  ; zet een extra spatie achter komma als nog niet aanwezig
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])VIII(?=[\/\-\s\.\,])", "8")
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])VII(?=[\/\-\s\.\,])", "7")
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])VI(?=[\/\-\s\.\,])", "6")
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])IV([ab])(?=[\/\-\s\.\,])", "4$1")
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])V(?=[\/\-\s\.\,])", "5")
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])III(?=[\/\-\s\.\,])", "3")
	inputtext := RegExReplace(inputtext, "(?<=[\/\-\s])II(?=[\/\-\s\.\,])", "2")
	inputtext := RegExReplace(inputtext, "segment I(?=[ ,\.])", "segment 1")
	inputtext := RegExReplace(inputtext, "(?<=\d )n?o?r?maal(?= \d)", "x")	; Corrigeert een veel gemaakte fout van de speech
	inputtext := RegExReplace(inputtext, "i)(peri|infra|supra) (?=[\w\-])", "$1")	; peri centimetrisch -> pericentimetrisch

       ;; mammo:
	inputtext := RegExReplace(inputtext, "(eefseltype(?:ring)?)\:? ([a-d])", "$1 $U2")	; zet het weefseltype in hoofdletters

	inputtext := RegExReplace(inputtext, "[\r\n]GECOMMUNICEERDE DRINGENDE BEVINDINGEN:[\n\r]?$", "")	; verwijderd die zin
	inputtext := RegExReplace(inputtext, "im)^Vergelijking met ", "Vergeleken met ")
	; inputtext := RegExReplace(inputtext, "i)gekende?", "\#\#\#")
	inputtext := RegExReplace(inputtext, "im)[\ \t]*supervis.*$", "") ; verwijderd supervisie.

	;;; inputtext := RegExReplace(inputtext, "([\n\r\.]) +(?=[\n\r])", "$1")  ; zorgt dat er geen spatie achter . of op nieuwe lijn komt
	inputtext := RegExReplace(inputtext, "([A-Z])([A-Z][a-z]{3,})", "$U1$L2")				; corrigeert WOord naar Woord
	inputtext := RegExReplace(inputtext, "(?<=^|[\n\r])\*\s?(.+?):? ?(?=\R)", "* $U1:")			; adds : at end of string with * and makes uppercase. Not done with m) because of strange bug where it would only capture the first
	inputtext := RegExReplace(inputtext, "m)([\w\d\)\%\°])\ ?(?=\R|$)", "$1.")				; adds . to end of string, word, digit or )
	inputtext := RegExReplace(inputtext, "m)(?<=(?<! a)\. |\? |^- |^)(\w)", "$U1")				; converts to uppercase after ., newline or newline - (exception for a. hepatica etc)
	inputtext := RegExReplace(inputtext, "([a-z])([\:\.])([a-zA-Z])", "$1$2 $3")				; makes sure there is a space after a colon or point (if not number)...
	inputtext := RegExReplace(inputtext, "(?<=[\:\;])\ ?([A-Z][^A-Z])", " $L1")				; converts after : or ; to lowercase (escept if 2x capital letter) for eg. DD, FLAIR, ...
;;	inputtext := RegExReplace(inputtext, "(?<=[\-\/\ ])[D](?=[1-9](?:[0-2]|[\ \:\ ]))", "T") ; Corrects -T10 of /D10 naar -Th10
	;;inputtext := RegExReplace(inputtext, "(?<=[\-\/])[DT](?=[1-9](?:[0-2]|[\ \:]))", "Th") ; Corrects -T10 of /D10 naar -Th10
	;;inputtext := RegExReplace(inputtext, "(?<=\ )[DT](?=[1-9][0-2]?[\-\/])", "Th") ; Corrects T10- naar Th10
	;;inputtext := RegExReplace(inputtext, "([CThD]\d{1,2}[\/-])[TD](?=\d{1,2})", "$1Th")			; corrects T1/X to Th1 TODO: werkt niet T11-L3
	;;inputtext := RegExReplace(inputtext, "[TD](?=\d{1,2}[\/-][ThDL]{1,2}\d{1,2})", "Th")			; corrects X/T1 to Th1
	inputtext := RegExReplace(inputtext, "((?:C|Th|L|S)\d{1,2})\/((?:C|Th|L|S)\d{1,2})", "$1-$2")	; corrects L1/L2 to L1-L2
	inputtext := RegExReplace(inputtext, "(\d{1,2})\/(\d{1,2})\/(\d{2,4})", "$1-$2-$3")			; corrects d/m/y tot d-m-y
	inputtext := RegExReplace(inputtext, "\R{3,}", "`n`n")											; replaces triple+ newline with double
;; TODO: checken of die [A-Z] ok is, want is toch met case insensitive gedaan...
	inputtext := RegExReplace(inputtext, "im)^\-?(?<=\-)?(?=\w|\(|\`")(?![A-Z -]{5,}\:[\r\n]|CONCLUSIE|Vergeleken|Mede in|In (?:vergel|vgl)|NB|Nota|Storende|Suboptim|Opname in|Reserve|Naar [lr]|[PBT]IRADS|\d[\/\)\.])", "- ")	; adds - to all words and (, excluding BESLUIT, vergeleken...
	inputtext := RegExReplace(inputtext, "(CONCLUSIE:[\n\r\R])\-\ (.+)(?:[\n\r\R]|$)(?!-)", "$1$2")	; Als maar 1 lijn conclusie, zal het het streepje weglaten. WERKT NOG NIET
	;;inputtext := RegExReplace(inputtext, "(\d )a( \d)", "$1Ãƒ$2")							; maakt  als a tussen 2 getallen.
	inputtext := RegExReplace(inputtext, "([\-\.]) {2,}(?=[\R\n\r\w])", "$1 ")  ; zorgt dat er niet meer dan 1 spatie na een streepje komt
	inputtext := RegExReplace(inputtext, "\ +, \ +", ", ")  ; verwijdert te veel spaties rond een komma
	inputtext := RegExReplace(inputtext, "\x{2013}", "-")  ; veranderd het unicode streepje (â€“ aka \x{2013}) naar een ASCII streepje. Nog niet getest.
	return inputtext
}

cleanReport_KWS() {
	RegexQuery := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:ONDERZOEKE?N?:\R{1,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<content>[\s\S]+?)(?:\R*$|[\n\r]{2,}\*\* Eind)"
	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch Error as e {
		MsgBox(type(e) " in " e.What ", which was called at line " e.Line)
		return
	}
	RegExMatch(A_Clipboard, RegexQuery, &report)
	if (isSet(report))
		_KWS_PasteToReport(report["header"] . "`n" . cleanreport(report["content"]))
}


mergeReport(currentreportunclean, oldreportunclean) {
	RegexQuery := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).*)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"

	foundcurrent  := RegExMatch(currentreportunclean, RegexQuery, &currentreport)
	foundprevious := RegExMatch(oldreportunclean, RegexQuery, &oldreport)
	;; Checkt of de regex van het vorige verslag gelukt is, en zo niet verwijderd het hele gedoe.
	if (not foundprevious) {
		_makeSplashText(title := "ERROR", text := "Probleem met de layout van het vorige verslag: is het een extern onderzoek?", time := -2000)
		return ""
	}
	if (not foundcurrent) {
		_makeSplashText(title := "ERROR", text := "Probleem met het huidige verslag te kopieren of layout herkennen: probeer opnieuw.", time := -2000)
		return ""
	}
	oldreportcontent := oldreport["content"]
	currentreportheader := currentreport["header"]
	;; Veranderd de datum in een reeds bestaande vergelijking, of voegt de vergelijktekst toe indien die nog niet aanwezig was.
	compared := ""
	if (oldreport["comparedwith"] != "" and oldreport["compdate"] = "") {
		compared := StrReplace(oldreport["comparedwith"], oldreport["compdate"], oldreport["date"])
		oldreportcontent :=  StrReplace(oldreportcontent, oldreport["compdate"], oldreport["date"])
	} else
		compared := "In vergelijking met het voorgaande onderzoek van " . oldreport["date"] . ":"

	; zoekt naar het besluit, en als het het vindt voegt het "vergelijking met" toe of veranderdt het de datum van "in vergelijking met"
	conclusielocatie := RegExMatch(oldreportcontent, "(BESLUIT|CONCLUSIE).*[\n\r]", &conclusieText)
	if (conclusielocatie) {
		if (RegExMatch(oldreportcontent, "(?:ergel(?:ij|e)k|opzichte|vgl |tov ).{5,55}?(?<date>\d+[-\/.]\d+[-\/.]\d+)", &conclcompare, conclusielocatie - 5)) {
			oldreportcontent := StrReplace(oldreportcontent, conclcompare["date"], oldreport["date"])
		} else {
			oldreportcontent := RegExReplace(oldreportcontent, "((?:BESLUIT|CONCLUSIE).*)\R", "$1`nIn vergelijking met het voorgaande onderzoek van " . oldreport["date"] . ":`n", , 1, conclusielocatie-5)
		}
	}
	if (SubStr(currentreportheader, -1) != "`n") ; Fix dat soms de "compare" tegen de onderzoeken wordt gezet. eventueel te fixen in de REGEX of gewoon extra newlines hier vanonder en dan de overschot newlines wegdoen
		MsgBox("Testnotification: adding newline")
		currentreportheader := currentreportheader . "`n"
	if (InStr(currentreport["type"], "RX thorax"))
		oldreportcontent := cleanreport(oldreportcontent) ;; automatisch maakt het verslag proper als het een RX thorax is.
	oldreportcontent := RegExReplace(oldreportcontent, "im)[\ \t]*supervis.*$", "") ; verwijder supervisie.
	; report := currentreportheader . "`n" . compared . "`n`n" . oldreportcontent
	return currentreportheader . "" . compared . "`n`n" . oldreportcontent
}


copyLastReport_KWS() {
	tempClip := A_Clipboard
	A_Clipboard := ""
	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch Error as e {
		return
	}
	currentreportunclean := A_Clipboard
	A_Clipboard := ""

	MouseGetPos(&mouseX, &mouseY)
	_BlockUserInput(True)
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\NieuweMededelingHeader.png") ;; klikt eerst de mededeling weg als die er is
	if (FoundX) {
		MouseClick("left", FoundX+400, FoundY+15)
		Sleep(50)
	}
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\toonLaatstVerslagKnop.png")
	if (FoundX = "") {
		ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\toonLaatstVerslagKnopSelected.png")
		if (FoundX = "") {
			_makeSplashText(title := "Error", text := "Geen vorig verslag aanwezig!", time := -2000)
			MouseMove(mouseX, mouseY)
			return
		}
	}
	MouseClick("left", FoundX+10, FoundY+10)
	Sleep(300)

	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\laatstVerslagHeader.png")
	if (FoundX = "") {
		_makeSplashText(title := "Error", text := "Verslag popup niet gevonden!", time := -2000)
		MouseMove(mouseX, mouseY)
		return
	}
	MouseClick("left", FoundX+100, FoundY+200)
	Sleep(100)
	Send("^a")
	Send("{Ctrl down}c{Ctrl up}") ;; zou ook reliablity verhogen
	Errorlevel := !ClipWait(1)
	oldreportunclean := A_Clipboard			; zet variabele gelijk aan clipboard

	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\sluitLaatstVerslagKnop.png")
	MouseClick("left", FoundX+5, FoundY+5)
	Sleep(100)
	MouseClick("left", FoundX+5, FoundY+650) ;; Selecteerd Textbox van KWS
	Sleep(100) ;; even tijd geven
	A_Clipboard := ""
	if (RegExMatch(oldreportunclean, "m)^ingevoerde beelden$")) { ; undoes the whole operation
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		_BlockUserInput(false)
		return
	}
	mergedReport := mergeReport(currentreportunclean, oldreportunclean)
	If (not mergedReport = "") ;; Checkt dat het effectief gelukt is om samen te voegen
		_KWS_PasteToReport(mergedReport, true)
	Send("^{F8}")							; Initieer dictee (ctrl F8)
	MouseMove(mouseX, mouseY)
	_BlockUserInput(false)
	A_Clipboard := tempClip
	return								; Klaar
}

selectLastReport_KWS() {
	RegexQuery := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).*)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"
	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch Error as e {
		return
	}
	_BlockUserInput(True)
	Send("{Down}") ;; deselecteert de inhoud
	currentreportunclean := A_Clipboard
	A_Clipboard := ""
	_KWS_SelectReportBox("right")
	_BlockUserInput(True)
	Send("{Down}")					; klik pijltje naar beneden
	Send("{Down}")
	Send("{Enter}")					; selecteerd "neem laatste verslag over"
	if (A_UserName = "cvmarc2") {
		;; Mag enkel uitgevoerd worden als de technische dienst de toegangsactie "laatsteGelijkaardigVerslagDetails" heeft opengezet! Indien niet, is dit niet nodig.
		WinWait("Neem een gelijkaardig verslag over ahk_class SunAwtDialog ahk_exe javaw.exe", , 3)
		WinActivate()
		_BlockUserInput(false)
		WinWaitClose("Neem een gelijkaardig verslag over ahk_class SunAwtDialog ahk_exe javaw.exe")
		Sleep(200)
		WinActivate("KWS ahk_exe javaw.exe")
	}
	Sleep(400)					; geeft tijd om vorig verslag te laden, kan evt verhoogd of verlaagd worden (400 werkt sowieso)
	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch Error as e {
		_makeSplashText(title := "ERROR", text := "Probleem met het voorgaande verslag te kopieren!", time := -3000)
		_log("Catch clause regel 113, poging tot _KWS_copyreporttoclipboard")
		_log("A_Clipboard: `n" . A_Clipboard)
		Send("{Enter}")
		return
	} finally {
		_BlockUserInput(false)
	}

	if (RegExMatch(A_Clipboard, "m)^ingevoerde beelden$")) { ; undoes the whole operation
		Send("^z")
		Send("^z")
		Send("^{F8}")					; Initieer dictee (ctrl F8)
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		if (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\samenvattingLabel.png")) { ;; verwijderd "ingevoerde beelden" onderaan
			MouseGetPos(&mouseX, &mouseY)
			MouseClick("left", FoundX+20, FoundY+30)
			Send("^a")
			Send("{Backspace}")
			MouseMove(mouseX, mouseY)
		}
		return
	}

	oldreportunclean := A_Clipboard			; zet variabele gelijk aan clipboard
	A_Clipboard := ""				; maakt het clipboard leeg
	RegExMatch(oldreportunclean, RegexQuery, &oldreport)
	RegExMatch(currentreportunclean, RegexQuery, &currentreport)
	compared := ""
	;; Checkt of de regex van het vorige verslag gelukt is, en zo niet verwijderd het hele gedoe.
;; TODO NOT CHECKED FOR ERRORS
	; if (oldreport["content"] = "") {
	;	_log("Catch clause regel 143, error bij regexmatch")
	;	_log("old report: `n" . oldreportunclean)
	;	_makeSplashText(title := "ERROR", text := "Probleem met REGEX van het A_Clipboard van het vorige verslag: sluit de patient en probeer opnieuw.", time := -2000)
	;	Send("^z")
	;	Send("^z")
	;	Send("^{F8}")					; Initieer dictee (ctrl F8)
	; }
	; ;; Veranderd de datum in een reeds bestaande vergelijking, of voegt de vergelijktekst toe indien die nog niet aanwezig was.
	; if (oldreport["comparedwith"] && oldreport["compdate"]) {
	;	compared := StrReplace(oldreport["comparedwith"], oldreport["compdate"], oldreport["date"])
	;	oldreportcontent :=  StrReplace(oldreport["content"], oldreport["compdate"], oldreport["date"])
	; } else {
	;	compared := "In vergelijking met het voorgaande onderzoek van " . oldreport["date"] . ":"
	; }

	; ; zoekt naar het besluit, en als het het vindt voegt het "vergelijking met" toe of veranderdt het de datum van "in vergelijking met"
	; conclusielocatie := RegExMatch(oldreport["content"], "(BESLUIT|CONCLUSIE).*[\n\r]", &conclusieText)
	; if (conclusielocatie) {
	;	RegExMatch(oldreport["content"], "(?:ergel(?:ij|e)k|opzichte|vgl |tov ).{5,55}?(?<date>\d+[-\/.]\d+[-\/.]\d+)", &conclcompare, (conclusielocatie - 5)<1 ? (conclusielocatie - 5)-1 : (conclusielocatie - 5))
	;	if (conclcompare["date"]) {
	;		oldreportcontent := StrReplace(oldreport["content"], conclcompare["date"], oldreport["date"])
	;	} else {
	;		oldreportcontent := RegExReplace(oldreport["content"], "((?:BESLUIT|CONCLUSIE).*)\R", "$1`nIn vergelijking met het voorgaande onderzoek van " . oldreport["date"] . ":`n", , 1, (conclusielocatie-5)<1 ? (conclusielocatie-5)-1 : (conclusielocatie-5))
	;	}
	; }
	; if (SubStr(currentreport["header"], (StrLen(currentreport["header"]))<1 ? (StrLen(currentreport["header"]))-1 : (StrLen(currentreport["header"]))) != "`n") { ; Fix dat soms de "compare" tegen de onderzoeken wordt gezet. eventueel te fixen in de REGEX of gewoon extra newlines hier vanonder en dan de overschot newlines wegdoen
	;	currentreport["header"] .= "`n"
	; }
	; if (InStr(currentreport["type"], "RX thorax")) {
	;	oldreportcontent := cleanreport(oldreport["content"]) ;; automatisch maakt het verslag proper als het een RX thorax is.
	; }
	; oldreportcontent := RegExReplace(oldreport["content"], "im)[\ \t]*supervis.*$", "") ; verwijder supervisie.
	; _KWS_PasteToReport(currentreport["header"] . "`n" . compared . "`n`n" . oldreport["content"])
	; Send("^{F8}")							; Initieer dictee (ctrl F8)
	; MouseMove(mouseX, mouseY)
	return								; Klaar
}

onveranderdMetVorigVerslag() {
	tempClip := A_Clipboard
	A_Clipboard := ""
	try _KWS_CopyReportToClipboard(selectReportBox := True)
	catch Error as e {
		return
	}
	currentreportunclean := A_Clipboard
	A_Clipboard := ""

	MouseGetPos(&mouseX, &mouseY)
	_BlockUserInput(True)
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\NieuweMededelingHeader.png") ;; klikt eerst de mededeling weg als die er is
	if (FoundX) {
		MouseClick("left", FoundX+400, FoundY+15)
		Sleep(50)
	}
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\toonLaatstVerslagKnop.png")
	if (FoundX = "") {
		ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\toonLaatstVerslagKnopSelected.png")
		if (FoundX = "") {
			_makeSplashText(title := "Error", text := "Geen vorig verslag aanwezig!", time := -2000)
			MouseMove(mouseX, mouseY)
			return
		}
	}
	MouseClick("left", FoundX+10, FoundY+10)
	Sleep(300)

	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\laatstVerslagHeader.png")
	if (FoundX = "") {
		_makeSplashText(title := "Error", text := "Verslag popup niet gevonden!", time := -2000)
		MouseMove(mouseX, mouseY)
		return
	}
	MouseClick("left", FoundX+100, FoundY+200)
	Sleep(100)
	Send("^a")
	Send("{Ctrl down}c{Ctrl up}") ;; zou ook reliablity verhogen
	Errorlevel := !ClipWait(1)
	oldreportunclean := A_Clipboard			; zet variabele gelijk aan clipboard

	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\sluitLaatstVerslagKnop.png")
	MouseClick("left", FoundX+5, FoundY+5)
	Sleep(100)
	MouseClick("left", FoundX+5, FoundY+650) ;; Selecteerd Textbox van KWS
	Sleep(100) ;; even tijd geven
	A_Clipboard := ""
	if (RegExMatch(oldreportunclean, "m)^ingevoerde beelden$")) { ; undoes the whole operation
		_makeSplashText(title := "ERROR", text := "Fout: het voorgaande verslag zijn ingevoerde beelden!", time := -2000)
		_BlockUserInput(false)
		return
	}


	RegexQuery := "(?<header>(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?:Onderzoeksdatum: )(?<date>\d{2}-\d{2}-\d{4})[\s\S]+(?:ONDERZOEKE?N?:\R{0,2})(?<type>(?:.+\R?)+)(?:\R*TOEGEDIENDE MEDI[CK]ATIE.+:\R(?:.+\R?)+)?)\R*(?<comparedwith>.{0,22}(?:(?:ergel(?:ij|e)k)|opzichte|vgl\.? |tov\.? |Ivm).{3,55}?(?:(?<compdate>\d+[-\/.]\d+[-\/.]\d+)|gisteren|vandaag).*)?\R+(?<content>[\s\S]+?)(?:\R*$|\*\* Eind)"

	foundcurrent  := RegExMatch(currentreportunclean, RegexQuery, &currentreport)
	foundprevious := RegExMatch(oldreportunclean, RegexQuery, &oldreport)
	;; Checkt of de regex van het vorige verslag gelukt is, en zo niet verwijderd het hele gedoe.
	if (not foundprevious) {
		_makeSplashText(title := "ERROR", text := "Probleem met de layout van het vorige verslag: is het een extern onderzoek?", time := -2000)
		return ""
	}
	if (not foundcurrent) {
		_makeSplashText(title := "ERROR", text := "Probleem met het huidige verslag te kopieren of layout herkennen: probeer opnieuw.", time := -2000)
		return ""
	}

	compared := "In vergelijking met het voorgaande onderzoek van " . oldreport["date"] . ":"

	mergedReport := currentreport["header"] . "`n" . compared . "`n`n- Globaal ongewijzigde positie van het supportmateriaal.`n- Globaal ongewijzigd cardiopulmonaal beeld."

	If (not mergedReport = "") ;; Checkt dat het effectief gelukt is om samen te voegen
		_KWS_PasteToReport(mergedReport, true)
	Send("^{F8}")							; Initieer dictee (ctrl F8)
	MouseMove(mouseX, mouseY)
	_BlockUserInput(false)
	A_Clipboard := tempClip
}

validateAndClose_KWS() {
	global splashExists
	if (not isSet(splashExists) or splashExists == False) {
		_makeSplashText("Validate function", "Press the button again to validate and close.", time := -3000, doublePressMode := True)
		Send("^s")
		return
	}
	_destroySplash()
	WinActivate("KWS ahk_exe javaw.exe")
	ead := _getEAD(true)
	Send("{Ctrl down}{Shift down}v{Ctrl up}{Shift up}") ;; KWS knop om te valideren
	_log(ead, "Gevalideerd en gesloten")
	_makeSplashText(title := "Gevalideerd", text := "Gevalideerd.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000)
}

saveAndClose_KWS() {
	global splashExists
	WinActivate("KWS ahk_exe javaw.exe")
	if (not isSet(splashExists) or splashExists == False) {
		_makeSplashText("Save function", "Press save button again to close.", -3000, doublePressMode := True)
		Send("{Ctrl down}s{Ctrl up}")
		return
	}
	_destroySplash()
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\bewarenGreyedButton.png")
	if (ErrorLevel) {
		_makeSplashText(title := "Save function", text := "Something went wrong when checking if it was alreade saved, or report was not saved", time := -2000)
		Send("^s")
		return
	}
	ead := _getEAD(true)

	Send("{Ctrl down}{F4}{Ctrl up}") ;; KWS knop om huidig formulier te sluiten
	_log(ead, "Opgeslagen en gesloten")
	_makeSplashText(title := "Opgeslagen", text := "Opgeslagen.`n`nEAD (" . ead . ") opgeslagen in de logfile.", time := -1000)
}

closeWithoutSaving() {
	Send("{Ctrl down}{F4}{Ctrl up}") ;; KWS knop om huidig formulier te sluiten
	_BlockUserInput(True)
	Sleep(100)
	Send("{Enter}")
	_BlockUserInput(false)
}

heightLossGUI() {
	heightLoss := Gui()
	heightLoss.OnEvent("Close", heightLossGuiHandler.bind("Close", heightLoss))
	heightLoss.OnEvent("Escape", heightLossGuiHandler.bind("Close", heightLoss))
	heightLoss.Opt("+LastFound")
	GuiHWND := WinExist()
	heightLoss.Add("Text", , "Height collapsed and normal vertebra")
	ogcEditHeight1 := heightLoss.Add("Edit", "vHeight1")
	ogcEditHeight2 := heightLoss.Add("Edit", "vHeight2")
	ogcButtonOK := heightLoss.Add("Button", "Default", "OK")
	ogcButtonOK.OnEvent("Click", heightLossGuiHandler.bind("OK", heightLoss))
	heightLoss.Title := "Vertebral height loss calculator"
	heightLoss.Show()
}

heightLossGuiHandler(A_GuiEvent, GuiCtrlObj, info, *) {
	if (A_GuiEvent = "Close") {
		GuiCtrlObj.Destroy()
		Return
	}
	if (A_GuiEvent = "OK") {
		h1 := GuiCtrlObj["Height1"].Text
		h2 := GuiCtrlObj["Height2"].Text
		hl := _calcHeightLoss(h1, h2)
		result := "hoogteverlies van " . hl[1] . " mm of " . hl[2] . "%"
		WinActivate("KWS ahk_exe javaw.exe")
		Sleep(50)
		_KWS_PasteToReport(result, false)
		GuiCtrlObj.Destroy()
	}
}


VolumeCalculator() {
	volCalc := Gui()
	volCalc.OnEvent("Close", volCalcGuiHandler.Bind("Close", volCalc))
	volCalc.OnEvent("Escape", volCalcGuiHandler.Bind("Close", volCalc))
	volCalc.Opt("+LastFound")
	GuiHWND := WinExist()
	volCalc.Add("Text", , "X, Y and Z")
	ogcEditvolX := volCalc.Add("Edit", "vvolX")
	ogcEditvolY := volCalc.Add("Edit", "vvolY")
	ogcEditvolZ := volCalc.Add("Edit", "vvolZ")
	ogcButtonOK := volCalc.Add("Button", "Default", "OK")
	ogcButtonOK.OnEvent("Click", volCalcGuiHandler.Bind("OK", volCalc))
	volCalc.Title := "Volume calculator"
	volCalc.Show()
}
volCalcGuiHandler(A_GuiEvent, GuiCtrlObj, info, *) {
	if (A_GuiEvent = "Close") {
		GuiCtrlObj.Destroy()
		return
	}
	if (A_GuiEvent = "OK") {
		x := toInteger(GuiCtrlObj["volX"].Text)
		y := toInteger(GuiCtrlObj["volY"].Text)
		z := toInteger(GuiCtrlObj["volZ"].Text)
		volume := Round(x * y * z * 0.52, 1)
		WinActivate("KWS ahk_exe javaw.exe")
		Sleep(50)
		_KWS_PasteToReport(volume, false)
		GuiCtrlObj.Destroy()
	}
}

VDTCalculator() {
	currentdate := FormatTime("A_Now", "yyy-MM-dd")
	VDTCalc := Gui()
	VDTCalc.OnEvent("Close", VDTCalcGuiHandler.bind("Close", VDTCalc))
	VDTCalc.OnEvent("Escape", VDTCalcGuiHandler.bind("Close", VDTCalc))
	VDTCalc.Opt("+LastFound")
	GuiHWND := WinExist()
	VDTCalc.Add("Text", "x60", "Date (yyyy-MM-dd)")
	VDTCalc.Add("Text", "xp+110 yp+0", "Size (mm)")
	VDTCalc.Add("Text", "x10 yp+20", "Previous")
	ogcEditdate1 := VDTCalc.Add("Edit", "x60 yp+0 w100 R1 vdate1")
	ogcEditdate1.OnEvent("Change", VDTCalcGuiHandler.Bind("Change", VDTCalc))
	ogcEditdiameter1 := VDTCalc.Add("Edit", "xp+110 yp+0 w70 R1 vdiameter1  number")
	ogcEditdiameter1.OnEvent("Change", VDTCalcGuiHandler.Bind("Change", VDTCalc))
	VDTCalc.Add("Text", "x10 yp+25", "Current")
	ogcEditdate2 := VDTCalc.Add("Edit", "x60 yp+0 w100 R1 vdate2", currentdate)
	ogcEditdate2.OnEvent("Change", VDTCalcGuiHandler.Bind("Change", VDTCalc))
	ogcEditdiameter2 := VDTCalc.Add("Edit", "xp+110 yp+0 w70 R1 vdiameter2  number")
	ogcEditdiameter2.OnEvent("Change", VDTCalcGuiHandler.Bind("Change", VDTCalc))
	ogcTextVDTDays := VDTCalc.Add("Text", "x60 yp+25 w100 vVDTDays", "Days")
	ogcTextVDTresult := VDTCalc.Add("Text", "xp+110 yp+0 w70 vVDTresult", "VDT")
	ogcButtonOK := VDTCalc.Add("Button", "x60 yp+25 w180 Default", "OK")
	ogcButtonOK.OnEvent("Click", VDTCalcGuiHandler.Bind("OK", VDTCalc))
	VDTCalc.Title := "Volume Doubling Time calculator"
	VDTCalc.Show()
}

VDTCalcGuiHandler(A_GuiEvent, GuiCtrlObj, Info := "", *) { ; V1toV2: Added bracket
	diameter1 := toInteger(GuiCtrlObj["diameter1"].value)
	diameter2 := toInteger(GuiCtrlObj["diameter2"].value)
	; date1 := toInteger(GuiCtrlObj["date1"].value)
	; date2 := toInteger(GuiCtrlObj["date2"].value)
	volume1 := Round(diameter1**3 * 0.52, 1)
	volume2 := Round(diameter2**3 * 0.52, 1)
	if (RegexMatch(GuiCtrlObj["date1"].value, "(\d{4}).?(\d{2}).?(\d{2})", &date1) and
	RegexMatch(GuiCtrlObj["date2"].value, "(\d{4}).?(\d{2}).?(\d{2})", &date2) ) {
		; if (date1.count = 3 and date2.count = 3) {
		date1 := date1.1 . date1.2 . date1.3
		date2 := date2.1 . date2.2 . date2.3
		daysDifference := DateDiff(date2, date1, "Days")
		GuiCtrlObj["VDTDays"].value := "Interval: " daysDifference " days"
		if (volume1 and volume2 != volume1) {
			VDT := Round((ln(2) * daysDifference)/(ln(volume2/volume1)), 0)
			GuiCtrlObj["VDTresult"].value := "VDT: " VDT " days"
		} else
		GuiCtrlObj["VDTresult"].value := "VDT: "
	}
	If (A_GuiEvent = "OK") {
		WinActivate("KWS ahk_exe javaw.exe")
		Sleep(50)
		_KWS_PasteToReport(VDT, false)
		GuiCtrlObj.Destroy()
	}
	If (A_GuiEvent = "close")
		GuiCtrlObj.Destroy()
}

toInteger(int) {
	return Integer(IsNumber(int) ? int : 0)
}

RIcalculatorGUI() {
	RIcalc := Gui()
	RIcalc.Opt("+LastFound")
	RIcalc.OnEvent("Close", RIcalcGuiHandler.bind("Close", RIcalc))
	RIcalc.OnEvent("Escape", RIcalcGuiHandler.bind("Close", RIcalc))
	GuiHWND := WinExist()

	RIcalc.Add("Text", , "Calculate RI from PSV and EDV")
	ogcEditvel1 := RIcalc.Add("Edit", "vvel1")
	ogcEditvel1.OnEvent("Change", RIcalcGuiHandler.Bind("Change", RIcalc))
	ogcEditvel2 := RIcalc.Add("Edit", "vvel2")
	ogcEditvel2.OnEvent("Change", RIcalcGuiHandler.Bind("Change", RIcalc))
	ogcTextRIresult := RIcalc.Add("Text", "vRIresult w150", "RI:")
	ogcButtonOK := RIcalc.Add("Button", "Default", "OK")
	ogcButtonOK.OnEvent("Click", RIcalcGuiHandler.Bind("Normal", RIcalc))
	RIcalc.Title := "RI calculator"
	RIcalc.Show()
}

RIcalcGuiHandler(A_GuiEvent, GuiCtrlObj, info, *) {
	if (A_GuiEvent = "Change") {
		v1 := toInteger(GuiCtrlObj["vel1"].value)
		v2 := toInteger(GuiCtrlObj["vel2"].value)
		absolute := Round((Max(v1,v2) - Min(v1,v2)) / Max(v1,v2), 2)
		percentage := Round(((Max(v1,v2) - Min(v1,v2)) / Max(v1,v2)) * 100, 1)
		GuiCtrlObj["RIresult"].value := "PSV: " . Max(v1, v2) . " cm/s; RI: " . absolute
	}
	if (A_GuiEvent = "Normal") {
		result := GuiCtrlObj["RIresult"].value
		WinActivate("KWS ahk_exe javaw.exe")
		Sleep(50)
		_KWS_PasteToReport(result, false)
		GuiCtrlObj.Destroy()
	}
	if (A_GuiEvent = "Close") {
		GuiCtrlObj.Destroy()
	}
}

openEAD_KWS(ead := "") {
	if RegExMatch(ead, "[^1-9]?(\d{8})[^1-9]?", &matchEAD) {
		A_Clipboard := ""
		Sleep(100) ;; 100 werkt dacht ik
		A_Clipboard := matchEAD[1]
		Errorlevel := !ClipWait(1)
		WinActivate("KWS ahk_exe javaw.exe")
		Send("^+z") ; Hotkey voor zoek patient
		Sleep(400)
		Send("{Ctrl down}v{Ctrl up}")
		Send("{Enter}")
	} else if (ead = "") {
		A_Clipboard := ""
		Sleep(100) ;; 150 werkt, te zien of lager werk
		Send("{Ctrl down}c{Ctrl up}") ;; zou ook reliablity verhogen
		Errorlevel := !ClipWait(1)
		openEAD_KWS(A_Clipboard)
	}
}

openLastPtInLog_KWS() {
	global logfile
	Loop read, logfile
		lastline := A_loopreadline
	if (RegExMatch(lastline, "[^1-9]?(\d{8})[^1-9]?", &matchEAD)) {
		;; MsgBox, %lastline% en %matchEAD1%
		Sleep(50) ;; needed for some reason ....
		openEAD_KWS(matchEAD[1])
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
	now := FormatTime(, "yyyyMMddHHmmss")
	;; age := CalcAge(birthdate . "000000", now)
	age := _CalcAge(birthdate, now)
	year := age[1]
	months := age[2]
	days := age[3]

	pedAbdGui := Gui()
	pedAbdGui.Opt("+LastFound")
	GuiHWND := WinExist()
	pedAbdGui.OnEvent("Close", pedAbdGuiCancelButton.Bind("Close"))
	pedAbdGui.OnEvent("Escape", pedAbdGuiCancelButton.Bind("Escape"))
	pedAbdGui.Add("Text", "x2 y19 w140 h20", "Leeftijd:")
	pedAbdGui.Add("Text", "x42 y39 w100 h20", "Jaar:")
	pedAbdGui.Add("Text", "x42 y59 w100 h20", "Maand:")
	pedAbdGui.Add("Text", "x42 y79 w100 h20", "Dag:")
	pedAbdGui.Add("Text", "x142 y-1 w230 h20")
	pedAbdGui.Add("Text", "x142 y39 w230 h20", age[1])
	pedAbdGui.Add("Text", "x142 y59 w230 h20", age[2])
	pedAbdGui.Add("Text", "x142 y79 w230 h20", age[3])
	pedAbdGui.Add("Text", "x2 y109 w130 h20", "Miltspan (mm):")
	pedAbdGui.Add("Text", "x2 y129 w130 h20", "Linkernier (mm):")
	pedAbdGui.Add("Text", "x2 y149 w130 h20", "Leverspan (mm):")
	pedAbdGui.Add("Text", "x2 y169 w130 h20", "Rechternier (mm):")
	ogcEditmilt := pedAbdGui.Add("Edit", "x132 y109 w100 h20 vmilt", "0")
	ogcEditmilt.OnEvent("Change", pedAbdGuiRefresh.Bind("Change"))
	ogcEditlinkerNier := pedAbdGui.Add("Edit", "x132 y129 w100 h20 vlinkerNier", "0")
	ogcEditlinkerNier.OnEvent("Change", pedAbdGuiRefresh.Bind("Change"))
	ogcEditlever := pedAbdGui.Add("Edit", "x132 y149 w100 h20 vlever", "0")
	ogcEditlever.OnEvent("Change", pedAbdGuiRefresh.Bind("Change"))
	ogcEditrechterNier := pedAbdGui.Add("Edit", "x132 y169 w100 h20 vrechterNier", "0")
	ogcEditrechterNier.OnEvent("Change", pedAbdGuiRefresh.Bind("Change"))
	ogcTextSDmilt := pedAbdGui.Add("Text", "x242 y109 w30 h20 vSDmilt", "0")
	SDmilt := ogcTextSDmilt.hwnd
	ogcTextSDliNier := pedAbdGui.Add("Text", "x242 y129 w30 h20 vSDliNier", "0")
	SDliNier := ogcTextSDliNier.hwnd
	ogcTextSDlever := pedAbdGui.Add("Text", "x242 y149 w30 h20 vSDlever", "0")
	SDlever := ogcTextSDlever.hwnd
	ogcTextSDreNier := pedAbdGui.Add("Text", "x242 y169 w30 h20 vSDreNier", "0")
	SDreNier := ogcTextSDreNier.hwnd
	ogcButtonOK := pedAbdGui.Add("Button", "x2 y199 w300 h30 Default", "OK")
	ogcButtonOK.OnEvent("Click", pedAbdGuiButtonOK.Bind("Normal"))
	pedAbdGui.Title := "Echografie Pediatrie Afmetingen"
	pedAbdGui.Show("x759 y391 h236 w305")
	; --------
	pedAbdGuiButtonOK(A_GuiEvent, GuiCtrlObj, Info := "", *) { ; V1toV2: Added bracket
		oSaved := pedAbdGui.Submit(1)
		result := _makePedReport(age, toInteger(oSaved.milt), toInteger(oSaved.linkerNier), toInteger(oSaved.lever), toInteger(oSaved.rechterNier))
		WinActivate("KWS ahk_exe javaw.exe")
		sleep(100)
		if (result != "")
			_KWS_PasteToReport(result, false)		; returning value
		pedAbdGui.Destroy()
	} ; V1toV2: Added Bracket before label
	; -------
	pedAbdGuiCancelButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
		pedAbdGui.Destroy()
	}
	pedAbdGuiRefresh(A_GuiEvent, GuiCtrlObj, Info := "", *) { ; V1toV2: Added bracket
		oSaved := pedAbdGui.Submit(0)
		SDs := _getStandardDevsPedAbd(age, toInteger(oSaved.milt), toInteger(oSaved.linkerNier), toInteger(oSaved.lever), toInteger(oSaved.rechterNier))
		; milt := ogcEditmilt.Text
		; linkerNier := ogcEditlinkerNier.Text
		; lever := ogcEditlever.Text
		; rechterNier := ogcEditrechterNier.Text
		; SDs := _getStandardDevsPedAbd(age, milt, linkerNier, lever, rechterNier)
		ogcTextSDmilt.Text := Round(SDs[3], 2)
		ogcTextSDliNier.Text := Round(SDs[1], 2)
		ogcTextSDlever.Text := Round(SDs[4], 2)
		ogcTextSDreNier.Text := Round(SDs[2], 2)
	}
} ; V1toV2: Added bracket before function

pressOKButton() {
	;; not really used, might be used in the future.
	#SuspendExempt
	MouseGetPos(&mouseX, &mouseY)
	if(ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\okButton.png")) {
		MouseClick("left", FoundX+5, FoundY+5)
		MouseMove(mouseX, mouseY)
	} else {
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000)
		return
	}
}

aanvaardOnderzoek(contrast := 2, opmerking := "") {
	MouseGetPos(&mouseX, &mouseY)
	If (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\contrastLabel.png")) {
		Switch contrast
		{
			Case 0: contrastX := FoundX + 80 ;; Zonder
			Case 1: contrastX := FoundX + 165 ;; Met
			Default: contrastX := FoundX + 250 ;; Zonder/Met
		}
		MouseClick("left", contrastX, FoundY)
		;; sleep, 100
		if (opmerking != "") {
			LabelFieldX := FoundX + 250
			LabelFieldY := FoundY + 200
			MouseClick("Left", LabelFieldX, LabelFieldY)
			Sleep(100)
			SendInput("^a" . opmerking)
		}
		MouseMove(mouseX, mouseY)
	}
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
	RegexQuery := "(?:Leuven|Pellenberg|Diest|Sint-Truiden)[\s\S]+(?<datum>\d{2}-\d{2}-\d{4})[\s\S]+(?:KLINISCHE INLICHTINGEN:[\n\r]+)(?<klinlicht>[\s\S]+)[\n\r]{2,}(?:(?:DIAGNOSTISCHE|RADIOLOGISCHE) VRAAGSTELLING:[\n\r]+)(?<diagvraag>[\s\S]+)[\n\r]{2,}(?:ONDERZOEKE?N?:[\n\r]+)(?<onderzoek>(?:.+\R{0,1}?)+)\R{2,}(?<content>[\s\S]+)"

	RegExMatch(A_Clipboard, RegexQuery, &report)
	ead := _getEAD()
	_BlockUserInput(false)
	if FileExist(excelSavePath) {
		if not WinExist("ahk_exe EXCEL.EXE") {
			Run(excelSavePath)
			_makeSplashText(title := "Opening excel", text := "Opening " . excelSavePath, time := 3000)
			ErrorLevel := WinWait("ahk_exe EXCEL.EXE") , ErrorLevel := ErrorLevel = 0 ? 1 : 0
		}
	} else {
		ErrorLevel := WinWait("ahk_exe EXCEL.EXE") , ErrorLevel := ErrorLevel = 0 ? 1 : 0
		MsgBox("Excel file niet gevonden! Maak een nieuw excel bestand in dezelfde folder als dit script met de naam " excelSavePath "`nEr zal nu geprobeerd een te maken (niet getest)...")
		XL := ComObject("Excel.Application")
		XL.Workbooks.Add
		XL.ActiveWorkbook.SaveAs(excelSavePath)
	}
	categoryList := ["","neuro","thorax","abdomen","spine","MSK","NKO","mammo","urogen","vasc","cardio"]
	indexCategory := 1
	category := ""
	Switch
	{
		case RegExMatch(report["onderzoek"], "i)hersen|schedel|hypophyse"):
		indexCategory := 2
		case RegExMatch(report["onderzoek"], "i)abdomen|lever|pancreas"):
		indexCategory := 4
		case RegExMatch(report["onderzoek"], "i)thorax|longen|embol"):
		indexCategory := 3
		case RegExMatch(report["onderzoek"], "i)wervel"):
		indexCategory := 5
		case RegExMatch(report["onderzoek"], "i)knie|schouder|heup|arm|been|pols|hand"):
		indexCategory := 6
		case RegExMatch(report["onderzoek"], "i)oor|hals|rots|schild|speeksel"):
		indexCategory := 7
		case RegExMatch(report["onderzoek"], "i)mammo|borst"):
		indexCategory := 8
		case RegExMatch(report["onderzoek"], "i)nier|blaas|prostaat|gyna|vrouw|bekken"):
		indexCategory := 9
		case RegExMatch(report["onderzoek"], "i)vascul"):
		indexCategory := 10
		case RegExMatch(report["onderzoek"], "i)hart|coronair"):
		indexCategory := 11

	}
	ExcelGUI := Gui()
	ExcelGUI.OnEvent("Close", ExcCancelButton.Bind("Close"))
	ExcelGUI.OnEvent("Escape", ExcCancelButton.Bind("Escape"))
	ExcelGUI.Opt("+LastFound")
	ExcelGuiHWND := WinExist()
	ExcelGUI.Add("Text", "x10 y10 w60 h30", report["datum"])
	ExcelGUI.Add("Text", "xp+70 yp+0 w60 h30", ead)
	ogcDropDownListExcCategorySelect := ExcelGUI.Add("DropDownList", "xp+70 yp+0 w130 R1 vExcCategorySelect R11 Choose" . indexCategory, categoryList)
	ogcEditExcOnderzoek := ExcelGUI.Add("Edit", "x10  yp+30 R1 vExcOnderzoek", report["onderzoek"])
	ExcelGUI.Add("Text", "x10 yp+40 w130 R2", "Comment")
	ogcEditExcComment := ExcelGUI.Add("Edit", "xp+140 yp+0  w190 R2 vExcComment")
	ExcelGUI.Add("Text", "xp-140 yp+40 w130 R2", "Klin. inlichtingen")
	ogcEditExcKlinInl := ExcelGUI.Add("Edit", "xp+140 yp+0  w190 R2 vExcKlinInl", report["klinlicht"])
	ExcelGUI.Add("Text", "xp-140 yp+40 w130 R2", "Diagn. vraagstelling")
	ogcEditExcDiagnVraag := ExcelGUI.Add("Edit", "xp+140 yp+0  w190 R2 vExcDiagnVraag", report["diagvraag"])
	ExcelGUI.Add("Text", "xp-140 yp+40 w130 R2", "Tags")
	ogcEditExcTags := ExcelGUI.Add("Edit", "xp+140 yp+0  w190 R2 vExcTags")
	ExcelGUI.Add("Text", "xp-140 yp+40 w130 R2", "Op te volgen?")
	ogcDropDownListOpTeVolgen := ExcelGUI.Add("DropDownList", "xp+140 yp+0 w190 R5 vOpTeVolgen Choose1", ["", "Over 1 dag", "Over 1 week", "Over 1 maand", "Op te volgen"])
	ogcButtonOK := ExcelGUI.Add("Button", "x10    yp+40 w130 h20 Default", "OK")
	ogcButtonOK.OnEvent("Click", ExcelGUIButtonOK.Bind("Normal"))
	ogcButtonCancel := ExcelGUI.Add("Button", "xp+140 yp+0 w130 h20", "Cancel")
	ogcButtonCancel.OnEvent("Click", ExcCancelButton.Bind("Normal"))
	ExcelGUI.Title := "Save to excel script"
	ExcelGUI.Show("x360 y233")
	; ErrorLevel := WinWaitClose("ahk_id " ExcelGuiHWND,,)
;; todo: add a check to see if excel has been found and XL object initialized
	; return
	ExcelGUIButtonOK(A_GuiEvent, GuiCtrlObj, Info := "", *)		{ ; V1toV2: Added bracket
		XL := ComObjGet(excelSavePath) ;; looks for excel
		oSaved := ExcelGUI.Submit()
		ExcCategorySelect := oSaved.ExcCategorySelect
		ExcOnderzoek := oSaved.ExcOnderzoek
		ExcComment := oSaved.ExcComment
		ExcKlinInl := oSaved.ExcKlinInl
		ExcDiagnVraag := oSaved.ExcDiagnVraag
		ExcTags := oSaved.ExcTags
		OpTeVolgen := oSaved.OpTeVolgen
		lastCell := excelFindLastCell(XL).row + 1

		if (InStr(OpTeVolgen, "Over")) { ;; This function finds the future date to follow up on
			FutureDate := A_Now
			Switch
			{
				case InStr(OpTeVolgen, "day"):   FutureDate := DateAdd(FutureDate, "1", "days")
				case InStr(OpTeVolgen, "week"):  FutureDate := DateAdd(FutureDate, "7", "days")
				case InStr(OpTeVolgen, "month"): FutureDate := DateAdd(FutureDate, "32", "days")
			}
			opTeVolgen := FormatTime(FutureDate, "yyy-MM-dd")
		}

		XL.Application.ActiveSheet.range("A" . lastCell).value := ead
		XL.Application.ActiveSheet.range("B" . lastCell).value := report["datum"]
		XL.Application.ActiveSheet.range("C" . lastCell).value := ExcCategorySelect
		XL.Application.ActiveSheet.range("D" . lastCell).value := ExcComment
		XL.Application.ActiveSheet.range("E" . lastCell).value := opTeVolgen
		XL.Application.ActiveSheet.range("F" . lastCell).value := RegExReplace(ExcOnderzoek, "\.?[\r\n]", ". ")
		XL.Application.ActiveSheet.range("G" . lastCell).value := RegExReplace(ExcKlinInl, "\.?[\r\n]", ". ")
		XL.Application.ActiveSheet.range("H" . lastCell).value := RegExReplace(ExcDiagnVraag, "\.?[\r\n]", ". ")
		XL.Application.ActiveSheet.range("I" . lastCell).value := ExcTags
		_makeSplashText(title := "Saved to excel", text := ead . " is saved to excel", time := -1500)
		ExcelGUI.Destroy()
	} ; V1toV2: Added Bracket before label

	ExcCancelButton(A_GuiEvent, GuiCtrlObj, Info := "", *) { ; V1toV2: Added bracket
		ExcelGUI.Destroy()
	}
} ; V1toV2: Added bracket before function

excelFindLastCell(objExcel, sheet := 1) {
	static xlByRows    := 1	     , xlByColumns := 2	     , xlPrevious  := 2
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
	A_Clipboard := ""
	Send("{End}")
	Send("+{Home}")
	Send("+{Left}")
	Send("^x")
	Errorlevel := !ClipWait(1)
	Send("{Up}")
	Send("{End}")
	Send("^v")
	Return
}

MoveLineDown() {
	A_Clipboard := ""
	Send("{End}")
	Send("+{Home}")
	Send("+{Left}")
	Send("^x")
	Errorlevel := !ClipWait(1)
	Send("{Down}")
	Send("{End}")
	Send("^v")
	Return
}

IndentLine(direction) {
	temp := A_Clipboard
	A_Clipboard := ""
	Send("{Home}")
	Send("+{End}")
	Send("^x")
	Errorlevel := !ClipWait(1)
	line := A_Clipboard
	A_Clipboard := ""
	if (direction = "promote") {
		line := RegExReplace(line, "^\s*[\.\-]?\s*(.+)$", "- $1")
	} else if (direction = "demote") {
		line := RegExReplace(line, "^\s*[\.\-]?\s*(.+)$", "  . $1")
	}
	A_Clipboard := line
	Errorlevel := !ClipWait(1)
	if (GetKeyState("{Alt}","P")) {
		Send("{Alt Up}")
	}
	Send("^v")
}


deleteLine() {
	Send("{End}")
	Send("+{Home}")
	Send("{Delete}{Delete}")
}

yankLine() {
	Send("{Home}")
	Send("+{Down}")
	Send("^c")
}

insertDatePeriod(daysInFuture := 0) {
	Send("{Home}") ;; zorgt dat de functie eender waar in het datumvak gestart kan worden
	CurrentDateTime := FormatTime(, "ddMMyyyy")
	SendInput(CurrentDateTime "0000")
	if (daysInFuture >= 0) {
		; FutureDate := A_Now
		FutureDate := DateAdd(A_Now, daysInFuture, "days")
		morgen := FormatTime(FutureDate, "ddMMyyyy")
		sleep(10)
		Send("{tab}")
		sleep(10)
		SendInput(morgen "2359")
		;; MsgBox, %morgen%2359
	}
}

insertPastDatePeriod(daysInPast := 1) {
	Send("{Home}") ;; zorgt dat de functie eender waar in het datumvak gestart kan worden
	daysInPast := -1 * Abs(daysInPast)
	;;PastDate := A_Now
	Pastdate := DateAdd(A_Now, daysInPast, 'Days')
	PastDate := FormatTime(PastDate, "ddMMyyyy")
	SendInput(PastDate "0000")
	sleep(10)
	Send("{tab}")
	sleep(10)
	CurrentDateTime := FormatTime(, "ddMMyyyy")
	SendInput(CurrentDateTime "0000")
}


initKWSWindows() {
	;; Uitvoeringswerklijst
	Send("^{Space}")
	Sleep(150)
	SendInput("Uitvoeringswerklijst")
	Sleep(400)
	Send("{Enter}")
	ErrorLevel := WinWait("Uitvoeringswerklijst parameters ahk_class SunAwtDialog", , 3) , ErrorLevel := ErrorLevel = 0 ? 1 : 0
	if WinExist("Uitvoeringswerklijst parameters ahk_class SunAwtDialog") {
		WinActivate()
		Sleep(450)
		Send("{Tab}")
		Send("e")
		Sleep(350)
		Send("{Enter}")
	}

	;; receptieeenheid
	Send("^{Space}")
	Sleep(150)
	SendInput("zet receptie eenheid")
	Sleep(450)
	Send("{Enter}")
	ErrorLevel := WinWait("Kies een receptie eenheid ahk_class SunAwtDialog", , 3) , ErrorLevel := ErrorLevel = 0 ? 1 : 0
	if WinExist("Kies een receptie eenheid ahk_class SunAwtDialog") {
		WinActivate()
		Sleep(450)
		MouseGetPos(&mouseX, &mouseY)
		if ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\EenheidLabel.png") {
			MouseClick("Left", FoundX+60, FoundY+5)
			MouseMove(mouseX, mouseY)
		} else {
			_makeSplashText("Error", "Something went wrong when looking for the field", time := -2000)
			return
		}
		SendInput(555)
		Send("{Enter}")
		Sleep(150)
		Send("{Tab}")
		Sleep(450)
		Send("{Enter}")
		Sleep(600)
	}

	;; receptie lijst
	Send("^{Space}")
	Sleep(150)
	SendInput("receptie werklijst hos")
	Sleep(400)
	Send("{Enter}")
	ErrorLevel := WinWait("parameters voor receptiewerklijst ahk_class SunAwtDialog", , 3) , ErrorLevel := ErrorLevel = 0 ? 1 : 0
	if WinExist("parameters voor receptiewerklijst ahk_class SunAwtDialog") {
		WinActivate()
		Sleep(450)
		;; Send, {Down}
		Send("{Tab}")
		Send("{Down}{Down}")
		Send("{Tab}")
		Sleep(150)
		insertDatePeriod(0)
		Sleep(250)
		Send("{Enter}")
		Sleep(2200)
	}
}

; HELPER FUNCTIONS
; --------------------------------------

_KWS_CopyReportToClipboard(selectReportBox := True) {
	_BlockUserInput(true)
	If (not WinActive("KWS ahk_exe javaw.exe")) {
		if WinExist("KWS ahk_exe javaw.exe")
			WinActivate()
		else
			throw Error("KWS is not open!", -1)						; STOPT als geen data in clipboard
	}
	if (selectReportBox) {
		_KWS_SelectReportBox()
	}
	A_Clipboard := ""				; maakt het clipboard leeg
	Send("^a")					; select all
	Send("{Ctrl down}c{Ctrl up}") ;; zou ook reliablity verhogen
	_BlockUserInput(false)
	Errorlevel := !ClipWait(1)					; wacht tot er data in het clipboard is
	if (ErrorLevel)					; als NOT, is er data in clipboard
		throw Error("Could not copy data to A_Clipboard!", -1)						; STOPT als geen data in clipboard
}

_KWS_SelectReportBox(mousebutton := "left") {
	;; assumes KWS already active
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\bevindingenLabel.png")
	if (ErrorLevel = 2)
		_makeSplashText(title := "Error", text := "Something went wrong when looking for enter field", time := -2000)
	else if (ErrorLevel = 1)
		_makeSplashText(title := "Error", text := "Error finding the report[0] field: did the action succeed nonetheless?", time := -2000)
	_BlockUserInput(true)
	MouseGetPos(&mouseX, &mouseY)
	MouseClick(mousebutton, FoundX+100, FoundY+200)
	MouseMove(mouseX, mouseY)
	_BlockUserInput(false)
}

_KWS_PasteToReport(text, overwrite := true) {
	If WinActive("KWS ahk_exe javaw.exe") {
		_BlockUserInput(True)
		tempclip := A_Clipboard
		A_Clipboard := ""
		A_Clipboard := text		; maakt het clipboard leeg
		Errorlevel := !ClipWait(1)			; wacht tot er data in het clipboard is
		if (overwrite) {
			Send("^a")
			Sleep(50)
		}
		Send("{Ctrl down}v{Ctrl up}") ;; zou ook reliablity verhogen
		Sleep(50)
		if WinExist("Foutboodschap JavaKWS") { ; Fixes "could not access clipboard"
			WinActivate()
			SendInput("{Enter}")
			Sleep(50)
			WinActivate("KWS ahk_exe javaw.exe")
			_KWS_PasteToReport(text, overwrite)
		}
		A_Clipboard := tempclip
		_BlockUserInput(false)
		return
	}
	if WinExist("KWS ahk_exe javaw.exe")
		WinActivate()
	else
		throw Error("KWS is not open!", -1)
	_KWS_SelectReportBox()
	_KWS_PasteToReport(text, overwrite)
	return
}

_sorttext(inputtext) {
	if (RegExMatch(inputtext, "([\s\S]+)(\RCONCLUSIE[\s\S]+)", &split)) {
		outputtext := _sorttext(split[1]) . split[2]
		return outputtext
	} else if (RegExMatch(inputtext, "([\s\S]+)(\*.+)([\s\S]+)", &split)) {
		outputtext := _sorttext(split[1]) . split[2] . _sorttext(split[3])
		return outputtext
	} else {
		dotlines := ""
		;; TODO: evt via Loop, Parse, inputtext, "`n`r" en dan A_LoopField
		while (RegExMatch(inputtext, "m)(?:\R|^)\.?\ ?(.+#.?)$", &line)) {	;gets all lines with XXXXX #
			dotlines := dotlines . "`n. " . line[1]
			inputtext := StrReplace(inputtext, line[0], "")
		}
		while (RegExMatch(inputtext, "m)(?:\R|^)\.\ ?(.+[^\/])?$", &line)) {		;gets all lines with . XXXXX (as long as not ending with /)
		; while (RegExMatch(inputtext, "m)(?:\R|^)\.\ ?(.+)$" , line)) {		;gets all lines with . XXXXX
			dotlines := dotlines . "`n. " . line[1]
			inputtext := StrReplace(inputtext, line[0], "")
		}
		outputtext := inputtext . "`n" . StrReplace(StrReplace(dotlines, "#", ""), ". - ", ". ") . "`n`n"
		return outputtext
	}
}

_getEAD(returnMouse := false) {
	if (returnMouse) {
		MouseGetPos(&mouseX, &mouseY)
	}
	temp := A_Clipboard
	A_Clipboard := ""
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\eadnrLabel.png")
	_BlockUserInput(true)
	MouseClick("left", FoundX+70, FoundY+10)
	_BlockUserInput(false)
	Errorlevel := !ClipWait(1)
	ead := A_Clipboard
	A_Clipboard := temp
	if (returnMouse) {
		MouseMove(mouseX, mouseY)
	}
	return ead
}

_getBirthDate(returnMouse := false) {
	if (returnMouse) {
		MouseGetPos(&mouseX, &mouseY)
	}
	temp := A_Clipboard
	A_Clipboard := ""
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\eadnrLabel.png")
	_BlockUserInput(True)
	MouseClick("left", FoundX-25, FoundY+10)
	Errorlevel := !ClipWait(1)
	date := SubStr(A_Clipboard, 2)
	A_Clipboard := temp
	if (returnMouse) {
		MouseMove(mouseX, mouseY)
	}
	_BlockUserInput(false)
	return date
}

_makeSplashText(title := "Splash title", text := "Splash text", time := -3000, doublePressMode := false) {
	time := Abs(time) * (-1) ;; Zorgt dat time altijd negatief is (dat removesplash dus maar 1 keer wordt uitgevoerd ipv in loop)
	global splashExists
	global splashGui

	_destroySplash()
	splashExists := doublePressMode
	splashGui := Gui()
	splashGui.Opt("+AlwaysOnTop +Disabled -SysMenu +Owner")		; +Owner avoids a taskbar button.
	splashGui.Add("Text", , text)
	splashGui.Title := title
	splashGui.Show("NoActivate")		; NoActivate prevents taking the focus
	;; destroySplash_fn := _destroySplash.bind() ; Settimer aanvaard enkel een label, iets wat ik probeer te vermijden. Op deze manier kan ik toch een functie aan setTimer geven.
	SetTimer(_destroySplash.bind(), time)
	return
}

_destroySplash() {
	global splashExists
	global splashGui
	splashExists := False
	if (not isSet(splashGui) or splashGui = "") {
		return
	}
	splashGui.Destroy()
}

_calcHeightLoss(h1, h2) {
	absolute := Max(h1,h2) - Min(h1,h2)
	percentage := Round((1 - (Min(h1,h2) / Max(h1,h2))) * 100)
	return [absolute, percentage]
}

auto_scroll(richting := 1, decreaseKey := "&", increaseKey := "Ã©", directionKey := "Space", pauseKey := "`""){ ; Automatisch scrollen. Versnel & vertraag. Gemaakt door Johannes Devos
	Suspend(true)
	; Hotkey, If
	; try Hotkey, %decreaseKey%, Off
	; try Hotkey, %increaseKey%, Off
	; try Hotkey, %directionKey%, Off
	MouseGetPos(, , &windowUnderCursor)
	temp := WinGetProcessName("ahk_id " windowUnderCursor)
	If (temp = "impax-client-main.exe" OR temp = "syngo.Common.Container.exe" OR temp = "javaw.exe" OR temp = "javawClinapps.exe"){
		keys := "{" . decreaseKey  . "}{" . increaseKey . "}{" . directionKey . "}{" . pauseKey . "}"
		hook := InputHook("L0", keys)
		hook.VisibleNonText := false
		hook.Start()
		static sleep_delay := 300
		endloop := 0
		hook.OnChar := _auto_scroll_down_helper
		Loop{
			if not GetKeyState(pauseKey) {
				if (richting = 1)
					Send("{wheeldown 1}")
				if (richting = -1)
					Send("{wheelup 1}")
			}
			Sleep(sleep_delay)
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
		Suspend(false)
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
	if not IsSet(logfile)
		logfile := "logfile.csv"
	timestring := FormatTime("A_Now", "yyyy-MM-dd HH:mm:ss")
	output := timestring  . "|" . str
	for index, s in extrastr
		output := output . "|" . s
	if FileExist(logfile)
		output := "`n" . output
	FileAppend(output, logfile)
}

_KWS_closeMededelingen() {
	if Winexist("Mededelingen ahk_exe javaw.exe") {
		WinClose()
	}
}

clipboardcleaner() {
	Errorlevel := !ClipWait(1)					; wacht tot er data in het clipboard is
	Sleep(50)
	if (RegExMatch(A_Clipboard, "(verslag \*\*[\r\n]{2,})([\s\S]+?)([\r\n]{2,}\*\* Einde)", &KWSfiltered)) {
		A_Clipboard := KWSfiltered[2]
		return
	} else {
		Return
	}
}

_BlockUserInput(block := true) {
	if (block) {
		BlockInput("On")
		MouseClick("left", 0, 0, , , "U", "R") ;; zorgt dat als de muis ingeduwd was dat er geen error komt
		;; BlockInput, Mousemove
		BlockInput("Send")
	} else {
		BlockInput("Off")
		BlockInput("Default")
		;; BlockInput, MouseMoveOff
	}
	;; Settimer die de blokkage opheft na 1 seconde, voor moest het programma crashen of vastlopen
	SetTimer(_BlockUserInput.bind(false),-1000)
}

_MouseIsOver(vWinTitle:="", vWinText:="", vExcludeTitle:="", vExcludeText:="") {
	MouseGetPos(, , &hWnd)
	return WinExist(vWinTitle (vWinTitle=""?"":" ") "ahk_id " hWnd, vWinText, vExcludeTitle, vExcludeText)
}

findAndReplaceGUI() {
	active_id := WinGetID("A") ;; gets the window where the script was activated
	A_Clipboard := ""
	Send("{Ctrl down}c{Ctrl up}") ;; zou ook reliablity verhogen
	ClipWait(1)
	originalText := A_Clipboard
	repGUI := Gui()
	repGUI.Opt("+LastFound -DPIscale")
	repGuiHWND := WinExist()
	repGUI.OnEvent("Close", repGuiCancel.Bind("Close"))
	repGUI.OnEvent("Escape", repGuiCancel.Bind("Close"))
	repGUI.MarginX := "10", repGUI.MarginY := "10"
	repGUI.Add("Text", "w130", "Find")
	ogcEditfindText := repGUI.Add("Edit", "w500 R2  vfindText")
	ogcEditfindText.OnEvent("Change", updateRepGUI.Bind("Change"))
	repGUI.Add("Text", "w130", "Replace")
	ogcEditreplaceText := repGUI.Add("Edit", "w500 R2  vreplaceText")
	ogcEditreplaceText.OnEvent("Change", updateRepGUI.Bind("Change"))
	ogcButtonReplace := repGUI.Add("Button", , "Replace!")
	ogcButtonReplace.OnEvent("Click", ReplaceButton.Bind("Normal"))
	ogcCheckboxRegexToggle := repGUI.Add("Checkbox", "yp+5 xp+65 vRegexToggle checked", "Enable regex")
	ogcCheckboxRegexToggle.OnEvent("Click", updateRepGUI.Bind("Normal"))
	ogcCheckboxIgnoreCaseFlag := repGUI.Add("Checkbox", "yp+0	xp+90 vIgnoreCaseFlag checked", "IgnoreCase")
	ogcCheckboxIgnoreCaseFlag.OnEvent("Click", updateRepGUI.Bind("Normal"))
	ogcCheckboxMultilineFlag := repGUI.Add("Checkbox", "yp+0	xp+90 vMultilineFlag  checked", "Multiline-mode")
	ogcCheckboxMultilineFlag.OnEvent("Click", updateRepGUI.Bind("Normal"))
	ogcCheckboxSinglelineFlag := repGUI.Add("Checkbox", "yp+0	xp+90 vSinglelineFlag", "Singleline")
	ogcCheckboxSinglelineFlag.OnEvent("Click", updateRepGUI.Bind("Normal"))
	ogcCheckboxUngreadyFlag := repGUI.Add("Checkbox", "yp+0	xp+70 vUngreadyFlag", "Ungready")
	ogcCheckboxUngreadyFlag.OnEvent("Click", updateRepGUI.Bind("Normal"))
	ogcActiveXrepTextBox := repGUI.Add("ActiveX", "vrepTextBox x10 w500 h400", "htmlfile")
	repTextBox := ogcActiveXrepTextBox.Value
	repTextBox.write(_getHTMLReplaceBox(originalText, "", "", False))
	repGUI.Title := "Find and replace"
	repGUI.Show()
	;-------
	ReplaceButton(A_GuiEvent, GuiCtrlObj, Info := "", *) { ; V1toV2: Added bracket
		needle := ogcEditfindText.Text
		replacement := ogcEditreplaceText.Text
		flag := _findreplaceConstructRegexFlags(ogcCheckboxIgnoreCaseFlag.Text, ogcCheckboxMultilineFlag.Text, ogcCheckboxSinglelineFlag.Text, ogcCheckboxUngreadyFlag.Text)
		temp := A_Clipboard
		A_Clipboard := ""

		if (RegExMatch(originalText, "(verslag \*\*[\r\n]{4})([\s\S]+)([\r\n]{4}\*\* Einde tekst)", &KWSfiltered)) {
			originalText := KWSfiltered[2]
		}
		if (ogcCheckboxRegexToggle.Text) {
			A_Clipboard := RegExReplace(originalText, flag . needle, replacement)
			if (A_Clipboard = "")
				A_Clipboard := originalText
		} else {
			A_Clipboard := StrReplace(originalText, needle, replacement)
		}
		WinActivate("ahk_id " active_id)
		ClipWait(1)
		Sleep(400)
		Send("^v")
		Sleep(100)
		A_Clipboard := temp
		repGui.Destroy()

	} ; V1toV2: Added bracket before function

	repGuiCancel(A_GuiEvent, GuiCtrlObj, Info := "", *) {
		repGui.Destroy()
	}
	updateRepGUI(A_GuiEvent, GuiCtrlObj, Info := "", *) { ; V1toV2: Added bracket
		flag := _findreplaceConstructRegexFlags(ogcCheckboxIgnoreCaseFlag.Text, ogcCheckboxMultilineFlag.Text, ogcCheckboxSinglelineFlag.Text, ogcCheckboxUngreadyFlag.Text)
		repTextBox.open()
		repTextBox.write(_getHTMLReplaceBox(originalText, ogcEditfindText.Text, ogcEditreplaceText.Text, ogcCheckboxRegexToggle.Text, flag))
		repTextBox.close()
	}
} ; V1toV2: Added bracket before function

_getHTMLReplaceBox(haystack, needle, replacement, regextoggle := True, flag := "") {
	; https://www.autohotkey.com/boards/viewtopic.php?t=84074
	html := "
	(
		<style>
		body {
			font-family: calibri;
			font-size: 13px;
			white-space: pre-line;
			overflow-wrap: normal;
		}		.red {
			color: red;
			font-weight: bold;
			text-decoration: line-through;
		}		.replacement {
			color: blue;
			font-weight: bold;
		}
		</style>
		<body>
		<div>INSERTLOCATION</div>
		</body>
	)"
	;; TODO werkt niet
	;; Haalt "de zin van "uit ongevalideerd verslag" weg indien aanwezig
	if (RegExMatch(haystack, "(verslag \*\*[\r\n]{4})([\s\S]+)([\r\n]{4}\*\* Einde)", &KWSfiltered)) {
		haystack := KWSfiltered[2]
	}
	if (regextoggle) {
		; TODO: regexreplace in try / catch, met als catch gewoon strReplace met "regex error"
		try {
			replacedText := RegExReplace(haystack, flag . needle, "<span class=`"red`">$0</span><span class=`"replacement`">" . replacement . "</span>")
		} catch Error as e {
			replacedText := "<span class=`"replacement`">---------------`nError in Regex code; fix or disable regex`n---------------</span>`n`n" . haystack
		}
	}
	else {
		replacedText := StrReplace(haystack, needle, "<span class=`"red`">" . needle . "</span><span class=`"replacement`">" . replacement . "</span>")
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
; REMOVED:	SetFormat, Float, 0.1
	Verslag := "Normale ligging van de retroperitoneale grote vaten.`nNormale ligging van de organen.`n`nLeverspan: " . Round(Lever/10, 1) . " cm (SD: " . Round(Result[4], 2) . ").`nHomogeen leverparenchym met normale reflectiviteit.`nNormale portahoofdstam en intrahepatische portatakken.`nNormale hepatische venen met normale hepatofugale flow.`nNormale hepatopetale portale flow.`nNormale flow in de a. hepatica.`nGeen gedilateerde intrahepatische of extrahepatische galwegen aangetoond.`nNormale galblaas.`nNormale pancreas. Geen visualisatie van de ductus van Wirsung.`nMilt: " . Round(Milt/10, 1) . " cm (SD: " . Round(Result[3], 2) . ").`nNormale milt.`n`nNormale bijnieren en bijnierloges.`nLinkernier: " . Round(Linkernier/10, 1) . " cm (SD: " . Round(Result[1],2) . ").`nRechternier: " . Round(Rechternier/10, 1) . " cm (SD: " . Round(Result[2], 2) . ").`nNormale reflectiviteit van het nierparenchym met corticomedullaire differentiatie.`nGeen hydro-ureteronefrose.`nNormale blaasvulling.`nNormale aflijning en dikte van de blaaswand.`n`nNormale ligging van de A. en V. Mesenterica Superior.`nGeen adenopathieen aangetoond.`nNormale darmwanden.`n###Normaal terminale ileum.`n###Normale appendix.`n`nCONCLUSIE:`n###`n`nGECOMMUNICEERDE DRINGENDE BEVINDINGEN:`n"
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
; REMOVED:	SetFormat, Float, 0.2

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
	FromDay := SubStr(FromDay, 1, 8)
	ToDay := SubStr(ToDay, 1, 8)
	Years := 0
	Months := 0
	Days := 0
	;; Global Years,Months,Days
	; If born on February 29
	If SubStr(FromDay, 5, 4) = 0229 and Mod(SubStr(ToDay, 1, 4), 4) != 0 and SubStr(ToDay, 5, 4) = 0228
		PlusOne := "1"
	ThisMonth := SubStr(ToDay, 1, 6)
	; Set ThisMonthLength equal to next month
	ThisMonthLength := SubStr(ToDay, 5, 2) = "12" ? SubStr(ToDay, 1, 4)+1 . "01" : SubStr(ToDay, 1, 4) . SubStr("0" . SubStr(ToDay, 5, 2)+1, -2)
	; Days in this month saved in ThisMonthLength
	ThisMonthLength := DateDiff(ThisMonthLength, ThisMonth, "d")
	; Set ThisMonthday to FromDay or  (if FromDay higher) last day of this month
	If SubStr(FromDay, 7, 2) > ThisMonthLength
		ThisMonthDay :=  ThisMonth . ThisMonthLength
	Else
		ThisMonthDay :=  ThisMonth . SubStr(FromDay, 7, 2)
	; Calculate last month's length
	LastMonthLength := SubStr(ToDay, 5, 2) = "01" ? SubStr(ToDay, 1, 4)-1 . "12" : SubStr(ToDay, 1, 4) . SubStr("0" . SubStr(ToDay, 5, 2)-1, -2)
	LastMonth := LastMonthLength
	; Days in last month saved in LastMonthLength
	LastMonthLength := DateDiff(LastMonthLength, ThisMonth, "d")
	LastMonthLength := LastMonthLength*(-1)
	; Set LastMonthday to FromDay or (if FromDay higher) last day of last month
	If SubStr(FromDay, 7, 2) > LastMonthLength
		LastMonthDay :=  LastMonth . LastMonthLength
	Else
		LastMonthDay :=  LastMonth . SubStr(FromDay, 7, 2)
	; Calculate years
	Years  := SubStr(ToDay, 5, 4) - SubStr(FromDay, 5, 4) < 0 ? SubStr(ToDay, 1, 4)-SubStr(FromDay, 1, 4)-1 : SubStr(ToDay, 1, 4)-SubStr(FromDay, 1, 4)
	; Calculate months
	Months := SubStr(ToDay, 5, 2)-SubStr(FromDay, 5, 2) < 0 ? SubStr(ToDay, 5, 2)-SubStr(FromDay, 5, 2)+12	: SubStr(ToDay, 5, 2)-SubStr(FromDay, 5, 2)
	Months := SubStr(ToDay, 7, 2) - SubStr(ThisMonthDay, 7, 2) < 0 ? Months -1 : Months
	Months := Months = -1 ? 11 : Months
	; Calculate days
	TodayDate := SubStr(ToDay, 1, 8)          ; Remove any time portion of stamp
	ThisMonthDay := DateDiff(ThisMonthDay, ToDayDate, "d")
	LastMonthDay := DateDiff(LastMonthDay, ToDayDate, "d")
	Days  := ThisMonthDay <= 0 ? -1*ThisMonthDay : -1*LastMonthDay
	; If February 28
	Years := isSet(PlusOne) ? Years + PlusOne : Years
	days := isSet(PlusOne) ? 0 : days
	If (TodayDate <= FromDay)
		Years := 0, Months := 0,Days := 0
	age := [Years, Months, days]
	return age
}
