;****************************************************************
;								*
;     		Luie radiologen TIRADS hulpje			*
;		TIRADS calculator + verslag maker 		*
;								*
;****************************************************************
;								*
; 	Auteur: Cedric Vanmarcke				*
; 								*
; 	Handleiding: zie "readme.txt" bestand			*
;								*
; 	Dit programma is bedoeld voor iedereen die 		*
;		het kan en wil gebruiken			*
;	Voor vragen: cedric.vanmarcke@uzleuven.be		*
;								*
;****************************************************************

NoduleList := []

Gui, Add, GroupBox, x10 w300 R2, Location
Gui, Add, DropDownList, xp+10 yp+20 w70 vSide gCheck, linker|rechter|isthmus
Gui, Add, DropDownList, xp+80 w80 vLocation gCheck, bovenpool|middenpool|onderpool

Gui, Add, GroupBox, x10 w300 R6, Characteristics
Gui, Add, DropDownList, xp+10 yp+20 w250 vComposition gCheck Choose1, cystisch (0 punten)|spongiform (0 punten)|gemengd cystisch en solide (1 punt)|solide of bijna volledig solid (2 punten)
Gui, Add, DropDownList, w100 w250 vEchogenicity gCheck Choose1, anechogeen (0 punten)|hyperechogeen of isoechogeen (1 punt)|hypoechogeen (2 punten)|zeer hypoechogeen (3 punten)
Gui, Add, DropDownList, w100 w250 vShape gCheck Choose1, wider-than-tall (0 punten)|taller-than-wide (3 punten)
Gui, Add, DropDownList, w100 w250 vMargin gCheck Choose1, scherp en glad (0 punten)|wazig afgelijnd (0 punten)|gelobuleerd of onregelmatig (2 punten)|extra-thyroidale extensie (3 punten)


Gui, Add, GroupBox, x10 w300 r4, Echogenic Foci (check all that apply)
;; TODO: label met dan dropdown nee (default) of ja
Gui, Add, CheckBox, xp+10 yp+20 gCheck vFoci1, Comet-tail artifacten (0 punten)
Gui, Add, CheckBox, gCheck vFoci2, Macrocalcificaties (1 punt)
Gui, Add, CheckBox, gCheck vFoci3, Perifere (rim) verkalkingen (2 punten)
Gui, Add, CheckBox, gCheck vFoci4, Punctiforme echogene foci (3 punten)

Gui, Add, GroupBox, xp-10 yp+20 w300 h50 r1, Size (ML x AP x CC) in mm
Gui, Add, Edit, xp+10 yp+20 w30 Number vSizeML gCheck,
Gui, Add, Edit, xp+40 yp+0 w30 Number vSizeAP gCheck,
Gui, Add, Edit, xp+40 yp+0 w30 Number vSizeCC gCheck,

Gui, Add, Text, x20 yp+30 vTiradsGradeText HwndTiradsGradeText, TI-RADS grade: 0

;; TODO: knop: add new nodule, wordt dan aan lijst toegevoegd en reset.
Gui, Add, Button, x10 w145 h40 gAddNoduleButton, Add another nodule
Gui, Add, Button, xp+155 yp+0 w145 h40 gAddNoduleAndCopyButton, Add nodule and copy to KWS
Gui, Add, Button, xp+155 yp+0 w145 h40 gInsertButton, Copy to KWS!
;; Gui, Add, Button, xp+85 yp+0 w80 h20 gResetNoduleButton, Reset nodule
;; Gui, Add, CheckBox, x10 vAutoClipboard gCopyClipboardButton, Automatically copy to clipboard

;; Gui, Show, x279 y217 h760 w320

Gui, Add, Edit, x320 y10 w300 r9 vDescriptionText,
Gui, Add, Edit, xp+0 yp+130 w300 r3 vInterpretationText,
Gui, Add, ListView, r5 w300, Location                            |TIRADS|Max size
Gui, Add, Button, w100 h20 gAddToListButton, Add to list (no reset)
Gui, Add, Button, xp+105 yp+0 w100 h20 gRemoveFromListButton, Remove from list
Gui, Add, Button, xp+105 yp+0 w80 h20 gClearListButton, Clear list


;; Gui, Add, Button, x320 yp+20 w50 h20 gButtonOK, OK
Gui, Add, Button, x320 yp+25 w100 h20 gCopyClipboardButton, Copy to clipboard
Gui, Add, Button, xp+105 yp+0 w50 h20  gResetButton, Reset
Gui, Add, Button, xp+55 yp+0 w50 h20 gButtonCancel, Exit

Gui, Show, x279 y217 w630
Description := ""
Tirads_grade := 0
Return 

ButtonCancel:
GuiClose:
ExitApp

AddNoduleButton:
gosub, AddToListButton
gosub, ResetNoduleButton
return

AddNoduleAndCopyButton:
gosub, AddToListButton
gosub, InsertButton
return

AddToListButton:
;; TODO: support voor # nodule nog bij inbouwen
LV_Add("", Side . " " . Location, Tirads_grade, GetMaxSize() . " mm")
NoduleList.Push([Description, Tirads_grade, Side . " " . Location, GetMaxSize()])
return

RemoveFromListButton:
selectedIndex := LV_GetNext(0, "Focused")
LV_Delete(selectedIndex)
NoduleList.RemoveAt(selectedIndex)
return

ClearListButton:
LV_Delete()
NoduleList := []
return

ButtonOK:
clipboard := constructReport(NoduleList)
if WinExist("KWS ahk_exe javaw.exe")
	WinActivate
Send, ^v
ExitApp

InsertButton:
Clipboard := ""
clipboard := constructReport(NoduleList)
ClipWait, 1
if WinExist("KWS ahk_exe javaw.exe")
	WinActivate
Send, ^v
;; TODO: Aanpassen als filename veranderd
Winactivate, TIRADSv2.ahk
return

CopyClipboardButton:
;; TODO: if NoduleList length = 0, add to list
clipboard := constructReport(NoduleList)
Return

ResetNoduleButton:
;; GuiControl, Choose, Side, 1
;; GuiControl, Choose, Location, 1

GuiControl, Choose, Composition, 1
GuiControl, Choose, Echogenicity, 1
GuiControl, Choose, Shape, 1
GuiControl, Choose, Margin, 1
GuiControl,, Foci1, 0
GuiControl,, Foci2, 0
GuiControl,, Foci3, 0
GuiControl,, Foci4, 0
GuiControl,, SizeML,
GuiControl,, SizeAP,
GuiControl,, SizeCC,
Send, +{tab}
GuiControl, Focus, Side
return

ResetButton:
;; Coordinaten van window halen
Reload
;; move window naar oude coordinaten
Return

Check:			; recalculate Tirads_sum and populate description
Gui, Submit, NoHide
Tirads_sum := 0
Tirads_grade := 0
Possibilities := [["grotendeels cystisch", "spongiform", "gemengd cystisch-vastweefsel", "vastweefsel"], ["anechogeen", "hypo- tot isoechogeen", "hypoechogeen", "zeer hypoechogeen"], ["wider-than-tall", "taller-than-wide"], ["glad afgelijnd", "wazig afgelijnd", "gelobuleerd/onregelmatig", "extra-thyroidale extensie"], ["comet tail artefacten (0)", "macrocalcificaties (1)", "perifere rim calcificaties (2)", "punctiforme echogene foci (3)"]]
Description := "- Nodule" . SizeStringMaker() . " in de " . Side . " " . Location . " met kenmerken: `n  . Compositie: " . Composition . "`n  . Echogeniciteit: " . Echogenicity . "`n  . Vorm: " . Shape . "`n  . Aflijning: " . Margin

if (Foci1 or Foci2 or Foci3 or Foci4)
	Description .= "`n  . Foci/calcificaties: "
if (Foci1)
	Description .= Possibilities[5,1] . "; "
if (Foci2)
	Description .= Possibilities[5,2] . "; "
if (Foci3)
	Description .= Possibilities[5,3] . "; "
if (Foci4)
	Description .= Possibilities[5,4]

Tirads_sum := 0
Pos := 1
While Pos := RegExMatch(Description, "\((\d+)", Match, Pos + StrLen(Match)) {
	Tirads_sum += Match1
}

Switch
{
	case Tirads_sum <= 1: Tirads_grade := 1
	case Tirads_sum = 2:  Tirads_grade := 2
	case Tirads_sum = 3:  Tirads_grade := 3
	case Tirads_sum <= 6: Tirads_grade := 4
	case Tirads_sum >= 7: Tirads_grade := 5
}
Description := Description . "`n  => ACR TI-RADS: " . Tirads_grade

GuiControl, Text, %TiradsGradeText%, TI-RADS grade: %Tirads_grade%
GuiControl,,DescriptionText,%Description%
GuiControl,,InterpretationText, % Interpretation(Tirads_grade, GetMaxSize())
If (AutoClipboard)
	gosub, CopyClipboardButton
Return

SizeStringMaker() {
	Gui, Submit, NoHide
	GuiControlGet, SizeML
	GuiControlGet, SizeAP
	GuiControlGet, SizeCC
	substr := ""
	substr2 := ""
	if (SizeML) {
		substr .= SizeML
		substr2 .= "ML"
	}
	if (SizeAP) {
		if (substr) {
			substr .= " x "
			substr2 .= " x "
		}
		substr .= SizeAP
		substr2 .= "AP"
	}
	if (SizeCC) {
		if (substr) {
			substr .= " x "
			substr2 .= " x "
		}
		substr .= SizeCC
		substr2 .= "CC"
	}
	if (substr) {
		return, " van " . substr . " mm (" . substr2 . ")"
	}
	return, ""
}

Interpretation(tiradsgrade, maxsize) {
	if (tiradsgrade = 1) {
		Return, "ACR TI-RADS 1 nodule, benigne."
	} else if (tiradsgrade = 2) {
		Return, "ACR TI-RADS 2 nodule, benigne."
	} else if (tiradsgrade = 3) {
		if (maxsize >= 25) {
			Return, "ACR TI-RADS 3 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return, "ACR TI-RADS 3 nodule van max. " . maxsize . " mm: echografisch te controleren op evolutiviteit over 1, 3, en 5 jaar."
		}
	} else if (tiradsgrade = 4) {
		if (maxsize >= 15) {
			Return, "ACR TI-RADS 4 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return, "ACR TI-RADS 4 nodule van max. " . maxsize . " mm: echografisch te controleren op evolutiviteit over 1, 2, 3, en 5 jaar."
		}
	} else if (tiradsgrade = 5) {
		if (maxsize >= 10) {
			Return, "Voor maligniteit verdachte ACR TI-RADS 5 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return, "Voor maligniteit verdachte ACR TI-RADS 5 nodule van max. " . maxsize . " mm: jaarlijks echografisch te controleren op evolutiviteit gedurende 5 jaar."
		}
	}
	Return, "Error"
}

GetMaxSize(){
	GuiControlGet, SizeML
	GuiControlGet, SizeAP
	GuiControlGet, SizeCC
	return (SizeML>SizeAP ? SizeML:SizeAP) > SizeCC ? (SizeML>SizeAP ? SizeML:SizeAP):SizeCC
}


constructNoduleListDescription(NoduleList) {
	fullDescr := ""
	for iter, nodule in NoduleList {
		fullDescr .= RegExReplace(nodule[1], "- Nodule v?a?n? ?", "- Nodule " . iter . ": ")
		fullDescr .= "`n"
	}
	return fullDescr
}

constructConclusion(NoduleList) {
	fullConclusion := ""
	for iter, nodule in NoduleList {
		fullConclusion .= "- nodule " . iter . ": " . Interpretation(nodule[2], nodule[4]) . "`n"
	}
	return fullConclusion
}

constructReport(NoduleList) {
	return constructNoduleListDescription(NoduleList) . "`n`n===== CONCLUSIE =====`n" . constructConclusion(NoduleList)
}
