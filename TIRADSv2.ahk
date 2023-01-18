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

global Side, Location, Tirads_grade, Shape, Echogenicity, Composition, Location, Margin
NoduleList := []
Description := ""
Tirads_grade := 0

myGui := Gui()
myGui.OnEvent("Escape", ButtonCancel.Bind("Close"))
myGui.OnEvent("Close", ButtonCancel.Bind("Close"))
myGui.Add("GroupBox", "x10 w300 R2", "Location")
ogcDropDownListSide := myGui.Add("DropDownList", "xp+10 yp+20 w70 vSide", ["linker", "rechter", "isthmus"])
ogcDropDownListSide.OnEvent("Change", Check.Bind("Change"))
ogcDropDownListLocation := myGui.Add("DropDownList", "xp+80 w80 vLocation", ["bovenpool", "middenpool", "onderpool"])
ogcDropDownListLocation.OnEvent("Change", Check.Bind("Change"))

myGui.Add("GroupBox", "x10 w300 R6", "Characteristics")
ogcDropDownListComposition := myGui.Add("DropDownList", "xp+10 yp+20 w250 vComposition  Choose1", ["cystisch (0 punten)", "spongiform (0 punten)", "gemengd cystisch en solide (1 punt)", "solide of bijna volledig solid (2 punten)"])
ogcDropDownListComposition.OnEvent("Change", Check.Bind("Change"))
ogcDropDownListEchogenicity := myGui.Add("DropDownList", "w100 w250 vEchogenicity  Choose1", ["anechogeen (0 punten)", "hyperechogeen of isoechogeen (1 punt)", "hypoechogeen (2 punten)", "zeer hypoechogeen (3 punten)"])
ogcDropDownListEchogenicity.OnEvent("Change", Check.Bind("Change"))
ogcDropDownListShape := myGui.Add("DropDownList", "w100 w250 vShape  Choose1", ["wider-than-tall (0 punten)", "taller-than-wide (3 punten)"])
ogcDropDownListShape.OnEvent("Change", Check.Bind("Change"))
ogcDropDownListMargin := myGui.Add("DropDownList", "w100 w250 vMargin  Choose1", ["scherp en glad (0 punten)", "wazig afgelijnd (0 punten)", "gelobuleerd of onregelmatig (2 punten)", "extra-thyroidale extensie (3 punten)"])
ogcDropDownListMargin.OnEvent("Change", Check.Bind("Change"))

myGui.Add("GroupBox", "x10 w300 r4", "Echogenic Foci (check all that apply)")
;; TODO: label met dan dropdown nee (default) of ja
ogcCheckBoxFoci1 := myGui.Add("CheckBox", "xp+10 yp+20  vFoci1", "Comet-tail artifacten (0 punten)")
ogcCheckBoxFoci1.OnEvent("Click", Check.Bind("Normal"))
ogcCheckBoxFoci2 := myGui.Add("CheckBox", "vFoci2", "Macrocalcificaties (1 punt)")
ogcCheckBoxFoci2.OnEvent("Click", Check.Bind("Normal"))
ogcCheckBoxFoci3 := myGui.Add("CheckBox", "vFoci3", "Perifere (rim) verkalkingen (2 punten)")
ogcCheckBoxFoci3.OnEvent("Click", Check.Bind("Normal"))
ogcCheckBoxFoci4 := myGui.Add("CheckBox", "vFoci4", "Punctiforme echogene foci (3 punten)")
ogcCheckBoxFoci4.OnEvent("Click", Check.Bind("Normal"))

myGui.Add("GroupBox", "xp-10 yp+20 w300 h50 r1", "Size (ML x AP x CC) in mm")
ogcEditSizeML := myGui.Add("Edit", "xp+10 yp+20 w30 Number vSizeML")
ogcEditSizeML.OnEvent("Change", Check.Bind("Change"))
ogcEditSizeAP := myGui.Add("Edit", "xp+40 yp+0 w30 Number vSizeAP")
ogcEditSizeAP.OnEvent("Change", Check.Bind("Change"))
ogcEditSizeCC := myGui.Add("Edit", "xp+40 yp+0 w30 Number vSizeCC")
ogcEditSizeCC.OnEvent("Change", Check.Bind("Change"))

ogcTiradsGradeText := myGui.Add("Text", "x20 yp+30 vTiradsGradeText", "TI-RADS grade: 0")
TiradsGradeText := ogcTiradsGradeText.hwnd

;; TODO: knop: add new nodule, wordt dan aan lijst toegevoegd en reset.
ogcButtonAddanothernodule := myGui.Add("Button", "x10 w145 h40", "Add another nodule")
ogcButtonAddanothernodule.OnEvent("Click", AddNoduleButton.Bind("Normal"))
ogcButtonAddnoduleandcopytoKWS := myGui.Add("Button", "xp+155 yp+0 w145 h40", "Add nodule and copy to KWS")
ogcButtonAddnoduleandcopytoKWS.OnEvent("Click", AddNoduleAndCopyButton.Bind("Normal"))
ogcButtonCopytoKWS := myGui.Add("Button", "xp+155 yp+0 w145 h40", "Copy to KWS!")
ogcButtonCopytoKWS.OnEvent("Click", InsertButton.Bind("Normal"))

ogcEditDescriptionText := myGui.Add("Edit", "x320 y10 w300 r9 vDescriptionText")
ogcEditInterpretationText := myGui.Add("Edit", "xp+0 yp+130 w300 r3 vInterpretationText")
ogcListViewLocationTIRADSMaxsize := myGui.Add("ListView", "r5 w300", ["Location              ", "TIRADS", "Max size"])
ogcButtonAddtolistnoreset := myGui.Add("Button", "w100 h20", "Add to list (no reset)")
ogcButtonAddtolistnoreset.OnEvent("Click", AddToListButton.Bind("Normal"))
ogcButtonRemovefromlist := myGui.Add("Button", "xp+105 yp+0 w100 h20", "Remove from list")
ogcButtonRemovefromlist.OnEvent("Click", RemoveFromListButton.Bind("Normal"))
ogcButtonClearlist := myGui.Add("Button", "xp+105 yp+0 w80 h20", "Clear list")
ogcButtonClearlist.OnEvent("Click", ClearListButton.Bind("Normal"))

ogcButtonCopytoA_Clipboard := myGui.Add("Button", "x320 yp+25 w100 h20", "Copy to Clipboard")
ogcButtonCopytoA_Clipboard.OnEvent("Click", CopyClipboardButton.Bind("Normal"))
ogcButtonReset := myGui.Add("Button", "xp+105 yp+0 w50 h20", "Reset")
ogcButtonReset.OnEvent("Click", ResetButton.Bind("Normal"))
ogcButtonExit := myGui.Add("Button", "xp+55 yp+0 w50 h20", "Exit")
ogcButtonExit.OnEvent("Click", ButtonCancel.Bind("Normal"))

myGui.Show("x279 y217 w630")

ButtonCancel(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	ExitApp()
}

AddNoduleButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	AddToListButton(A_GuiEvent, GuiCtrlObj, Info)
	ResetNoduleButton()
	return
}

AddNoduleAndCopyButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	AddToListButton(A_GuiEvent, GuiCtrlObj, Info)
	InsertButton(A_GuiEvent, GuiCtrlObj, Info)
	return
}

AddToListButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	ogcListViewLocationTIRADSMaxsize.Add("", Side . " " . Location, Tirads_grade, GetMaxSize() . " mm")
	NoduleList.Push([ogcEditDescriptionText.Text, Tirads_grade, ogcDropDownListSide.Text . " " . Location, GetMaxSize()])
	return
}

RemoveFromListButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	selectedIndex := ogcListViewLocationTIRADSMaxsize.GetNext(0,"Focused")
	ogcListViewLocationTIRADSMaxsize.Delete(selectedIndex)
	NoduleList.RemoveAt(selectedIndex)
	return
}

ClearListButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	ogcListViewLocationTIRADSMaxsize.Delete()
	NoduleList := []
	return
}


ButtonOK(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	A_Clipboard := constructReport(NoduleList)
	if WinExist("KWS ahk_exe javaw.exe")
		WinActivate()
	Send("^v")
	ExitApp()
}

InsertButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	A_Clipboard := ""
	A_Clipboard := constructReport(NoduleList)
	Errorlevel := !ClipWait(1)
	if WinExist("KWS ahk_exe javaw.exe")
		WinActivate()
	Send("^v")
	;; TODO: Aanpassen als filename veranderd
	WinActivate("TIRADSv2.ahk")
	return
}

CopyClipboardButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	;; TODO: if NoduleList length = 0, add to list
	A_Clipboard := constructReport(NoduleList)
	Return
}

ResetNoduleButton() {
	ogcDropDownListComposition.Choose(1)
	ogcDropDownListEchogenicity.Choose(1)
	ogcDropDownListShape.Choose(1)
	ogcDropDownListMargin.Choose(1)
	ogcCheckBoxFoci1.Value := 0
	ogcCheckBoxFoci2.Value := 0
	ogcCheckBoxFoci3.Value := 0
	ogcCheckBoxFoci4.Value := 0
	ogcEditSizeML.Value := ""
	ogcEditSizeAP.Value := ""
	ogcEditSizeCC.Value := ""
	Send("+{tab}")
	ogcDropDownListSide.Focus()
	return
}

ResetButton(A_GuiEvent, GuiCtrlObj, Info := "", *) {
;; Coordinaten van window halen
	Reload()
	Return
}

Check(A_GuiEvent, GuiCtrlObj, Info := "", *)	{		; recalculate Tirads_sum and populate description
	global Side, Location, Tirads_grade, Shape, Echogenicity, Composition, Location, Margin
	oSaved := myGui.Submit("0")
	Side := oSaved.Side
	Location := oSaved.Location
	Composition := oSaved.Composition
	Echogenicity := oSaved.Echogenicity
	Shape := oSaved.Shape
	Margin := oSaved.Margin
	Foci1 := oSaved.Foci1
	Foci2 := oSaved.Foci2
	Foci3 := oSaved.Foci3
	Foci4 := oSaved.Foci4
	SizeML := oSaved.SizeML
	SizeAP := oSaved.SizeAP
	SizeCC := oSaved.SizeCC
	DescriptionText := oSaved.DescriptionText
	InterpretationText := oSaved.InterpretationText
	Tirads_sum := 0
	Tirads_grade := 0
	Possibilities := [["grotendeels cystisch", "spongiform", "gemengd cystisch-vastweefsel", "vastweefsel"], ["anechogeen", "hypo- tot isoechogeen", "hypoechogeen", "zeer hypoechogeen"], ["wider-than-tall", "taller-than-wide"], ["glad afgelijnd", "wazig afgelijnd", "gelobuleerd/onregelmatig", "extra-thyroidale extensie"], ["comet tail artefacten (0)", "macrocalcificaties (1)", "perifere rim calcificaties (2)", "punctiforme echogene foci (3)"]]
	Description := "- Nodule" . SizeStringMaker() . " in de " . Side . " " . Location . ", ACR TI-RADS. `n  . Compositie: " . Composition . "`n  . Echogeniciteit: " . Echogenicity . "`n  . Vorm: " . Shape . "`n  . Aflijning: " . Margin

	if (Foci1 or Foci2 or Foci3 or Foci4)
		Description .= "`n  . Foci/calcificaties: "
	if (Foci1)
		Description .= Possibilities[5][1] . "; "
	if (Foci2)
		Description .= Possibilities[5][2] . "; "
	if (Foci3)
		Description .= Possibilities[5][3] . "; "
	if (Foci4)
		Description .= Possibilities[5][4]

	Tirads_sum := 0
	Pos := 1
	While Pos := RegExMatch(Description, "\((\d+)", &Match, Pos + 3) {
		Tirads_sum += Match[1]
	}

	Switch
	{
		case Tirads_sum <= 1: Tirads_grade := 1
		case Tirads_sum = 2:  Tirads_grade := 2
		case Tirads_sum = 3:  Tirads_grade := 3
		case Tirads_sum <= 6: Tirads_grade := 4
		case Tirads_sum >= 7: Tirads_grade := 5
	}
	; Description := Description . "`n  => ACR TI-RADS: " . Tirads_grade
	Description := StrReplace(Description, "ACR TI-RADS", "ACR TI-RADS " . Tirads_grade)
	ogcTiradsGradeText.Text := "TI-RADS grade: " Tirads_grade
	ogcEditDescriptionText.Value := Description
	ogcEditInterpretationText.Value := Interpretation(Tirads_grade, GetMaxSize())
	Return
}

SizeStringMaker() {
	SizeML := ogcEditSizeML.Text
	SizeAP := ogcEditSizeAP.Text
	SizeCC := ogcEditSizeCC.Text
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
		return " van " . substr . " mm (" . substr2 . ")"
	}
	return ""
}

Interpretation(tiradsgrade, maxsize) {
	if (tiradsgrade = 1) {
		Return "ACR TI-RADS 1 nodule, benigne."
	} else if (tiradsgrade = 2) {
		Return "ACR TI-RADS 2 nodule, benigne."
	} else if (tiradsgrade = 3) {
		if (maxsize >= 25) {
			Return "ACR TI-RADS 3 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return "ACR TI-RADS 3 nodule van max. " . maxsize . " mm: echografisch te controleren op evolutiviteit over 1, 3, en 5 jaar."
		}
	} else if (tiradsgrade = 4) {
		if (maxsize >= 15) {
			Return "ACR TI-RADS 4 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return "ACR TI-RADS 4 nodule van max. " . maxsize . " mm: echografisch te controleren op evolutiviteit over 1, 2, 3, en 5 jaar."
		}
	} else if (tiradsgrade = 5) {
		if (maxsize >= 10) {
			Return "Voor maligniteit verdachte ACR TI-RADS 5 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return "Voor maligniteit verdachte ACR TI-RADS 5 nodule van max. " . maxsize . " mm: jaarlijks echografisch te controleren op evolutiviteit gedurende 5 jaar."
		}
	}
	Return "Error"
}

GetMaxSize(){
	SizeML := toInteger(ogcEditSizeML.Text)
	SizeAP := toInteger(ogcEditSizeAP.Text)
	SizeCC := toInteger(ogcEditSizeCC.Text)
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
	return constructNoduleListDescription(NoduleList) . "`n`nCONCLUSIE:`n" . constructConclusion(NoduleList)
}


toInteger(int) {
	return Integer(IsNumber(int) ? int : 0)
}

