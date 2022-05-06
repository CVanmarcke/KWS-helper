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


;                 ====== Start programma =========

Gui, Add, GroupBox, x10 w300 r4, Composition
Gui, Add, Radio, xp+10 yp+20 gCheck vComposition1, Cystic or almost completely cystic (0 points)
Gui, Add, Radio, gCheck vComposition2, Spongiform (0 points)
Gui, Add, Radio, gCheck vComposition3, Mixed cystic and solid (1 point)
Gui, Add, Radio, gCheck vComposition4, Solid or almost completely solid (2 points).
Gui, Add, GroupBox, xp-10 yp+20 w300 r4, Echogenicity
Gui, Add, Radio, xp+10 yp+20 gCheck vEchogenicity1, Anechoic (0 points)
Gui, Add, Radio, gCheck vEchogenicity2, Hyperechoic or isoechoic (1 point)
Gui, Add, Radio, gCheck vEchogenicity3, Hypoechoic (2 points)
Gui, Add, Radio, gCheck vEchogenicity4, Very hypoechoic (3 points)
Gui, Add, GroupBox, xp-10 yp+20 w300 r2, Shape
Gui, Add, Radio, xp+10 yp+20 gCheck vShape1, Wider-than-tall (0 points)
Gui, Add, Radio, gCheck vShape2, Taller-than-wide (3 points)
Gui, Add, GroupBox, xp-10 yp+20 w300 r4, Margin
Gui, Add, Radio, xp+10 yp+20 gCheck vMargin1, Smooth (0 points)
Gui, Add, Radio, gCheck vMargin2, Ill-defined (0 points)
Gui, Add, Radio, gCheck vMargin3, Lobulated or irregular (2 points)
Gui, Add, Radio, gCheck vMargin4, Extra-thyroidal extension (3 points)
Gui, Add, GroupBox, xp-10 yp+20 w300 r4, Echogenic Foci (check all that apply)
Gui, Add, CheckBox, xp+10 yp+20 gCheck vFoci1, None or large comet-tail artifacts (0 points)
Gui, Add, CheckBox, gCheck vFoci2, Macrocalcifications (1 point)
Gui, Add, CheckBox, gCheck vFoci3, Peripheral (rim) calcifications (2 points)
Gui, Add, CheckBox, gCheck vFoci4, Punctate echogenic foci (3 points)
Gui, Add, GroupBox, xp-10 yp+20 w300 h50 r1, Size (ML x AP x CC) in mm
Gui, Add, Edit, xp+10 yp+20 w30 Number vSizeML gCheck,
Gui, Add, Edit, xp+40 yp+0 w30 Number vSizeAP gCheck,
Gui, Add, Edit, xp+40 yp+0 w30 Number vSizeCC gCheck,

Gui, Add, Text, x20 yp+30 vTiradsGradeText HwndTiradsGradeText, TI-RADS grade: 0   
Gui, Add, Edit, xp-10 yp+20 w300 r9 vDescriptionText,
Gui, Add, Edit, xp+0 yp+130 w300 r3 vInterpretationText,
Gui, Add, CheckBox, vAutoClipboard gCopyClipboardButton, Automatically copy to clipboard

Gui, Add, Button, xp+0 yp+20 w120 h20 gCopyClipboardButton, Copy to clipboard
Gui, Add, Button, x147 yp+0 w70 h20 gResetButton, Reset
Gui, Add, Button, x227 yp+0 w70 h20, Cancel
Gui, Show, x279 y217 h760 w320
Description := ""
Tirads_grade := 0
Return 

ButtonCancel:
GuiClose:
ExitApp

CopyClipboardButton:
GuiControlGet, InterpretationText
clipboard := Description . "`n`n===== CONCLUSIE =====`n" . InterpretationText . "`n"
Return

ResetButton:
Reload
Return

Check:			; recalculate Tirads_sum and populate description
Gui, Submit, NoHide
Tirads_sum := 0
Tirads_grade := 0
Possibilities := [["grotendeels cystisch", "spongiform", "gemengd cystisch-vastweefsel", "vastweefsel"], ["anechogeen", "hypo- tot isoechogeen", "hypoechogeen", "zeer hypoechogeen"], ["wider-than-tall", "taller-than-wide"], ["glad afgelijnd", "wazig afgelijnd", "gelobuleerd/onregelmatig", "extra-thyroidale extensie"], ["geen echogene foci en/of comet tail artefacten", "macrocalcificaties", "perifere rim calcificaties", "punctiforme echogene foci"]]
Description := "- nodule " . SizeStringMaker() . "met kenmerken: `n`t. Compositie: "
if (Composition1) {
	Description .= Possibilities[1,1]
} else if (Composition2) {
	Description .= Possibilities[1,2]
} else if (Composition3) {
	Description .= Possibilities[1,3]
	Tirads_sum += 1
} else if (Composition4) {
	Description .= Possibilities[1,4]
	Tirads_sum += 2
} 
Description .= "`n`t. Echogeniciteit: "
if (Echogenicity1) {
	Description .= Possibilities[2,1]
} else if (Echogenicity2) {
	Description .= Possibilities[2,2]
	Tirads_sum += 1
} else if (Echogenicity3) {
	Description .= Possibilities[2,3]
	Tirads_sum += 2
} else if (Echogenicity4) {
	Description .= Possibilities[2,4]
	Tirads_sum += 3
}
Description .= "`n`t. Vorm: "
if (Shape2) {
	Tirads_sum += 3
	Description .= Possibilities[3,2]
} else {
	Description .= Possibilities[3,1]
}
Description .= "`n`t. Aflijning: "
if (Margin1) {
	Description .= Possibilities[4,1]
} else if (Margin2) {
	Description .= Possibilities[4,2]
} else if (Margin3) {
	Description .= Possibilities[4,3]
	Tirads_sum += 2
} else if (Margin4) {
	Description .= Possibilities[4,4]
	Tirads_sum += 3
}
Description .= "`n`t. Foci/calcificaties: "
if (Foci1) {
	Description .= Possibilities[5,1] . "; "
}
if (Foci2) {
	Description .= Possibilities[5,2] . "; "
	Tirads_sum += 1
}
if (Foci3) {
	Description .= Possibilities[5,3] . "; "
	Tirads_sum += 2
}
if (Foci4) {
	Description .= Possibilities[5,4] . "; "
	Tirads_sum += 3
}
if (Tirads_sum <= 1) {
	Tirads_grade := 1
} else if (Tirads_sum = 2) {
	Tirads_grade := 2
} else if (Tirads_sum = 3) {
	Tirads_grade := 3
} else if (Tirads_sum <= 6) {
	Tirads_grade := 4
} else if (Tirads_sum >= 7) {
	Tirads_grade := 5
}
Description := Description . "`n`t=> TI-RADS: " . Tirads_grade

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
		return, "van " . substr . " mm (" . substr2 . ") "
	}
	return, ""
}

Interpretation(tiradsgrade, maxsize) {
	if (tiradsgrade = 1) {
		Return, "TI-RADS 1 nodule, benigne."
	} else if (tiradsgrade = 2) {
		Return, "TI-RADS 2 nodule, benigne."
	} else if (tiradsgrade = 3) {
		if (maxsize >= 25) {
			Return, "TI-RADS 3 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return, "TI-RADS 3 nodule van max. " . maxsize . " mm: echografisch te controleren op evolutiviteit over 1, 3, en 5 jaar."
		}
	} else if (tiradsgrade = 4) {
		if (maxsize >= 15) {
			Return, "TI-RADS 4 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return, "TI-RADS 4 nodule van max. " . maxsize . " mm: echografisch te controleren op evolutiviteit over 1, 2, 3, en 5 jaar."
		}
	} else if (tiradsgrade = 5) {
		if (maxsize >= 10) {
			Return, "Voor maligniteit verdachte TI-RADS 5 nodule van max. " . maxsize . " mm: aan te vullen met FNAC ter histologische karakterisatie."
		} else {
			Return, "Voor maligniteit verdachte TI-RADS 5 nodule van max. " . maxsize . " mm: jaarlijks echografisch te controleren op evolutiviteit gedurende 5 jaar."
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


