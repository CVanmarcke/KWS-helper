﻿;; GECONVERTEERD VAN v1 naar v2: kan nog opgekuisd worden (TODO)
#SingleInstance Force
#Include "Sift.ahk" ; Sift library

SetTitleMatchMode(1)

  Data := ""
  Data_orl := "
  (
    Rotsbeenderen standaard				RAD mr rotsb (+)
    Verworven perceptiedoofheid/perifere vertigo		RAD mr orl 01 (+)
    Congenitale perceptiedoofheid			RAD mr orl 02 (-)
    Cholesteatoma/ontsteking				RAD mr orl 08 (-)
    Uitwendig Oor/Middenoor Tumor			RAD mr 10 (+)
    Uitwendig Oor/Middenoor Ontsteking			RAD mr orl 32 (+)
    ORL hydrops					RAD mr orl 24 (+)
    Pulsatiele tinnitus					RAD mr orl 06 (+)
    Niet-pulsatiele tinnitus				RAD mr orl 05 (+)
    Centrale vertigo/vertebrobasilaire insufficientie		RAD mr orl 04 (-)
    Aangezichtszenuwen				RAD mr orl 36 (+)
    Maxillofaciaal Massief tumor MFM			RAD mr orl 14 (+)
    Tong / snurken					RAD mr orl 16 (-)
    Tong/Tumor					RAD mr orl 22 (+)
    Reukverlies/anosmie/parosmie			RAD mr orl 15 (?)
    Hemifaciaal spasme/tics				RAD mr orl 07 (+)
    Trigeminusneuralgie Centraal				RAD mr orl 33 (+)
    Trigeminusneuralgie Perifeer				RAD mr orl 31 (+)
    Kaakgewricht/ATM					RAD mr orl 13 (-)
    Parotispathologie					RAD mr orl 19 (+)
    Hals tumor					RAD mr orl 17 (+)
    Schildklier					RAD mr 34 (+)
    Bijschildklier					RAD mr 35 (+)
    Plexus Brachialis standaard				RAD mr plex brach (?)
    Plexus Brachialis trauma				RAD mr orl 21 (-)
    Plexus Brachialis tumor				RAD mr orl 20 (+)
    zondercontrast     (-)
    metcontrast    (+)
  )"
  Data_neuro := "
  (
  Cervicobrachalgie				RAD mr wz 01  (-)
  CWK/DWK Post-Op			RAD mr wk 19  (-)
  Congenitale Aandoening LWK			RAD mr wk 21  (-)
  Full Spine					RAD mr full spine (-)
  Intraspinale Metastasen			RAD mr wk 23  (+)
  LWK Post-op				RAD mr wk 15  (+)
  LWK/DWK Standaard			RAD mr wk 18  (-)
  Metaal					RAD mr wk 22  (-)
  Ruggenmerg (MS of Myelitis)			RAD mr wz 13  (-)
  Ruggenmergletsel (geen full spine)		RAD mr wk 16  (+)
  Scoliose					RAD mr wk 24  (-)
  Spinale AV-Fistel(enkel bij controle)		RAD mr wk 25  (+)
  Liquorlek					RAD mr wk 26  (?)
  Aneurysma				RAD amr hersen 57 (-)
  Arterio-Veneuze Malformatie			RAD mr hersen 71  (+)
  Bloeding					RAD mr hersen 65  (+)
  Caverneuze Malformatie/Fabry		RAD mr hersen 67  (-)
  Cerebellair Letsel (tumor)			RAD mr hersen 13  (+)
  CVA-TIA					RAD mr hersen 04  (-)
  Dementie					RAD mr hersen 11  (-)
  Epilepsie, Hippocampus			RAD mr hersen 06  (-)
  Frameless stereotaxie (neurochirurgie)		RAD mr hersen 54  (+)
  Hoofdpijn					RAD mr hersen 24  (-)
  Hydrocefalie / NPH				RAD mr hersen 64  (-)
  Hypotensie				RAD mr hersen 72  (?)
  Idiopathische Intracraniale Hypertensie		RAD mr hersen 73  (-)
  Liquorlek					RAD mr hersen 68  (-)
  MS (Vermoeden MS)			RAD mr hersen 02  (?)
  MS Follow-up				RAD mr hersen 03  (-)
  Neurofibromatosis/Fakomatose		RAD mr hersen 74  (+)
  Neurovasculair Confict V/VII/IX		RAD mr hersen 75  (+)
  Parkinson					RAD mr hersen 57  (-)
  Post-operatieve Controle			RAD mr hersen 27  (+)
  Routineprotocol				RAD mr hersen 25  (-)
  Sella Macro-adenoma			RAD mr hersen 08  (+)
  Sella Micro-adenoma			RAD mr hersen 09  (+)
  Sella Post-operatief				RAD mr hersen 10  (+)
  Stereotactisch Radiotherapie			RAD mr hersen 89  (+)
  Tumor met Perfusie				RAD mr hersen 77  (+)
  Tumor/infectie zonder Perfusie		RAD mr hersen 78  (+)
  Vasculitis					RAD amr hersen 58 (+)
  Veneuze Trombose				RAD mr hersen 14  (-)
  Vertigo/Duizeligheid				RAD mr hersen 29  (-)
  Vestibulair Schwannoom			RAD mr hersen 70  (+)
  Trauma					RAD mr hersen 32  (-)
  Epilepsie, Hippocampus Pathologie (pediatrie)	RAD mr hersen 46  (-)
  Langerhanscelhistiocyste			RAD mr hersen 81  (+)
  Neonato					RAD mr hersen 84  (-)
  Neurofibromatose (pediatrie)			RAD mr hersen 85  (+)
  PVL					RAD mr hersen 23  (-)
  Routineprotocol (pediatrie)			RAD mr hersen 87  (-)
  Tumor (Pediatrie)				RAD mr hersen 88  (+)
  MR Orbita				RAD mr orbita (+)
  MR Halsvaten Dissectie			RAD amr hersen 59 (+)
  MR Hersenen (Pediatrie) Pubertas Praecox	RAD mr hersen 42 (-)
  Studie (?)
  )"

  Data_abdomen_MR := "
  (
	TODO: onvolledig!!

	Abdomen Peritoneaalmetastasen			RAD mr abd 41 (+)
	Pancreas Tumor					RAD mr abd 55 (+)
	Chronische Pancreatitis				RAD mr abd 34 (-)
	Pancreas Familiaal Carcinoma				RAD mr abd 33 (-)
	Pancreas voor Pancreatitis				RAD mr abd 08 (?)
	Lever met Contrast, zonder Laattijdige			RAD mr abd 48 (+)
	Lever met Contrast, met Laattijdige			RAD mr abd 49 (+)
	Lever zonder contrast				RAD mr abd 50 (-)
	Lever hemochromatose				RAD mr abd 35 (?)
	Levertransplantatie					RAD mr abd 03 (+)
	Cholangio zonder contrast				RAD mr abd 52 (-)
	Cholangio met contrast (PSC, IGG4, maligne stenose)	RAD mr abd 53 (+)
	Cholangio gallek					RAD mr abd 52 (?)
	Staging Recto-Anale Tumor				RAD mr abd 29 (+)
	Ovaria IOTA					RAD mr abd 44 (+)
	Entero						RAD mr abd 37

  )"

  Data_abdomen_CT := "
  (
  Klassiek abdomen			RAD ct abd 22 (+) [IV veneus + 3 PO]
  Bloeding				RAD ct abd 22 (+) [trifasisch]
  Buikwand				RAD ct abd 12 (-)
  Lever/pancreas			RAD ct abd 22 (+) [2 PO + a blanc en arterieel bovenbuik + veneus volledig abdomen]
  Bovenbuik pancreas		RAD ct abd 26 (+) [bovenbuik arterieel en veneus + 2 PO]
  Bovenbuik lever			RAD ct abd 25 (+) [bovenbuik en veneus + 2 PO]
  Combi abdomenlijst			RAD ct comb 02 (+) [IV veneus + 3 PO]
  Combi +hersenen			RAD ct comb 03 (+) [IV veneus + 3 PO]
  Combi URO lijst			RAD ct uro 24 (+)
  Combi thoraxlijst			RAD ct thorax 20 (+)

  )"

 Data_thorax := "
  (
  Longembolen				RAD ct thorax 28 (+)
  Longembolen (chronisch)			RAD ct thorax 35 (+)
  Klassiek					RAD ct thorax 24 (-)
  HRCT / ILD				RAD ct thorax 23 (-)
  Transplant				RAD ct thorax 33 (-)
  Low Dose				RAD ct thorax 34 (-)
  Combi thorax/bovenbuik/hersenen		RAD ct comb 06 (+)
  Combi thorax/bovenbuik			RAD ct thorax 20 (+)
  Mediastinum				RAD ct thorax 15 (+)
  )"

  Data_uro_CT  := "
  (
  Bijnier					RAD ct uro 08 (+)
  Nier vasculair				RAD ct uro 100 (+)
  Niertumor primair				RAD ct uro 19 (+)
  Pyelonefritis				RAD ct uro 10 (+)
  Low dose lithiasis				RAD ct uro 14 (-)
  IVU					RAD ct uro 20 (+)
  Follow up/screening abdomen met aflopen	RAD ct uro 30 (+)
  Combi thorax/abdomen			RAD ct uro 24 (+)
  Combi + hersenen				RAD ct uro 27 (+)
  Combi thorax/abdomen + afloop RAD ct uro 29 (+)
  Trauma					RAD ct uro 22 (+)
  Nierdonor					RAD ct uro 21 (+)
  Niervolumetrie				RAD ct uro 28 (-) [Indien ADPKD normale dosis]
  )"

  helptext := "
  (
  Enter		-> Vul het eerste onderzoek in in KWS.
  Ctrl-Enter 	-> Aanvaard het onderzoek in KWS (duwt op OK).
  Ctrl-NumpadEnter -> Aanvaard het onderzoek in KWS (duwt op OK).
  Ctrl-Numpad+ 	-> Selecteer 'met contrast'.
  Ctrl-Numpad- 	-> Selecteer 'zonder contrast'.
  Ctrl-i		-> Zet de cursor in 'Opmerkingen'.
  Ctrl-q		-> Sla de huidige patient over.

  NB: Als abdomen CT is geselecteerd als discipline, zal Ctrl-Numpad+ ook 'IV veneus {+} 3 PO invullen'.
)"

;;::startaanv::
  ;; Goto("start_aanvaardingen")

start_aanvaardingen:
  HotIfWinActive("Aanvaardingen helper ahk_class AutoHotkeyGUI")
  Hotkey("^Enter", druk_ok_aanvaarding)
  Hotkey("^NumpadEnter", druk_ok_aanvaarding)
  ;; Hotkey, Enter, aanvaard_onderzoek_knop ;; mss neit meer nodig met default
  Hotkey("^NumpadAdd", ctrnumplusHotkey)
  Hotkey("^NumpadSub", ctrlnumminHotkey)
  Hotkey("^i", selectOpmerking)
  Hotkey("^q", sluitaanvaardschermKWS)
  toon_onderzoeken := ""
  ;; TODO onthouden welke subdiscipline.
  global aanvaarder
  aanvaarder := Gui()
  aanvaarder.OnEvent("Close", aanvaarderGuiEscape)
  aanvaarder.OnEvent("Escape", aanvaarderGuiEscape)
  ogcEditonderzoek_naam := aanvaarder.Add("Edit", "x10 y9 w280 h20 vonderzoek_naam")
  ogcEditonderzoek_naam.OnEvent("Change", zoek_onderzoek_naam.Bind("Change"))
  onderzoek_naam := ogcEditonderzoek_naam.hwnd
  ogcButtonAanvaard := aanvaarder.Add("Button", "x295 y9 w180 h20  Default", "Aanvaard")
  ogcButtonAanvaard.OnEvent("Click", aanvaard_onderzoek_knop.Bind("Normal"))
  ogcDropDownListsubdiscipline := aanvaarder.Add("DropDownList", "x480 y9 w110 h20 R10 vsubdiscipline Choose1", ["neuro", "ORL", "abdomen (CT)", "abdomen (MR)", "thorax", "uro (CT)"])
  ogcDropDownListsubdiscipline.OnEvent("Change", set_subdiscipline.Bind("Change"))
  ogcButton := aanvaarder.Add("Button", "x595 y9 w15 h20", "?")
  ogcButton.OnEvent("Click", helpknop.Bind("Normal"))
  ogcEditGui_Display := aanvaarder.Add("Edit", "x10 y39 w600 h190 vGui_Display ReadOnly -wrap", toon_onderzoeken)
  aanvaarder.Title := "Aanvaardingen helper"
  aanvaarder.Show("x420 y781 h241 w620")
  WinSetAlwaysOnTop(1, "Aanvaardingen helper")
  set_subdiscipline("", aanvaarder)
  Return

set_subdiscipline(A_GuiEvent, GuiCtrlObj, Info := "", *)
{ ; V1toV2: Added bracket
	Global Data
	oSaved := aanvaarder.Submit("0")
	onderzoek_naam := oSaved.onderzoek_naam
	subdiscipline := oSaved.subdiscipline
	Gui_Display := oSaved.Gui_Display
	Switch subdiscipline
	{
		Case "neuro": Data := Data_neuro
		Case "ORL": Data := Data_orl
		Case "abdomen (CT)": Data := Data_abdomen_CT
		Case "abdomen (MR)": Data := Data_abdomen_MR
		Case "thorax": Data := Data_thorax
		Case "uro (CT)": Data := Data_uro_CT
		Default: Data := Data_neuro
	}
	zoek_onderzoek_naam(A_GuiEvent, GuiCtrlObj, Info)
} ; V1toV2: Added bracket before function

zoek_onderzoek_naam(A_GuiEvent, GuiCtrlObj, Info := "", *)
{ ; V1toV2: Added bracket
	Global Data, subdiscipline
	oSaved := aanvaarder.Submit("0")
	onderzoek_naam := oSaved.onderzoek_naam
	subdiscipline := oSaved.subdiscipline
	Gui_Display := oSaved.Gui_Display
	toon_onderzoek := Sift_Regex(&Data, &onderzoek_naam, "oc")
	ogcEditGui_Display.Value := toon_onderzoek
} ; V1toV2: Added Bracket before label

aanvaard_onderzoek_knop(A_GuiEvent, GuiCtrlObj, Info := "", *)
{ ; V1toV2: Added bracket
	global subdiscipline, Data
	oSaved := aanvaarder.Submit("0")
	onderzoek_naam := oSaved.onderzoek_naam
	subdiscipline := oSaved.subdiscipline
	Gui_Display := oSaved.Gui_Display
	onderzoek := Sift_Regex(&Data, &onderzoek_naam, "oc")
	onderzoek := StrSplit(onderzoek, "`n")[1]
	RegExMatch(onderzoek, "(RAD .*[a-zA-Z0-9])\ +\((.)\)(?:\ +\[(.+)\])?", &gekozenOnderzoek)
	aanvaardOnderzoek(gekozenOnderzoek.1, gekozenOnderzoek.2, gekozenOnderzoek.3)
} ; V1toV2: Added bracket before function

druk_ok_aanvaarding(ThisHotkey)
{ ; V1toV2: Added bracket
  WinActivate("KWS ahk_exe javaw.exe")
  MouseGetPos(&mouseX, &mouseY)
  ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\okButton.png")
  if (ErrorLevel = 2 or ErrorLevel = 1) {
	  MsgBox("Er is iets fout gegaan met zoeken naar de OK knop (niet gevonden of afbeelding bestaat niet)")
	  return
  }
  MouseClick("left", FoundX+5, FoundY+5)
  MouseMove(mouseX, mouseY)
  ogcEditonderzoek_naam.Text := "" ;;nodig?
  WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
  zoek_onderzoek_naam("Button", "")
  return
} ; V1toV2: Added Bracket before label

aanvaarderGuiEscape(*)
{ ; V1toV2: Added bracket
aanvaarderGuiClose:
	ExitApp()
Return
} ; V1toV2: Added bracket before function

selectOpmerking(ThisHotkey)
{ ; V1toV2: Added bracket
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\contrastLabel.png")
	If (ErrorLevel = 0) {
		MouseClick("Left", FoundX + 250, FoundY + 200)
		SendInput("^a")
		MouseMove(mouseX, mouseY)
	} else
		MsgBox("Het opmerkingen formulier werd niet gevonden.")
return
} ; V1toV2: Added Bracket before label

sluitaanvaardschermKWS(ThisHotkey)
{ ; V1toV2: Added bracket
  WinActivate("KWS ahk_exe javaw.exe")
  Send("^{F4}")
  WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
return
} ; V1toV2: Added bracket before function

ctrnumplusHotkey(ThisHotkey)
{ ; V1toV2: Added bracket
	global subdiscipline
	Switch subdiscipline {
		Case "abdomen (CT)": aanvaardOnderzoek("", "+", "IV veneus {+} 3 PO")
		Default: aanvaardOnderzoek("", "+", "")
	}
return
} ; V1toV2: Added Bracket before label

ctrlnumminHotkey(ThisHotkey)
{ ; V1toV2: Added bracket
  aanvaardOnderzoek("", "-", "")
return
} ; V1toV2: Added Bracket before label

helpknop(A_GuiEvent, GuiCtrlObj, Info := "", *)
{ ; V1toV2: Added bracket
MsgBox(helptext)
return
} ; V1toV2: Added bracket before function

aanvaardOnderzoek(onderzoekCode := "", contrast := "?", opmerking := "") {
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	;; MsgBox, % onderzoekCode " " contrast " " opmerking
	if (onderzoekCode != "") {
		ErrorLevel := !ImageSearch(&LocaX, &LocaY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\onderzoekLabel.png") ; Zoek de lijn onderzoek
		if (ErrorLevel > 1) ; Error.
			MsgBox("Error")
		else if (ErrorLevel = 1)
			MsgBox("the label was not found on the screen (error level 1)")
		else {
			MouseClick("Left", LocaX + 100, LocaY + 5)
			Sleep(50)
			SendInput("^a")
			SendInput("{backspace}")
			opmerking := StrReplace(opmerking, "+", "{+}")
			SendInput(onderzoekCode) ; typ het onderzoek in de regel van aanvaarding.
			Sleep(200)
		}
	}
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\contrastLabel.png")
	If (ErrorLevel = 0) {
		Switch contrast
		{
			Case "-": contrastX := FoundX + 80 ;; Zonder
			Case "+": contrastX := FoundX + 165 ;; Met
			Case "?": contrastX := FoundX + 250 ;; zonder/met
			Default: contrastX := FoundX + 250 ;; Zonder/Met
		}
		MouseClick("left", contrastX, FoundY)
		if (opmerking != "") {
			LabelFieldX := FoundX + 250
			LabelFieldY := FoundY + 200
			MouseClick("Left", LabelFieldX, LabelFieldY)
			Sleep(50)
			SendInput("^a")
			SendInput("{backspace}")
			SendInput(opmerking)
		}
		MouseMove(mouseX, mouseY)
	} else {
		MsgBox("The contrast label was not found.")
	}
	sleep(150)
	WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}