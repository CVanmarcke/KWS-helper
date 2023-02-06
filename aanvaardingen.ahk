﻿;; GECONVERTEERD VAN v1 naar v2: kan nog opgekuisd worden (TODO)
#SingleInstance Force
#Include "Sift.ahk" ; Sift library

SetTitleMatchMode(1)
SetMouseDelay(-1) ;remove delays from mouse actions
SetDefaultMouseSpeed 0

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

Data_abdomen_CT_oud := "
  (
  Klassiek abdomen			RAD ct abd 22 (+) [IV veneus + 3 PO]
  Bloeding				RAD ct abd 22 (+) [trifasisch]
  Buikwand				RAD ct abd 12 (-)
  Lever/pancreas			RAD ct abd 22 (+) [2 PO + a blanc en arterieel bovenbuik + veneus volledig abdomen]
  Bovenbuik pancreas		RAD ct abd 26 (+) [bovenbuik arterieel en veneus + 2 PO]
  Bovenbuik lever			RAD ct abd 25 (+) [bovenbuik arterieel en veneus + 2 PO]
  Combi abdomenlijst			RAD ct comb 02 (+) [IV veneus + 3 PO]
  Combi +hersenen			RAD ct comb 03 (+) [IV veneus + 3 PO]
  Combi URO lijst			RAD ct uro 24 (+)
  Combi thoraxlijst			RAD ct thorax 20 (+)

  NOTA: gebruik ctrl-m en ctrl-z om snel met of zonder contrast te kiezen. ctrl-m zet ook automatisch IV veneus + 3 PO.
  )"

Data_abdomen_CT := "
  (
  Klassiek abdomen				RAD ct abd 22 (+) [IV veneus + 3 PO]
  Ischemie (darm/parenchym)			RAD ct abd 22 (+) [abd art, abd veneus op 75s, 3-4 ml/s, geen PO]
  Obstructie zonder ischemie			RAD ct abd 22 (+) [IV portaal veneus, flow rate 3-4 ml/s, scan 75s na injectie, geen PO]
  Interne herniatie post GABY			RAD ct abd 22 (+) [abd veneus, 1 beker of zoveel mogelijk 10 min voor scan, halve beker op tafel]
  Bloeding					RAD ct abd 22 (+) [abd trifasisch angio/aorta protocol, 3-4 ml/s, geen PO (bellen afdeling indien hos)]
  GI lek distale slokdarm/maag			RAD ct comb 02 (+) [één scanrange thorax-abdomen, 1 beker of zoveel mogelijk 10 min voor scan, 1 beker op tafel]
  GI lek maag-duod-dundarm			RAD ct abd 22 (+) [abd veneus, 3 PO]
  GI lek distale dundarm-rechter colon		RAD ct abd 22 (+) [abd veneus, 4 PO over 2u + retro indien mogelijk]
  GI lek rectosigmoid - linker colon		RAD ct abd 22 (+) [abd veneus, + retro, geen PO]
  Enterografie (Crohn,..)			RAD ct abd 28 (+) [bespreken met supervisie. Drinken, buscopan + manitol. In spoedsetting zonder manitol en zonder buscopan. *bij vraag dundarm NET: +LA fase]
  Sonde					RAD ct abd 22 (-) [LOW DOSE CT uro, injectie 20-30 ml via sonde]

  Screening/staging pancreasCa/cholangioCa	RAD ct abd 22 (+) [bb art, abd veneus op 75s, 3-4 ml/s, 2 bekers WATER PO over 30 min (zeker bij duodenum)]
  Acuut bovenbuik (-itis)			RAD ct abd 22 (+) [bb art, abd veneus op 75s, 3-4 ml/s]
  Postop vasculair en collecties			RAD ct abd 22 (+) [bb art, abd veneus op 75s, 3-4 ml/s]
  FU hypervasc (NET, melanoom, schildklier)	RAD ct abd 22 (+) [bb art, abd veneus op 75s, 3-4 ml/s]
  FU acute BB (pancreatitis, cholecystitis)		RAD ct abd 22 (+) [abd veneus op 75s 3-4 ml/s]
  Screening/staging HCC / pretransplant		RAD ct abd 22 (+) [BB blanco/art, ABD veneus, nakijken of delayed fase (2-5 min postinjectie) nodig]

  Combi abdomenlijst				RAD ct comb 02 (+) [IV veneus + 3 PO]
  Combi Hypervasculair FU			RAD ct comb 02 (+) [BB art getriggerd cfr. abdomen bovenbuik protocol, thorax-abdomen veneus 75 sec]
  Combi +hersenen				RAD ct comb 03 (+) [IV veneus + 3 PO]
  Combi URO lijst				RAD ct uro 24 (+)
  Combi thoraxlijst				RAD ct thorax 20 (+)

  NOTA: gebruik ctrl-m en ctrl-z om snel met of zonder contrast te kiezen. ctrl-m zet ook automatisch IV veneus + 3 PO.
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


Data_uro_MR  := "
  (
  Prostaat standaard zc				RAD mr prostaat 01 zc (-)
  Prostaat met contrast				RAD mr prostaat mk (+)

  Nierletsel / niertumor (pre of postop)	RAD mr nier 08 (+)
  MR IVU / urinewegen / afloop			RAD mr nier 05 (+)

  MR Peritoneaalmetastasen 				RAD mr abd 41
  )"

helptext := "
  (
  Enter		-> Vul het eerste onderzoek in in KWS.
  Ctrl-Enter	-> Aanvaard het onderzoek in KWS (duwt op OK).
  Ctrl-NumpadEnter -> Aanvaard het onderzoek in KWS (duwt op OK).
  Ctrl-Numpad+	-> Selecteer 'met contrast'.
  Ctrl-Numpad-	-> Selecteer 'zonder contrast'.
  Ctrl-m		-> Selecteer 'met contrast'.
  Ctrl-z		-> Selecteer 'zonder contrast'.
  Ctrl-i		-> Zet de cursor in 'Opmerkingen'.
  Ctrl-l		-> Open het labo.
  Ctrl-q		-> Sluit het huidige KWS scherm (pt of labo).

  NB: Als abdomen CT is geselecteerd als discipline, zal Ctrl-Numpad+ ook 'IV veneus {+} 3 PO invullen'.
)"

inifile := "aanvaardingen.ini"

;;::startaanv::
  ;; Goto("start_aanvaardingen")

start_aanvaardingen:
  HotIfWinActive("Aanvaardingen helper ahk_class AutoHotkeyGUI")
  Hotkey("^Enter", druk_ok_aanvaarding)
  Hotkey("^NumpadEnter", druk_ok_aanvaarding)
  ;; Hotkey, Enter, aanvaard_onderzoek_knop ;; mss neit meer nodig met default
  Hotkey("^NumpadAdd", ctrnumplusHotkey)
  Hotkey("^NumpadSub", ctrlnumminHotkey)
  Hotkey("^m", ctrnumplusHotkey)
  Hotkey("^z", ctrlnumminHotkey)
  Hotkey("^i", selectOpmerking)
  Hotkey("^l", selectLabo)
  Hotkey("^q", sluitaanvaardschermKWS)
  toon_onderzoeken := ""
  global aanvaarder
  aanvaarder := Gui()
  aanvaarder.OnEvent("Close", aanvaarderGuiEscape)
  aanvaarder.OnEvent("Escape", aanvaarderGuiEscape)
  ogcEditonderzoek_naam := aanvaarder.Add("Edit", "x10 y9 w280 h20 vonderzoek_naam")
  ogcEditonderzoek_naam.OnEvent("Change", zoek_onderzoek_naam.Bind("Change"))
  onderzoek_naam := ogcEditonderzoek_naam.hwnd
  ogcButtonAanvaard := aanvaarder.Add("Button", "x295 y9 w180 h20  Default", "Aanvaard")
  ogcButtonAanvaard.OnEvent("Click", aanvaard_onderzoek_knop.Bind("Normal"))
  ogcDropDownListsubdiscipline := aanvaarder.Add("DropDownList", "x480 y9 w110 h20 R10 vsubdiscipline Choose1", ["neuro", "ORL", "abdomen (CT)", "abdomen (MR)", "thorax", "uro (CT)", "uro (MR)"])
  ogcDropDownListsubdiscipline.OnEvent("Change", set_subdiscipline.Bind("Change"))
  ogcButton := aanvaarder.Add("Button", "x595 y9 w15 h20", "?")
  ogcButton.OnEvent("Click", helpknop.Bind("Normal"))
  ogcEditGui_Display := aanvaarder.Add("Edit", "x10 y39 w600 h190 vGui_Display ReadOnly -wrap", toon_onderzoeken)
  aanvaarder.Title := "Aanvaardingen helper"
  aanvaarder.Show("x420 y781 h241 w620")
  WinSetAlwaysOnTop(1, "Aanvaardingen helper")
  if FileExist(inifile) {
	  subdiscipline := IniRead(inifile, "General", "subdiscipline", "neuro")
	  ogcDropDownListsubdiscipline.Choose(subdiscipline)
  }
  set_subdiscipline("", aanvaarder)
  Return

set_subdiscipline(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	Global Data, subdiscipline
	oSaved := aanvaarder.Submit("0")
	subdiscipline := oSaved.subdiscipline
	Switch subdiscipline
	{
		Case "neuro": Data := Data_neuro
		Case "ORL": Data := Data_orl
		Case "abdomen (CT)": Data := Data_abdomen_CT
		Case "abdomen (MR)": Data := Data_abdomen_MR
		Case "thorax": Data := Data_thorax
		Case "uro (CT)": Data := Data_uro_CT
		Case "uro (MR)": Data := Data_uro_MR
		Default: Data := Data_neuro
	}
	zoek_onderzoek_naam(A_GuiEvent, GuiCtrlObj, Info)
}

zoek_onderzoek_naam(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	Global Data
	oSaved := aanvaarder.Submit("0")
	onderzoek_naam := oSaved.onderzoek_naam
	toon_onderzoek := Sift_Regex(&Data, &onderzoek_naam, "uw")
	ogcEditGui_Display.Value := toon_onderzoek
}

aanvaard_onderzoek_knop(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	global Data
	oSaved := aanvaarder.Submit("0")
	onderzoek_naam := oSaved.onderzoek_naam
	onderzoek := Sift_Regex(&Data, &onderzoek_naam, "uw")
	onderzoek := StrSplit(onderzoek, "`n")[1]
	if RegExMatch(onderzoek, "(RAD .*[a-zA-Z0-9])\ +\((.)\)(?:\ +\[(.+)\])?", &gekozenOnderzoek)
		aanvaardOnderzoek(gekozenOnderzoek.1, gekozenOnderzoek.2, gekozenOnderzoek.3)
}

druk_ok_aanvaarding(ThisHotkey) {
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
  zoek_onderzoek_naam("Button", "")
  sleep(300)
  WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
  ; sleep(200)
  ; WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}

aanvaarderGuiEscape(*) {
	oSaved := aanvaarder.Submit("0")
	subdiscipline := oSaved.subdiscipline
	try
	IniWrite(subdiscipline, inifile,"General", "subdiscipline")
	ExitApp()
	Return
}

selectOpmerking(ThisHotkey) {
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	If (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\contrastLabel.png")) {
		MouseClick("Left", FoundX + 250, FoundY + 200)
		SendInput("^a")
		MouseMove(mouseX, mouseY)
	} else
		MsgBox("Het opmerkingen formulier werd niet gevonden.")
}

selectLabo(ThisHotkey) {
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	If (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\contrastLabel.png")) {
		MouseClick("Left", FoundX + 365, FoundY)
		sleep(550)
		If (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\eGFRLabel.png"))
			MouseClick("Left", FoundX + 3, FoundY + 3) ;; klik op GFR
		else
			MouseClick("Left", 950, 313) ;; even klikken op het KWS scherm om het gele vakje weg te krijgen.
		sleep(250)
		WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
		MouseMove(mouseX, mouseY)
	} else
		MsgBox("Het opmerkingen formulier werd niet gevonden.")
;	WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}

sluitaanvaardschermKWS(ThisHotkey) {
  WinActivate("KWS ahk_exe javaw.exe")
  sleep(200)
  Send("^{F4}")
  sleep(300)
  WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}

ctrnumplusHotkey(ThisHotkey) {
	global subdiscipline
	Switch subdiscipline {
		Case "abdomen (CT)": aanvaardOnderzoek("", "+", "IV veneus {+} 3 PO")
		Default: aanvaardOnderzoek("", "+", "")
	}
	WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}

ctrlnumminHotkey(ThisHotkey) {
	aanvaardOnderzoek("", "-", "")
	WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}

helpknop(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	MsgBox(helptext)
}

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
