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
  Cervicobrachialgie (CWK)				RAD mr wz 01  (-)
  Cervicaal CWK/DWK Post-Op			RAD mr wk 19  (-)
  Intraspinale Metastasen				RAD mr wk 23  (+)
  LWK/DWK Standaard				RAD mr wk 18  (-)
  LWK Post-op					RAD mr wk 15  (+)
  Metaal						RAD mr wk 22  (-)
  Congenitale Aandoening LWK (pediatrie)		RAD mr wk 21  (-)
  Full Spine						RAD mr full spine (-)
  Ruggenmerg (MS of Myelitis)				RAD mr wz 13  (-)
  Ruggenmergletsel (geen full spine)			RAD mr wk 16  (+)
  Scoliose						RAD mr wk 24  (-)
  Spinale AV-Fistel(enkel bij controle)			RAD mr wk 25  (+)
  Liquorlek						RAD mr wk 26  (?)
  Zonder contrast							(-)
  Met contrast							(+)
  Aneurysma					RAD amr hersen 57 (-)
  Arterio-Veneuze Malformatie (AVM)			RAD mr hersen 71  (+)
  Bloeding						RAD mr hersen 65  (+)
  Caverneuze Malformatie/Fabry			RAD mr hersen 67  (-)
  Cerebellair Letsel (tumor)				RAD mr hersen 13  (+)
  CVA-TIA						RAD mr hersen 04  (-)
  Dementie						RAD mr hersen 11  (-)
  Epilepsie, Hippocampus				RAD mr hersen 06  (-)
  Frameless stereotaxie (neurochirurgie/navig)		RAD mr hersen 54  (+)
  Hoofdpijn						RAD mr hersen 24  (-)
  Hydrocefalie / NPH					RAD mr hersen 64  (-)
  Hypotensie					RAD mr hersen 72  (?)
  Idiopathische Intracraniale Hypertensie			RAD mr hersen 73  (-)
  Liquorlek						RAD mr hersen 68  (-)
  MS (Vermoeden MS)				RAD mr hersen 02  (?)
  MS Follow-up					RAD mr hersen 03  (-)
  Neurofibromatosis/Fakomatose			RAD mr hersen 74  (+)
  Neurovasculair Conflict V/VII/IX			RAD mr hersen 75  (+)
  Parkinson						RAD mr hersen 57  (-)
  Post-operatieve Controle				RAD mr hersen 27  (+)
  Routineprotocol					RAD mr hersen 25  (-)
  Sella Macro-adenoma				RAD mr hersen 08  (+)
  Sella Micro-adenoma				RAD mr hersen 09  (+)
  Sella Post-operatief					RAD mr hersen 10  (+)
  Stereotactisch Radiotherapie				RAD mr hersen 89  (+)
  Tumor met Perfusie					RAD mr hersen 77  (+)
  Tumor/infectie zonder Perfusie			RAD mr hersen 78  (+)
  Vasculitis						RAD amr hersen 58 (+)
  Veneuze Trombose					RAD mr hersen 14  (-)
  Vertigo/Duizeligheid					RAD mr hersen 29  (-)
  Vestibulair Schwannoom				RAD mr hersen 70  (+)
  Trauma						RAD mr hersen 32  (-)
  Epilepsie, Hippocampus Pathologie (pediatrie)		RAD mr hersen 46  (-)
  Langerhanscelhistiocyste				RAD mr hersen 81  (+)
  Neonato						RAD mr hersen 84  (-)
  Neurofibromatose/fakomatose (pediatrie)		RAD mr hersen 85  (+) [kinderen > 3jaar gebruik T2_mv  i.p.v. destir]
  PVL						RAD mr hersen 23  (-)
  Routineprotocol (pediatrie)				RAD mr hersen 87  (-)
  Tumor (Pediatrie)					RAD mr hersen 88  (+)
  Ontwikkelingsstoornis (Pediatrie)			RAD mr hersen 86 (-) [Als >3 jaar gebruik T2_mv ipv destir]
  MR Orbita					RAD mr orbita (+)
  Sella (Pediatrie) Pubertas Praecox			RAD mr hersen 42  (-)
  MR orbita (Pediatrie) sinus cavernosus/orbita		RAD mr orbita 03  (+) [Indien < 3j: AX DESTIR ipv ax T2 TSE]
  MR orbita (Pediatrie) opticus hypoplasie			RAD mr orbita 02  (-) [Indien > 3jaar gebruik T2_mv i.p.v. destir]
  MR Halsvaten/craniocervicaal			RAD amr hersen 60  (?) [Beter met contrast als de bepaling van de stenosegraad verplicht is]
  MR Halsvaten + hersenen				RAD amr hersen 61  (?)
  MR Halsvaten Dissectie				RAD amr hersen 62  (?)
  MR Halsvaten + hersenen dissectie			RAD amr hersen 63  (?)
  Studie 						(?)
  )"

Data_abdomen_MR := "
  (
	Abdomen Peritoneaalmetastasen			RAD mr abd 41 (+)
	Pancreas Tumor					RAD mr abd 55 (+)
	Chronische Pancreatitis				RAD mr abd 34 (-)
	Pancreas Familiaal Carcinoma				RAD mr abd 33 (-)
	Pancreas voor (acute/collecties) pancreatitis		RAD mr abd 08 (+)
	Pancreas simpele IPMN				RAD mr abd 34 (-) (Als complexe cyste: tumorprotocol)
	Lever + Contrast, zonder Laattijdige			RAD mr abd 48 (+)
	Lever + Contrast, met Laattijdige			RAD mr abd 49 (+)
	HCC / FNH / adenoma (Lever met Laattijdige)		RAD mr abd 49 (+)
	Lever zonder contrast				RAD mr abd 50 (-) (zeldzaam: screening meta's, hoewel beter met contrast zonder late)
	Lever hemochromatose				RAD mr abd 35 (?)
	Levertransplantatie					RAD mr abd 03 (+)
	Cholangio zonder contrast (MRCP)			RAD mr abd 51 (-)
	Cholangio met contrast (PSC, IGG4, maligne stenose)	RAD mr abd 53 (+) (beter lever + MRCP)
	Cholangio gallek					RAD mr abd 52 (?) [Multihance en laattijdige niet vergeten!]
	Staging Recto-Anale Tumor				RAD mr abd 29 (+)
	Anale fistel (Crohn)					RAD mr abd 21 (+) [Als fistel zichtbaar, zeker zien dat die volledig in beeld is!]
	Ovaria IOTA					RAD mr abd 44 (+)
	Cervix  						RAD mr cervix (+)
	Endometriumcarcinoom 				RAD mr abd 63 (+)
	Uterusfibromen 					RAD mr abd 61 (+)
	Endometriose (lang: standaard)			RAD mr abd 59 (+)
	Endometriose (kort)					RAD mr abd 58 (-) (Nooit van extern/niet-gynaeco)
	Vrouwelijke genitalia (congenitaal, adenomoyose)		RAD mr abd 60 (-) [Bellen voor assen]
	Entero						RAD mr abd 37 (+) (door Ragna te aanvaarden!)
	Acuut bij zwangerschap				RAD mr abd 47 (-) (Te bespreken met supervisie)

	Lynch  						RAD mr abd 45 (+) [Enkel MR4, zo niet peritoneaalmetas protocol]
	Rectoanale Pouch					RAD mr abd 21 (+) [Buscopan en rectale vulling met water. T2 TRUFI 5 mm in 3 richtingen. T1 cor/ax vibe pre C 3 mm. 3D T1 caipi vibe ax pre C. Contrast + 2e helft buscopan. T1 cor/ax vibe richtingen post C 3 mm. 3D T1 caipi vibe ax post C. DWI ax. T2 TSE fs indien nodig in nuttigste vlak (bijvoorbeeld fistel)]
	Nierletsel / niertumor (pre of postop)			RAD mr nier 08 (+)
	Volumetrie
	Studie 						 (?) [Volgens studie]
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
  Metastasen				RAD ct thorax 21 (-)
  )"

Data_uro_CT  := "
  (
  Bijnier					RAD ct uro 08 (+)
  Nier vasculair				RAD ct uro 100 (+)
  Niertumor primair				RAD ct uro 19 (+)
  Pyelonefritis				RAD ct uro 10 (+)
  Low dose lithiasis				RAD ct uro 13 (-)
  IVU					RAD ct uro 20 (+)
  Follow up/screening abdomen met aflopen	RAD ct uro 30 (+)
  Combi thorax/abdomen			RAD ct uro 24 (+)
  Combi RCC				RAD ct uro 24 (+) [Graag arteriele van de thorax opentrekken over de bovenbuik; dan klassieke veneuze abdomen]
  Combi + hersenen				RAD ct uro 27 (+)
  Combi thorax/abdomen + afloop 		RAD ct uro 29 (+)
  Trauma					RAD ct uro 22 (+)
  Nierdonor					RAD ct uro 21 (+)
  Niervolumetrie				RAD ct uro 28 (-) [Indien ADPKD normale dosis]
  )"


Data_uro_MR  := "
  (
  Prostaat standaard zc				RAD mr prostaat 01 zc (-)
  Prostaat met contrast				RAD mr prostaat mk (+)
  Prostaat veld / radiotherapie 				RAD mr v prostaat 02 (-)

  Nierletsel / niertumor (pre of postop)			RAD mr nier 08 (+)
  IVU / urinewegen / afloop				RAD mr nier 05 (+)
  Transplantnier (cave geen angio!)			RAD mr tpnier 01 (+)
  Bijnier						RAD mr bijnier 02 mc (?) [graag bellen voor contrast wordt gegeven.]
  Bijnier met contrast					RAD mr bijnier 02 mc (+)
  Bijnier zonder contrast				RAD mr bijnier 01 zc (-)
  Gemetastaseerd RCC abdomen			RAD mr nier 09 (+)

  MR blaas 					RAD mr blaas (+)
  MR urethra zonder contrast 				RAD mr urethra 01 (-)
  MR urethra met contrast 				RAD mr urethra 02 (+)
  MR penis 					RAD mr penis 01 (+)
  MR scrotum 					RAD mr scrotum 01 (+)

  Angio renale vaten   		 		RAD amr nier 01 (+)
  Angio transplantnier vaten   		 		RAD amr tpnier 02 (+)
  Peritoneaalmetastasen 				RAD mr abd 41
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
  Ctrl-d		-> Toon document.
  Ctrl-q		-> Sluit het huidige KWS scherm (pt of labo).

  NB: Als abdomen CT is geselecteerd als discipline, zal Ctrl-Numpad+ ook 'IV veneus {+} 3 PO invullen'.
)"

inifile := "settings.ini"

HotIfWinActive("Aanvaardingen helper ahk_class AutoHotkeyGUI")
Hotkey("^Enter", druk_ok_aanvaarding)
Hotkey("^NumpadEnter", druk_ok_aanvaarding)
;; Hotkey, Enter, vulin_knop ;; mss neit meer nodig met default
Hotkey("^NumpadAdd", ctrnumplusHotkey)
Hotkey("^NumpadSub", ctrlnumminHotkey)
Hotkey("^m", ctrnumplusHotkey)
Hotkey("^z", ctrlnumminHotkey)
Hotkey("^i", selectOpmerking)
Hotkey("^l", selectLabo)
Hotkey("^d", selectToonDocument)
Hotkey("^q", sluitaanvaardschermKWS)
^Backspace::Send("^+{Left}{Backspace}{Backspace}")

toon_onderzoeken := ""
global aanvaarder
aanvaarder := Gui()
aanvaarder.OnEvent("Close", aanvaarderGuiEscape)
aanvaarder.OnEvent("Escape", aanvaarderGuiEscape)
ogcEditonderzoek_naam := aanvaarder.Add("Edit", "x10 y9 w220 h20 vonderzoek_naam")
ogcEditonderzoek_naam.OnEvent("Change", zoek_onderzoek_naam.Bind("Change"))
onderzoek_naam := ogcEditonderzoek_naam.hwnd
ogcButtonVulIn := aanvaarder.Add("Button", "x235 y9 w100 h20  Default", "Vul in [enter]")
ogcButtonVulIn.OnEvent("Click", vulin_knop.Bind("Normal"))
ogcButtonAanvaard := aanvaarder.Add("Button", "x340 y9 w135 h20", "Aanvaard [ctrl-enter]")
ogcButtonAanvaard.OnEvent("Click", druk_ok_aanvaarding.Bind())
ogcDropDownListsubdiscipline := aanvaarder.Add("DropDownList", "x480 y9 w110 h20 R10 vsubdiscipline Choose1", ["neuro", "ORL", "abdomen (CT)", "abdomen (MR)", "thorax", "uro (CT)", "uro (MR)"])
ogcDropDownListsubdiscipline.OnEvent("Change", set_subdiscipline.Bind("Change"))
ogcButton := aanvaarder.Add("Button", "x595 y9 w15 h20", "?")
ogcButton.OnEvent("Click", helpknop.Bind("Normal"))
ogcEditGui_Display := aanvaarder.Add("Edit", "x10 y39 w600 h190 vGui_Display ReadOnly -wrap", toon_onderzoeken)
aanvaarder.Title := "Aanvaardingen helper"
aanvaarder.Show("x1220 y780 h241 w620")
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

vulin_knop(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	global Data
	oSaved := aanvaarder.Submit("0")
	onderzoek_naam := oSaved.onderzoek_naam
	onderzoek := Sift_Regex(&Data, &onderzoek_naam, "uw")
	if onderzoek = ""
			return
	onderzoek := StrSplit(onderzoek, "`n")[1]
	if RegExMatch(onderzoek, "(?<oz>RAD .*[a-zA-Z0-9])?[ \t]+\((?<contrast>[+-?])\)(?:\ +\[(?<opm>.+)\])?", &gekozenOnderzoek)
		aanvaardOnderzoek(gekozenOnderzoek.oz, gekozenOnderzoek.contrast, gekozenOnderzoek.opm)
}

druk_ok_aanvaarding(ThisHotkey) {
	global Data
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	ErrorLevel := !ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\okButton.png")
	if (ErrorLevel = 2 or ErrorLevel = 1) {
			;; MsgBox("Er is iets fout gegaan met zoeken naar de OK knop (niet gevonden of afbeelding bestaat niet)")
			return
	}
	MouseClick("left", FoundX+5, FoundY+5)
	MouseMove(mouseX, mouseY)
	ogcEditonderzoek_naam.Text := "" ;; clear zoekbalk
	ogcEditGui_Display.Value := Data
	; zoek_onderzoek_naam("Button", "")
	focus_on_aanvaarder_timer(200)
}

aanvaarderGuiEscape(*) {
	oSaved := aanvaarder.Submit("0")
	subdiscipline := oSaved.subdiscipline
	try
	IniWrite(subdiscipline, inifile,"General", "subdiscipline")
	ExitApp()
	Return
}

selectToonDocument(ThisHotkey) {
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	If (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\okButton.png")) {
		MouseClick("left", FoundX + 180, FoundY + 5)
		MouseMove(mouseX, mouseY)
		focus_on_aanvaarder_timer()
	} else
		MsgBox("Het referentiepunt voor de knop werd niet gevonden!")
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
		loop(2) {
			If (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "images\eGFRLabel.png"))
				MouseClick("Left", FoundX + 3, FoundY + 3) ;; klik op GFR
			else
				MouseClick("Left", 950, 318) ;; even klikken op het KWS scherm om het gele vakje weg te krijgen.
			sleep(250)
		}
		focus_on_aanvaarder_timer()
		MouseMove(mouseX, mouseY)
	} else
		MsgBox("Het opmerkingen formulier werd niet gevonden.")
;	WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}

sluitaanvaardschermKWS(ThisHotkey) {
  WinActivate("KWS ahk_exe javaw.exe")
  sleep(300)
  Send("^{F4}")
  sleep(50)
  focus_on_aanvaarder_timer()
}

ctrnumplusHotkey(ThisHotkey) {
	global subdiscipline
	Switch subdiscipline {
		Case "abdomen (CT)": aanvaardOnderzoek("", "+", "IV veneus {+} 3 PO")
		Default: aanvaardOnderzoek("", "+", "")
	}
	focus_on_aanvaarder_timer()
}

ctrlnumminHotkey(ThisHotkey) {
	aanvaardOnderzoek("", "-", "")
	focus_on_aanvaarder_timer()
}

helpknop(A_GuiEvent, GuiCtrlObj, Info := "", *) {
	MsgBox(helptext)
}

aanvaardOnderzoek(onderzoekCode := "", contrast := "?", opmerking := "") {
	WinActivate("KWS ahk_exe javaw.exe")
	MouseGetPos(&mouseX, &mouseY)
	Clipboard := ""
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
			SendInput("^x")
			SendInput(onderzoekCode) ; typ het onderzoek in de regel van aanvaarding.
			Sleep(200)
			;; Checkt nog eens om zeker te zijn dat er niets verkeer aanvaard wordt.
			if (InStr(onderzoekCode, "mr hersen") and RegExMatch(A_Clipboard, "mr c?l?w[zk]"))
					MsgBox("CAVE: rug aangevraagd maar hersenen aanvaard!!!")
			if (RegExMatch(onderzoekCode, "mr w[kz] (01|19|18|15|22)") and RegExMatch(A_Clipboard, "mr wz"))
					MsgBox("CAVE: full spine aangevraagd maar LWZ/CWZ aanvaard!!!")
			if (RegExMatch(onderzoekCode, "mr cwz") and RegExMatch(A_Clipboard, "mr wk (18|15)"))
					MsgBox("CAVE: cervicale aangevraagd maar LWZ aanvaard!!!")
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
			opmerking := StrReplace(opmerking, "+", "{+}")
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
	;; focus_on_aanvaarder_timer(200)
	focus_on_aanvaarder_timer(600)
}

focus_on_aanvaarder_timer(time := -500) {
	WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
	time := Abs(time) * (-1) ;; Zorgt dat time altijd negatief is
	focus_on_aanvaarder()
	SetTimer(focus_on_aanvaarder.bind(), time) ;; do it again after x miliseconds
}

focus_on_aanvaarder() {
		WinActivate("Aanvaardingen helper ahk_class AutoHotkeyGUI")
}
		