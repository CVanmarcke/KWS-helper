#Include KWSHandler.ahk

; TODO: 
; GUI voor extra dingen
; In sheet gooien afh van type verslag (hersenen, ...)

; Script broke after adding the GUI.


KWStoExcel() {
	
	MsgBox, Got here
	path := "P:\Neuro cases-tips.xlsx"
	
	copyKWSreporttoclip()
	
	if not WinExist("Excel") {
		Run, %path%
		WinWait, "Excel"
	}
	try {

		;XL := ComObjCreate("Excel.Application")    ; create a (new) instance of Excel
		;XL.Visible := true                         ; make Excel visible
		;XL.Workbooks.Open("P:\Neuro cases-tips.xlsx")    
		; XL := ComObjActive("Excel.Application")

		XL := ComObjGet(path)
		
		RegexQuery := "(?:Leuven|Pellenberg)[\s\S]+(?<datum>\d{2}-\d{2}-\d{4})[\s\S]+(?:KLINISCHE INLICHTINGEN:[\n\r])(?<klinlicht>[\s\S]+)\R{2}(?:DIAGNOSTISCHE VRAAGSTELLING:[\n\r])(?<diagvraag>[\s\S]+)[\n\r]{2}(?:ONDERZOEKE?N?:[\n\r])(?<onderzoek>(?:.+\R?)+)\R{2,}(?:[\s\S]+)"
		RegExMatch(clipboard, RegexQuery, report)

		if WinExist("Pt. ") 
			WinActivate 
		WinGetTitle, title, A
		ead := SubStr(title, InStr(title, "(")+1, 8)
		
		sheetlist := ["Default", "thorax", "neuro", "abdomen"]
		
		sheet := "Default"
		if RegExMatch(reportonderzoek, "i)hersenen|schedel") {
			sheet := "neuro"
		} else if RegExMatch(reportonderzoek, "i)thorax") {
			sheet := "thorax"
		} else if RegExMatch(reportonderzoek, "i)abdomen") {
			sheet := "abdomen"
		}
		
		XL.Sheets(sheet).Activate
		lastCell := excelFindLastCell(XL, sheet).row + 1
		
		Gui, ExcelGUI:Add, Edit, x12 y9 w130 h30 vOnderzoek, %reportonderzoek%
		Gui, ExcelGUI:Add, Edit, x152 y9 w100 h30 +ReadOnly , %reportdatum%
		Gui, ExcelGUI:Add, Edit, x262 y9 w80 h30 +ReadOnly, %ead%
		Gui, ExcelGUI:Add, Text, x12 y52 w130 h30 , Klin. inlichtingen
		Gui, ExcelGUI:Add, Edit, x152 y49 w190 h30 vKlinInl, %reportklinlicht%
		Gui, ExcelGUI:Add, Text, x12 y89 w130 h30 , Diagn. vraagstelling
		Gui, ExcelGUI:Add, Edit, x152 y89 w190 h30 vDiagnVraag, %reportdiagvraag%
		Gui, ExcelGUI:Add, Text, x12 y129 w130 h30 , Diagnose
		Gui, ExcelGUI:Add, Edit, x152 y129 w190 h30 vDx, 
		Gui, ExcelGUI:Add, Text, x12 y169 w130 h30 , Tags
		Gui, ExcelGUI:Add, Edit, x152 y169 w190 h30 vTags ,
		Gui, ExcelGUI:Add, DropDownList, x12 y209 w130 h30 vSheetSelect, Default|thorax|neuro|abdomen
		Gui, ExcelGUI:Add, Button, x152 y209 w190 h20 , OK
		GuiControl, Choose, SheetSelect, ObjIndexOf(sheetlist, sheet)
		; Generated using SmartGUI Creator 4.0
		Gui, ExcelGUI:Show, x360 y233 h245 w357, Save to excel script
		
		return
		ExcelGUIButtonOK:
		/*
		XL.Sheets(sheet).range("A"lastCell).value := ead
		XL.Sheets(sheet).range("B"lastCell).value := reportonderzoek
		XL.Sheets(sheet).range("C"lastCell).value := reportdatum
		XL.Sheets(sheet).range("D"lastCell).value := reportklinlicht
		XL.Sheets(sheet).range("E"lastCell).value := reportdiagvraag
		*/
		XL.Sheets(sheet).range("A"lastCell).value := ead
		XL.Sheets(sheet).range("B"lastCell).value := Onderzoek
		XL.Sheets(sheet).range("C"lastCell).value := reportdatum
		XL.Sheets(sheet).range("D"lastCell).value := KlinInl
		XL.Sheets(sheet).range("E"lastCell).value := DiagnVraag
		XL.Sheets(sheet).range("F"lastCell).value := Dx
		XL.Sheets(sheet).range("G"lastCell).value := Tags
		
		_makeSplashText(title := "Saved to excel", text := ead . " is saved to excel", time := -1500)
		
		ExcelGUIGuiEscape:
		ExcelGUIGuiClose:
		ExitApp
		
	} catch e {
		MsgBox, An error was thrown: %e%
		MsgBox, 4112, Error - Edit Mode, The code cannot be executed while in edit mode (or other error).
	}
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