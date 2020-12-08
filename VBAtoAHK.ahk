; Created by AHK_User
; Used to convert recorded Excel VBA Code to AHK com Code
; Already working fine in a lot of cases, but a lot of improvements can be made
; 2020-11-22: Added Word translation and Macro Explorer (to make Macro Explorer work, please lower the macro security)
; 2020-12-08: Added automatic download of constants and set function to work global

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
global arrMacro := {}
global arrXlsConstants := VBAtoAHK_LoadArrConstants()
global vProgram

Gui, 1: Default

Menu, FileMenu, Add, &Open Macro Explorer, VBAtoAHK_GetExcelWindows  ; See remarks below about Ctrl+O.
Menu, MyMenuBar, Add, &File, :FileMenu  ; Attach the two sub-menus that were created above.
Gui, Menu, MyMenuBar

Gui, Add, Text,, VBA Code:
Gui, Add, DropDownList, vvProgram ggVBA_Code x430 yp w80, Excel||Word ; (idea to also make it work on other programs than Excel)
Gui, Add, Edit, vvVBA_Code ggVBA_Code xm y30 r30 w500 t5, %Clipboard%
Gui, Add, Text, ym x520, AHK Code:
Gui, Add, CheckBox, vvInitiation x900 yp ggVba_Code, Show initiation.
Gui, Add, Edit, vvAHK_Code x520 y30 r30 w500  t5, % VBAtoAHK(Clipboard)
Gui, Add, Button, ggRun_AHK_Code, Run  ; The label ButtonOK (if it exists) will be run when the button is pressed.
Gui, Add, Button, ggEdit_AHK_Code x+5 yp, Edit  ; The label ButtonOK (if it exists) will be run when the button is pressed.
Gui, Show,, VBA to AHK code
GoSub, gVBA_Code
return  ; End of auto-execute section. The script is idle until the user does something.

gVBA_Code:
Gui, Submit, NoHide
AHK_Code := VBAtoAHK(vVBA_Code)
if (vInitiation and !InStr(AHK_Code, "Create a connection to an Excel Application")){
	VBAtoAHK_AddInitiation(AHK_Code)
}
Guicontrol, Text, vAHK_Code, % AHK_Code
return

gRun_AHK_Code:
Gui, Submit, NoHide
FileAHK := A_ScriptDir "\AHK_Code.ahk"
FileDelete, %FileAHK%
FileAppend,  #SingleInstance`, Force`, %FileAHK%
VBAtoAHK_AddInitiation(vAHK_Code)
FileAppend, %vAHK_Code% , %FileAHK%
Run, % FileAHK
return

gEdit_AHK_Code:
Gui, Submit, NoHide
FileAHK := A_ScriptDir "\AHK_Code.ahk"
FileDelete, %FileAHK%
FileAppend,  #SingleInstance`, Force`, %FileAHK%
VBAtoAHK_AddInitiation(vAHK_Code)
FileAppend, %vAHK_Code% , %FileAHK%
Run edit "%FileAHK%"
return

GuiClose:
Gui, Submit  ; Save the input from the user to each control's associated variable.
ExitApp

MenuHandler:
MsgBox, test
return

VBAtoAHK(VBA_Code){
	global
	; Cleans the VBA returns
	VBA_Code:= RegExReplace(VBA_Code, "_\R\s*", "")
	
	AHK_Code := ""
	
	ArrWith := {}
	ArrWith["Level"] := 0
	WithCount := 0
	
	loop, Parse, VBA_Code, `n, `r
	{
		Line := A_LoopField
		
		
		Line := RegExReplace(Line, "^(\s*.*?=\s-?\d*)\.(\d.*)$", "$1,$2")
		FirstWord := RegExReplace(Line, "^\s*(\w*).*$", "$1")
		
		if !RegExMatch(Line, "^\s*If\s"){
			Line := StringCodeReplace(Line, " =", " :=")
		}
		
		if (vProgram="Excel"){
			; Adding missing leading Object definitions
			Line := RegExReplace(Line, "([\s(:=]+|^)(Sheets)", "$1oExcel.ActiveWorkbook.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Worksheets)", "$1oExcel.ActiveWorkbook.$2")
			
			Line := RegExReplace(Line, "([\s(:=]+|^)(Cells)", "$1oWorkSheet.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Columns)", "$1oWorkSheet.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Range\()", "$1oWorkSheet.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Rows\()", "$1oWorkSheet.$2")
			
			Line := RegExReplace(Line, "([\s(:=]+|^)(ActiveCell)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(ActiveSheet)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(ActiveWindow)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Application)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(InputBox)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Selection)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(UsedRange)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Visible)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Workbooks)", "$1oExcel.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(WorksheetFunction)", "$1oExcel.$2")
			
			Line := RegExReplace(Line, "([\s(:=]+|^)(ActiveWorkbook)", "$1oExcel.Application.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(DisplayAlerts)", "$1oExcel.Application$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(EnableEvents)", "$1oExcel.Application$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(ScreenUpdating)", "$1oExcel.Application$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Calculation)", "$1oExcel.Application$2")
			
			; For Range Object
			VBAtoAHK_ParameterCheck(Line, "AutoFilter", "Field|Criteria1|Operator|Criteria2|SubField|VisibleDropDown")
			VBAtoAHK_ParameterCheck(Line, "AutoFitt", "Destination|Type")
			
			; For Sheet Object
			VBAtoAHK_ParameterCheck(Line, "Add", "Before|After|Count|Type")
			
		}
		Else if (vProgram="Word"){
			; Adding missing leading Object definitions
			Line := RegExReplace(Line, "([\s(:=]+|^)(ActiveDocument)", "$1oWord.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Documents)", "$1oWord.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(Selection)", "$1oWord.$2")
			Line := RegExReplace(Line, "([\s(:=]+|^)(PageSetup)", "$1oDoc.$2")
			
			Line := RegExReplace(Line, "([\s(:=]+|^)(Visible)", "$1oWord.$2")
			
			VBAtoAHK_ParameterCheck(Line, "TypeText", "Text")
			VBAtoAHK_ParameterCheck(Line, "EndKey", "Unit|Extend")
			
			Line := RegExReplace(Line, "^(\s*.*?\.MoveLeft)\s(.*)$", "$1($2)")
			Line := RegExReplace(Line, "^(\s*.*?\.MoveRight)\s(.*)$", "$1($2)")
			Line := RegExReplace(Line, "^(\s*.*?\.Tables.Add)\s(Range.*)$", "$1($2)")
			Line := RegExReplace(Line, "^(\s*.*?\.InsertRowsBelow)\s(.*)$", "$1($2)")
			
		}
		
		Line := RegExReplace(Line, "i)^(\s*)Set\s(.*)$", "$1$2")
		
		; Converting comments
		Line := RegExReplace(Line, "\s'", " `;")
		Line := RegExReplace(Line, "^'", "`;")
		Line := RegExReplace(Line, "^(\s*)Dim\s", "$1`; Dim ")
		
		if RegExMatch(Line, "^\s*With\s"){
			WithCount++
			ArrWith["Level"]++
			ArrWith[ArrWith["Level"]] := RegExReplace(Line, "(^\s*)(\w*)\s(.*)$", "$3")
			VarWith := RegExReplace(Line, "(^\s*)(\w*)\s(.*)$", "$3")
			VarWithSpaces := RegExReplace(Line, "(^\s*)(\w*)\s(.*)$", "$1")
			
			ArrWith[ArrWith["Level"], "Spaces"] := RegExReplace(Line, "(^\s*)(\w*)\s(.*)$", "$1")
			if InStr(VarWith, "."){
				Line := RegExReplace(Line, "^(\s*)With\s(.*)$", "$1oWith" WithCount " := $2")
				ArrWith[ArrWith["Level"]] := "oWith" WithCount
			}
			Else{
				ArrWith[ArrWith["Level"]] := VarWith
				Continue
			}
		}
		
		; Correcting With...
		Line := RegExReplace(Line, "(\s|\()\.", "$1" ArrWith[ArrWith["Level"]] ".")
		if (VarWith!=""){
			Line := RegExReplace(Line, "^(\s*)(.*)$", VarWithSpaces "$2")
		}
		
		; Functions
		Line := RegExReplace(Line, "^(\s*.*?\.[A-Z0-9][A-Za-z0-9]*)\s("".*"")\s*$", "$1($2)")
		
		; Handling if else statements
		Line := RegExReplace(Line, "i)^(\s*If)\s(.*)\sThen$", "$1($2){")
		Line := RegExReplace(Line, "i)^(\s*)ElseIf\s(.*)\sThen$", "}`n$1Else If($2){")
		Line := RegExReplace(Line, "i)^(\s*)End If$", "$1}")
		Line := RegExReplace(Line, "i)^(\s*)Else$", "$1}`n$1Else{")
		
		; Defining variables
		PreConstants := "xl|mso|wd|sig|rgb|pp|pb|ol|_xl|Backstage|Broadcast|cert|cipher|contverres|empty|enc|mf"
		Variable := RegExReplace(Line, "(:=\s?|\()(" PreConstants ")[A-Z0-9][A-Za-z0-9]*", "$1", RegexCount)
		LineEnd := Line
		LineStart := ""
		Loop, %RegexCount%
		{
			Variable := RegExReplace(LineEnd, "^.*?(\s?:=\s?|\()((" PreConstants ")[A-Z0-9][A-Za-z0-9]*).*$", "$2", RegexCount)
			LineStart := LineStart RegExReplace(LineEnd, "^(.*?(\s?:=\s?|\()(" PreConstants ")[A-Z0-9][A-Za-z0-9]*)(.*)$", "$1 := " arrXlsConstants[Variable] , RegexCount)
			LineEnd := RegExReplace(LineEnd, "^(.*?(\s?:=\s?|\()(" PreConstants ")[A-Z0-9][A-Za-z0-9]*)(.*)$", "$4", RegexCount)
			Line := LineStart LineEnd
		}
		
		; With statements can be skipped, but we will safe the object and the leading spaces
		
		if RegExMatch(Line, "^\s*ChDir\s"){
			Continue
		}
		if RegExMatch(Line, "^\s*End\sSub"){
			Line := "return`n}"
		}
		if RegExMatch(Line, "^\s*End\sWith"){
			ArrWith["Level"]--
			VarWith := ""
			VarWithSpaces := ""
			Continue
		}
		if RegExMatch(Line, "^\s*Sub\s"){
			if (!InStr(VBA_Code, "End Sub")){
				Continue
			}
			Line := RegExReplace(Line, "(^\s*)(\w*)\s(.*)$", "$1$3{`n$1global")
		}
		
		Line := StringCodeReplace(Line, " & ", " ")
		
		Line := RegExReplace(Line, "i)^(.*:=\s?\d*)\.(\d*)$", "$1,$2")
		AHK_Code.= Line "`n"
		if (vProgram="Excel"){
			; Connect back to activesheet when deleting or adding sheets
			if (InStr(Line, "Sheets.Delete") or InStr(Line, "Sheets.Add")){
				AHK_Code.= "oWorkSheet := oExcel.ActiveSheet`n"
			}
		}
	}
	;~ DebugWindow(AHK_Code)
	return AHK_Code
}

VBAtoAHK_LoadArrConstants(){
	if !FileExist(A_ScriptDir "\Constants.txt"){
		URLDownloadToFile, https://github.com/dmtr99/VBAtoAHK/raw/main/Constants.txt, %A_ScriptDir%\Constants.txt
	}
	FileRead, xlsInterfaceConstants, %A_ScriptDir%\Constants.txt
	
	arrXlsConstants := {}
	loop, Parse, xlsInterfaceConstants, `n, `r
	{
		Variable := RegExReplace(A_LoopField, "^([^\s]*)\s.\s([^\s]*)$", "$1")
		Value := RegExReplace(A_LoopField, "^([^\s]*)\s.\s([^\s]*)$", "$2")
		arrXlsConstants[Variable] := Value
	}
	return arrXlsConstants
}

VBAtoAHK_ParameterCheck(ByRef Line, Method, ListParameters){
	if InStr(Line, "." Method " "){
		Line1 := RegExReplace(Line, "^(.*\." Method ")\s.*", "$1(")
		Loop, Parse, ListParameters , |
		{
			Var := RegExReplace(Line, ".*\." Method "\s.*(" A_LoopField "\s?:?=[^,]*).*", "$1", RegexCount)
			if (RegexCount=1){
				Line1 .= Var ","
			}
			Else{
				Line1 .= ","
			}
		}
		Line := Line1 ")"
		Line := RegExReplace(Line, "^(.*?),*\)$", "$1)")
	}
	return Line
}

VBAtoAHK_AddInitiation(ByRef Code){
	if (!InStr(Code, "Create a connection to an Excel Application") and vProgram="Excel"){
		Code := "Try{`n`toExcel := ComObjActive(""Excel.Application"") `; Create a connection to an Excel Application object`n}`nCatch{`n`toExcel := ComObjCreate(""Excel.Application"") `; Create an Excel Application object`n`toExcel.Visible := true`n}`n`noWorkSheet := oExcel.ActiveSheet`n" Code 
	}
	else if (!InStr(Code, "Create a connection to an Word Application") and vProgram = "Word"){
		Code := "Try{`n`toWord := ComObjActive(""Word.Application"") `; Create a connection to an Word Application object`n}`nCatch{`n`toWord := ComObjCreate(""Word.Application"") `; Create an Word Application object`n`toWord.Visible := true`n}`n`noDoc := oWord.ActiveDocument`n" Code
	}
	else if (!InStr(Code, "Create a connection to an Outlook Application") and vProgram = "Outlook"){
		Code := "Try{`n`toOutlook := ComObjActive(""Outlook.Application"") `; Create a connection to an Outlook Application object`n}`nCatch{`n`toOutlook := ComObjCreate(""Outlook.Application"") `; Create an Outlook Application object`n`oOutlook.Visible := true`n}`n`n" Code
	}
	Return
}

StringCodeReplace(Haystack, SearchText, ReplaceText:=""){
	ReplacedStr := ""
	StrReplace(Haystack, """", , OutputVarCount)
	if (OutputVarCount=0){
		return StrReplace(Haystack, SearchText, ReplaceText)
	}
	Loop, Parse, Haystack, `"
	{
		if (mod(A_Index, 2) = 1){
			ReplacedStr .= StrReplace(A_LoopField, SearchText, ReplaceText)
		}
		else{
			ReplacedStr .= A_LoopField
		}
		if (A_index!=OutputVarCount+1){
			ReplacedStr .= """"
		}
	}
	return ReplacedStr
}

VBAtoAHK_GetExcelWindows(){
	global
	oExcel := excel_get()
	if IsObject(oExcel){
		list := "Excel.exe`n"
		Loop % oExcel.Workbooks.count {
			;~ MsgBox, % xl.Workbooks(a_index).name
			list .= "`t" oExcel.Workbooks(a_index).name "`n"
			FileName := oExcel.Workbooks(a_index).name
			For cmpComponent In oExcel.Workbooks(a_index).VBProject.VBComponents
			{
				szFileName := cmpComponent.Name
				list .= "`t`t" szFileName "`n"
				; Read contents
				LinesContent := cmpComponent.CodeModule.Lines(1,cmpComponent.CodeModule.CountOfLines)
				arrMacro["Excel.exe\" FileName "\" szFileName] := LinesContent
				
				LinesContent0 := StrReplace(LinesContent, "`r", "")
				loop, 
				{
					CodeContent := RegexReplace(LinesContent0, "(^.*?Sub\s[^\(]*?\(.*?End Sub).*", "$1", RegexCount)
					Name_Macro := RegexReplace(CodeContent, "^.*?Sub ([^\(\s]*?)\(.*", "$1")
					LinesContent0 := RegexReplace(LinesContent0, "^.*?Sub ([^\(\s]*?\(.*?End Sub)(.*)$", "$2")
					if (RegexCount=0){
						Break
					}
					list .= "`t`t`t" Name_Macro "`n"
					arrMacro["Excel.exe\" FileName "\" szFileName "\" Name_Macro] := CodeContent
				}
			}	
		}
	}
	oWord := ComObjActive("Word.Application")
	if IsObject(oWord){
		list .= "Word.exe`n"
		Loop % oWord.documents.count {
			oDocument := oWord.documents(A_Index)
			
			list .= "`t" oDocument.name "`n"
			FileName := oWord.documents(A_Index).name
			
			For cmpComponent In oWord.documents(A_Index).VBProject.VBComponents
			{
				szFileName := cmpComponent.Name
				list .= "`t`t" szFileName "`n"
				; Read contents
				LinesContent := ""
				try LinesContent := cmpComponent.CodeModule.Lines(1,cmpComponent.CodeModule.CountOfLines)
				if (LinesContent =""){
					LinesContent := "Does not work, seems to work differently than Excel, sorry.`n`nSuggestions are welcome..."
				}
				arrMacro["Word.exe\" FileName "\" szFileName] := LinesContent
				
				LinesContent0 := StrReplace(LinesContent, "`r", "")
				loop, 
				{
					CodeContent := RegexReplace(LinesContent0, "(^.*?Sub\s[^\(]*?\(.*?End Sub).*", "$1", RegexCount)
					Name_Macro := RegexReplace(CodeContent, "^.*?Sub ([^\(\s]*?)\(.*", "$1")
					LinesContent0 := RegexReplace(LinesContent0, "^.*?Sub ([^\(\s]*?\(.*?End Sub)(.*)$", "$2")
					if (RegexCount=0){
						Break
					}
					list .= "`t`t`t" Name_Macro "`n"
					arrMacro["Word.exe\" FileName "\" szFileName "\" Name_Macro] := CodeContent
				}
			}	
		}
	}
	
	Gui, 2:New, +HwndhGui +Resize
	
	ImageListID := IL_Create(10)
	TreeViewList =
	
	IL_Add(ImageListID, "shell32.dll", "5") ; folder
	IL_Add(ImageListID, "shell32.dll", "3")
	
	Gui 2: Add, TreeView, xm w250 h400 0x400 vvMacroExplTree ggMacroExplTree hwndHTV ImageList%ImageListID% ; AltSubmit		;0x400 single expand, 0x200 hot tracking 
	Ar_Tree := CreateTreeView(list)
	Gui 2: Add, Edit, x+10 yp w400 h400 vvMacroContent, 
	Gui 2: Add, Button, gg2Translate, Translate 
	ItemID := 0  ; Causes the loop's first iteration to start the search at the top of the tree.
	Loop
	{
		ItemID := TV_GetNext(ItemID, "Full")  ; Replace "Full" with "Checked" to find all checkmarked items.
		if not ItemID  ; No more items in tree.
			break
		ParentID := TV_GetParent(ItemID)
		
		if (ParentID = "0"){
			TV_Modify(ItemID ,"Expand")
		}
	}
	
	Gui, 2:Add, StatusBar, , 
	Gui, 2:Show, , Macro Explorer
	
	GuiName := A_DefaultGui 
	return
	
	gMacroExplTree:
	Gui +OwnDialogs
	
	SelectedItemID := TV_GetSelection()
	TV_GetText(ItemText, SelectedItemID)
	
	if (A_GuiEvent = "DoubleClick"){
		TV_GetText(DoubleClickedItemText, A_EventInfo)
		TV_GetText(SelectedText, A_EventInfo)
		ItemID := TV_GetSelection()
		Path := ItemText
		ParentID := TV_GetParent(ItemID)
		Loop, {
			TV_GetText(ParentText, ParentID)
			if (ParentText =""){
				break
			}
			Path := ParentText "\" Path
			ParentID := TV_GetParent(ParentID)
		}
		GuiControl, , vMacroContent, % arrMacro[Path]
	}
	return
	g2Translate:
	Gui, 2:Submit, NoHide
	Gui, 1:Default
	Guicontrol, Text, vVBA_Code, % vMacroContent
	GoSub gVBA_Code
	return
}
	
	Excel_Get(WinTitle="ahk_class XLMAIN") {	; by Sean and Jethrow, minor modification by Learning one
		ControlGet, hwnd, hwnd, , Excel71, %WinTitle%
		if !hwnd
			return
		Window := Acc_ObjectFromWindow(hwnd, -16)
		Loop
			try
				Application := Window.Application
		catch
			ControlSend, Excel71, {esc}, %WinTitle%
		Until !!Application
		return Application
	}	; http://www.autohotkey.com/forum/viewtopic.php?p=492448#492448
	
	;---------------------------------------------------------------------------------------------------------
	;CreateTreeView function
	;adding auto "sort" items ---append two instances of match3 with " Sort"
	;---------------------------------------------------------------------------------------------------------
	CreateTreeView(TreeViewDefinitionString) {	; by Learning one
		Ar_Tree := {}
		IDs := {}
		Loop, parse, TreeViewDefinitionString, `n, `r
		{
			if A_LoopField is space
				continue
			Item := RegExReplace(A_LoopField,"AD)(\t*[^\t]+)\t.+","$1",Count)
			
			Item := RTrim(Item, A_Space A_Tab), Item := LTrim(Item, A_Space), Level := 0
			While (SubStr(Item,1,1) = A_Tab)
				Level += 1,	Item := SubStr(Item, 2)
			RegExMatch(Item, "([^`t]*)([`t]*)([^`t]*)", match)	; match1 = ItemName, match3 = Options
			if (Level=0){
				;~ ItemID := TV_Add(match1, 0, match3 "Icon2" )
				ItemID := TV2_Add(match1, 0, match3)
				IDs["Level0"] := ItemID
			}
			else{
				ItemID := TV2_Add(match1, IDs["Level" Level-1], match3)
				;~ if !InStr(match1,"."){
				;~ ItemID := TV_Add(match1, IDs["Level" Level-1], match3 "Icon2 Sort")
				;~ }
				IDs["Level" Level] := ItemID
			}
			Ar_Tree[ItemID]	:= A_LoopField
			
		}
		return Ar_Tree
	}	; http://www.autohotkey.com/board/topic/92863-function-createtreeview/
	
	TV2_Add(ItemName, ParentID="", Options=""){
		if !InStr(ItemName,"."){
			ItemID := TV_Add(ItemName, ParentID, "Icon2 " Options)
		}
		else{
			ItemID := TV_Add(ItemName, ParentID, " " Options)
		}
		return ItemID
	}
	
	
	
