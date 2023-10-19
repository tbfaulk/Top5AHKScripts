#Requires Autohotkey v2.0+
#SingleInstance Force
#Include Lib\WinClip_ahk2-master\WinClip.ahk
#Include Lib\WinClip_ahk2-master\WinClipAPI.ahk

^Numpad9:: ;Jump to unread emails (requires you to be on a current gmail window)
{
Send("/")
Sleep 75
Send("is:unread {Enter}")
}

^Numpad8:: ;Start Google Meet and copy link to clipboard
{
Run("chrome.exe https://meet.new/")
Sleep 3000
Send "^l"
Sleep 20
Send "^c"
}


^!+F1:: ;Copy PDF version of current Excel file to clipboard
{
hwnd := ControlGetHwnd("EXCEL71"," - Excel")
xl := ObjectFromWindow(hwnd,)
CurrentXL := xl.Application.ActiveWorkbook.FullName
CurrentPDF :=StrReplace(CurrentXL, ".xlsm", ".pdf")
CurrentPDF :=StrReplace(CurrentPDF, ".xlsx", ".pdf")
wc := WinClip()
wc.SetFiles( CurrentPDF )
}


^Numpad7:: ;Tech Support Assistant
{
CommandGUI := Gui()
CommandGUI.Opt("AlwaysOnTop")
TV := CommandGUI.Add("TreeView", "h450 w330")

Mod1 := TV.Add("General")
Mod2 := TV.Add("Survey Module")
Mod3 := TV.Add("CG Survey Module")
Mod4 := TV.Add("Civil Module")
Mod5 := TV.Add("Hydrology Module")
Mod6 := TV.Add("GIS Module")
Mod7 := TV.Add("Geotech Module")
Mod8 := TV.Add("Construction Module")
Mod9 := TV.Add("Trench Module")
Mod10 := TV.Add("CADNET Module")
Mod11 := TV.Add("Field Module")
Mod12 := TV.Add("Natural Regrade Module")
Mod13 := TV.Add("Point Cloud Module")
Mod14 := TV.Add("Geology Module")
Mod15 := TV.Add("Underground Mining Module")
Mod16 := TV.Add("Surface Mining Module")

Menu1 := TV.Add("File Pulldown Menu", Mod1)
Menu2 := TV.Add("Edit Pulldown Menu", Mod1)
Menu3 := TV.Add("View Pulldown Menu", Mod1)
Menu4 := TV.Add("Draw Pulldown Menu", Mod1)
Menu5 := TV.Add("Inquiry Pulldown Menu", Mod1)
Menu6 := TV.Add("Settings Pulldown Menu", Mod1)
Menu7 := TV.Add("Points Pulldown Menu", Mod1)
Menu8 := TV.Add("Help Pulldown Menu", Mod1)

SubMenu9 := TV.Add("Tools", Menu2)
SubMenu11 := TV.Add("Clipboard", Menu2)
SubMenu17 := TV.Add("Erase", Menu2)
SubMenu18 := TV.Add("Copy", Menu2)
SubMenu19 := TV.Add("Offset", Menu2)
SubMenu20 := TV.Add("Explode", Menu2)
SubMenu21 := TV.Add("Extend", Menu2)
SubMenu22 := TV.Add("Break", Menu2)

Child223 := TV.Add("Print", SubMenu9)
Child224 := TV.Add("Open Link", SubMenu9)
Child225 := TV.Add("Change Space", SubMenu9)
Child226 := TV.Add("Image Manager...", SubMenu9)
Child227 := TV.Add("Check Spelling...", SubMenu9)
Child228 := TV.Add("Chamfer", SubMenu9)
Child229 := TV.Add("Fillet", SubMenu9)
Child230 := TV.Add("Measure", SubMenu9)
Child231 := TV.Add("Divide", SubMenu9)
Child232 := TV.Add("Xref Manager...", SubMenu9)
Child233 := TV.Add("Clip Xref", SubMenu9)
Child238 := TV.Add("Cut", SubMenu11)
Child240 := TV.Add("Copy with Base Point", SubMenu11)
Child241 := TV.Add("Paste", SubMenu11)
Child242 := TV.Add("Paste to Original Coordinates", SubMenu11)
Child243 := TV.Add("Paste as Block", SubMenu11)
Child244 := TV.Add("Paste Special...", SubMenu11)
Child265 := TV.Add("Find and Replace", SubMenu9)
Child283 := TV.Add("Select", SubMenu17)
Child284 := TV.Add("Single", SubMenu17)
Child285 := TV.Add("Last", SubMenu17)
Child286 := TV.Add("Erase By Layer...", SubMenu17)
Child287 := TV.Add("Erase By Closed Polyline", SubMenu17)
Child288 := TV.Add("Erase Outside", SubMenu17)
Child289 := TV.Add("Standard Copy", SubMenu18)


CommandGUI.Add("button", "w75", "OK").OnEvent("Click", Ok_Click)
CommandGUI.Add("button", "w75 x+1", "Cancel").OnEvent("Click", Cancel_Click)

Cancel_Click(*)
{CommandGUI.Destroy()}

Ok_Click(*)
{
final := TV.GetSelection()
outputtext := TV.GetText(final)
newparent := TV.GetParent(final)

While (newparent)
    {
    newparenttext := TV.GetText(newparent)
    outputtext := newparenttext " > " outputtext
    newparent := TV.GetParent(newparent)
    }

CommandGUI.Destroy()
outputtext := StrReplace(outputtext, "General > ", "")
Send outputtext
}
CommandGUI.Show()
}