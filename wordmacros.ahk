#Requires AutoHotKey v2
#SingleInstance Force ;when you run another instance it will replace previous one

;The following hotstrings will run only when Windows Word application is active.

#HotIf WinActive("ahk_class OpusApp")
;Currently the following word macros are triggered with hotstrings.
;you can replace with them hotkeys
;
::;;rel;;::
{
    Reload() ;reloads the current AHK script. This is handy if you made any modifications to this script.
}
; insert table - opens dialog with selection on "only label and number" (For example   "Table 1")
::;;instable::
{
    Send("{LAlt Down}n{LAlt Up}rf")
    Sleep (500)
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Enter}")
    Send ("{Tab}")
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Enter}")
    Send("{Tab}{Down}")
}

::;;insfig::
{
    Send("{LAlt Down}n{LAlt Up}rf")
    Sleep (500)
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Down}{Down}{Down}{Down}{Down}{Enter}")
    Send ("{Tab}")
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Enter}")
    Send("{Tab}{Down}")
}

::;;insxref::
{
    Send("{LAlt Down}n{LAlt Up}rf")
    Sleep (500)
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Down}{Enter}")
    Send ("{Tab}")
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Enter}")
    Send("{Tab}{Down}")
}
;Here is another hotkey to open bookmarks GUI (AHK gui) and insert a field (cross-reference)
;Shows using Word COM OBJECT
;Open the bookmarks window and insert the selected crossreference
^!+b::
{
    oWord := ComObjActive("Word.Application")
    oDocument := oWord.ActiveDocument
    oCount := oDocument.Bookmarks.Count
    
    myGUI := Gui()
    myGUI.MarginX := 16
    myGUI.MarginY := 16
    myGUI.BackColor := "Black"
    MyGui.Opt("+AlwaysOnTop -SysMenu +Owner -Caption")  
    myGUI.SetFont("s10 cGreen q5 w400", "Segoe UI")
    myGUI.Add('Text', ,"Bookmarks")
    
    listCtrl := myGUI.Add('ListBox', 'vSkull r10 w220', [])
    listCtrl.SetFont("s12 cBlack q5 w400", "Consolas")
    listCtrl.opt("-Redraw") ;for huge numbers this will speed up. First add and then redraw
    ;Add bookmarks to the listbox
    while (oCount > 0)
    {
        listCtrl.Add([oDocument.Bookmarks.Item(oCount).Name])
        oCount--
    }
    listCtrl.opt("+Redraw")
    MyBtn := MyGui.Add("Button", "Default w80", "OK")
    MyBtn.OnEvent("Click", MyBtn_Click)  ; Call MyBtn_Click when clicked.
    myGUI.OnEvent("Escape", GuiEscape)
    myGUI.OnEvent("Close", GuiEscape)
    ;listCtrl.OnEvent("Change",Findus)
    myGUI.Show()

    ;-------------------------------------------------------------------------------
    ; GUI FUNCTIONS AND SUBROUTINES
    ;-------------------------------------------------------------------------------
    MyBtn_Click(*)
    {
        myGUI.Submit
        your_bookmark := listCtrl.Text
        gui_destroy()
       
        ;The following code shows inserting various fields at the current cursor selection
        ;Field add is described at https://learn.microsoft.com/en-us/office/vba/api/Word.Fields.Add
        ;Field type values can be identified at https://learn.microsoft.com/en-us/office/vba/api/word.wdfieldtype
        wdFieldDate := 31
        wdFieldRef := 3 
        ;insert date field
        ;oField := oDocument.Fields.Add(oWord.Selection.Range, wdFieldDate)
        oField := oDocument.Fields.Add(oWord.Selection.Range, wdFieldRef, your_bookmark, True)
    }
    ; Automatically triggered on Escape key:
    GuiEscape(*)
    {
        gui_destroy()
    }
    ; The callback function when the text changes in the input field.
    Findus(*)
    {
        Saved := myGUI.Submit(false)
        your_bookmark := Saved.Skull
        
    }
    ; gui_destroy: Destroy the GUI after use.
    ;
    #WinActivateForce
    gui_destroy() {
        myGUI.Destroy()
    }  
}
;Here is another hotkey to insert a date field
;Shows using Word COM OBJECT
;Open the bookmarks window and insert the selected crossreference
::;;datefield::
{
    Sleep(400) ;sleep otherwise while deleting the hotstring the date is inserted and gets deleted

    oWord := ComObjActive("Word.Application")
    oDocument := oWord.ActiveDocument
    ;The following code shows inserting various fields at the current cursor selection
    ;Field add is described at https://learn.microsoft.com/en-us/office/vba/api/Word.Fields.Add
    ;Field type values can be identified at https://learn.microsoft.com/en-us/office/vba/api/word.wdfieldtype
    wdFieldDate := 31
    wdFieldRef := 3 
    ;insert date field
    oField := oDocument.Fields.Add(oWord.Selection.Range, wdFieldDate) 
}
#HotIf
