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
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Enter}")
    Send ("{Tab}")
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Enter}")
    Send("{Tab}{Down}")
}

::;;insfig::
{
    Send("{LAlt Down}n{LAlt Up}rf")
    Sleep (500)
    Send("{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Down}{Down}{Down}{Down}{Down}{Enter}")
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
;get the current selected text and create a bookmark
^!+b::
{
    oWord := ComObjActive("Word.Application")
    ;MsgBox(oWord.selection.text)
    ;TODO
}

#HotIf