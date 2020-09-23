Attribute VB_Name = "basHelp"
Option Explicit

Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3

Private Type tagHH_FTS_QUERY
cbStruct As Long
fUniCodeStrings As Long
pszSearchQuery As String
iProximity As Long
fStemmedSearch As Long
fTitleOnly As Long
fExecute As Long
pszWindow As String
End Type
Private Declare Function HTMLHelpStdCall Lib "hhctrl.ocx" _
Alias "HtmlHelpA" (ByVal hwnd As Long, _
ByVal lpHelpFile As String, _
ByVal wCommand As Long, _
ByVal dwData As Long) As Long
Private Declare Function HTMLHelpCallSearch Lib "hhctrl.ocx" _
Alias "HtmlHelpA" (ByVal hwnd As Long, _
ByVal lpHelpFile As String, _
ByVal wCommand As Long, _
ByRef dwData As tagHH_FTS_QUERY) As Long
Public Function ShowContents() As Long
    ShowContents = HTMLHelpStdCall(0, App.Path & "\QHelp.chm", HH_DISPLAY_TOC, 0)
End Function







