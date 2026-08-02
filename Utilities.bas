Attribute VB_Name = "Utilities"
Option Explicit

Public Function EncodeURL(ByVal strText As String) As String
    Dim objScr As Object
    Dim strResult As String

    Set objScr = CreateObject("MSScriptControl.ScriptControl")
    objScr.Language = "JScript"
    strResult = objScr.CodeObject.encodeURIComponent(strText)
    EncodeURL = strResult
    Set objScr = Nothing
End Function

