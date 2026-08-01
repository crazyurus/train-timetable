Attribute VB_Name = "Utilities"
Option Explicit

Public Function CalcDuration(ByVal strStartTime As String, ByVal strEndTime As String) As String
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim lngMinutes As Long
    Dim intHours As Integer
    Dim intMins As Integer
    On Error Resume Next
    If Len(strStartTime) >= 5 And Len(strEndTime) >= 5 Then
        dblStart = TimeValue(Left(strStartTime, 5))
        dblEnd = TimeValue(Left(strEndTime, 5))
        lngMinutes = DateDiff("n", dblStart, dblEnd)
        If lngMinutes < 0 Then
            lngMinutes = lngMinutes + 24 * 60
        End If
        intHours = lngMinutes \ 60
        intMins = lngMinutes Mod 60
        CalcDuration = Format(intHours, "00") & ":" & Format(intMins, "00")
    Else
        CalcDuration = "----"
    End If
End Function

Public Function GetJsonArray(ByVal strJson As String) As String
    Dim lngBracket As Long
    Dim i As Long
    Dim blnInStr As Boolean
    Dim strChr As String
    lngBracket = 0
    blnInStr = False
    For i = 1 To Len(strJson)
        strChr = Mid(strJson, i, 1)
        If strChr = """" Then
            blnInStr = Not blnInStr
        ElseIf Not blnInStr Then
            If strChr = "[" Then
                lngBracket = lngBracket + 1
            ElseIf strChr = "]" Then
                lngBracket = lngBracket - 1
                If lngBracket = 0 Then
                    GetJsonArray = Left(strJson, i)
                    Exit Function
                End If
            End If
        End If
    Next i
    GetJsonArray = strJson
End Function

Public Function HttpGet(ByVal strUrl As String) As String
    Dim objHttp As Object
    Dim strResult As String
    On Error GoTo ErrorHandler
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error Resume Next
    objHttp.setOption 2, 13056
    On Error GoTo ErrorHandler
    With objHttp
        .Open "GET", strUrl, False
        .Send
        .WaitForResponse 10000
        strResult = .ResponseText
    End With
    HttpGet = strResult
    Set objHttp = Nothing
    Exit Function
ErrorHandler:
    Err.Raise vbObjectError + 1001, "HttpGet", "HTTPÇëÇóÊ§°Ü: " & Err.Description
End Function

Public Function EncodeURL(ByVal strText As String) As String
    Dim objScr As Object
    Dim strResult As String

    Set objScr = CreateObject("MSScriptControl.ScriptControl")
    objScr.Language = "JScript"
    strResult = objScr.CodeObject.encodeURIComponent(strText)
    EncodeURL = strResult
    Set objScr = Nothing
End Function

Public Function SplitJsonObject(ByVal strJson As String) As String()
    Dim arrResult() As String
    Dim lngCount As Long
    Dim i As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngBrace As Long
    Dim blnInStr As Boolean
    Dim strChr As String
    ReDim arrResult(0 To 99)
    lngCount = 0
    lngStart = 0
    lngBrace = 0
    blnInStr = False
    For i = 1 To Len(strJson)
        strChr = Mid(strJson, i, 1)
        If strChr = """" Then
            blnInStr = Not blnInStr
        ElseIf Not blnInStr Then
            If strChr = "{" Then
                If lngBrace = 0 Then
                    lngStart = i
                End If
                lngBrace = lngBrace + 1
            ElseIf strChr = "}" Then
                lngBrace = lngBrace - 1
                If lngBrace = 0 Then
                    lngEnd = i
                    If lngCount <= UBound(arrResult) Then
                        arrResult(lngCount) = Mid(strJson, lngStart, lngEnd - lngStart + 1)
                        lngCount = lngCount + 1
                    End If
                End If
            End If
        End If
    Next i
    If lngCount > 0 Then
        ReDim Preserve arrResult(0 To lngCount - 1)
    Else
        ReDim arrResult(0 To -1)
    End If
    SplitJsonObject = arrResult
End Function

Public Function GetJsonValue(ByVal strJson As String, ByVal strKey As String) As String
    Dim strSearch As String
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strChr As String
    strSearch = """" & strKey & """:"
    lngPos = InStr(strJson, strSearch)
    If lngPos = 0 Then
        GetJsonValue = ""
        Exit Function
    End If
    lngStart = lngPos + Len(strSearch)
    Do While lngStart <= Len(strJson)
        strChr = Mid(strJson, lngStart, 1)
        If strChr <> " " And strChr <> vbTab And strChr <> vbCr And strChr <> vbLf Then
            Exit Do
        End If
        lngStart = lngStart + 1
    Loop
    If lngStart > Len(strJson) Then
        GetJsonValue = ""
        Exit Function
    End If
    strChr = Mid(strJson, lngStart, 1)
    If strChr = """" Then
        lngStart = lngStart + 1
        lngEnd = InStr(lngStart, strJson, """")
        If lngEnd = 0 Then
            GetJsonValue = Mid(strJson, lngStart)
        Else
            Dim strTemp As String
            strTemp = Mid(strJson, lngStart, lngEnd - lngStart)
            strTemp = Replace(strTemp, "\""", """")
            strTemp = Replace(strTemp, "\\", "\")
            strTemp = Replace(strTemp, "\/", "/")
            strTemp = Replace(strTemp, "\n", vbLf)
            strTemp = Replace(strTemp, "\r", vbCr)
            strTemp = Replace(strTemp, "\t", vbTab)
            GetJsonValue = strTemp
        End If
    Else
        lngEnd = lngStart
        Do While lngEnd <= Len(strJson)
            strChr = Mid(strJson, lngEnd, 1)
            If strChr = "," Or strChr = "}" Or strChr = "]" Then
                Exit Do
            End If
            lngEnd = lngEnd + 1
        Loop
        GetJsonValue = Trim(Mid(strJson, lngStart, lngEnd - lngStart))
    End If
End Function

