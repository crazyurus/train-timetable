Attribute VB_Name = "StationData"
Option Explicit

Private Type StationEntry
    Name As String
    Code As String
    Pinyin As String
    PinyinShort As String
End Type

Private m_stations() As StationEntry
Private m_stationCount As Long
Private m_dictByName As Object
Private m_loading As Boolean
Private m_loaded As Boolean

Public Sub ParseStationList(ByVal jsContent As String)
    Dim lines() As String
    Dim parts() As String
    Dim i As Long, j As Long
    
    If m_loaded Then Exit Sub
    
    On Error GoTo ErrHandler
    Set m_dictByName = CreateObject("Scripting.Dictionary")
    
    lines = Split(jsContent, "@")
    ReDim m_stations(0 To UBound(lines))
    
    j = 0
    For i = 1 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo ContinueLoop
        parts = Split(lines(i), "|")
        If UBound(parts) >= 2 Then
            m_stations(j).Name = parts(1)
            m_stations(j).Code = parts(2)
            If UBound(parts) >= 3 Then m_stations(j).Pinyin = parts(3)
            If UBound(parts) >= 4 Then m_stations(j).PinyinShort = parts(4)
            m_dictByName(parts(1)) = j
            j = j + 1
        End If
ContinueLoop:
    Next i
    
    m_stationCount = j
    If j > 0 Then ReDim Preserve m_stations(0 To j - 1)
    m_loaded = True
    m_loading = False
    Exit Sub
    
ErrHandler:
    m_loading = False
End Sub

Public Function SearchStations(ByVal keyword As String) As Collection
    Dim result As Collection
    Dim i As Long
    Dim entry As Object
    
    Set result = New Collection
    If Len(keyword) = 0 Or m_stationCount = 0 Then
        Set SearchStations = result
        Exit Function
    End If
    
    keyword = UCase(keyword)
    For i = 0 To m_stationCount - 1
        If InStr(1, UCase(m_stations(i).Name), keyword, vbTextCompare) > 0 Or _
           InStr(1, UCase(m_stations(i).Pinyin), keyword, vbTextCompare) > 0 Or _
           InStr(1, UCase(m_stations(i).PinyinShort), keyword, vbTextCompare) > 0 Then
            Set entry = CreateObject("Scripting.Dictionary")
            entry("name") = m_stations(i).Name
            entry("code") = m_stations(i).Code
            result.Add entry
            If result.Count >= 20 Then Exit For
        End If
    Next
    Set SearchStations = result
End Function

Public Function GetStationCode(ByVal name As String) As String
    If m_dictByName Is Nothing Then
        GetStationCode = ""
        Exit Function
    End If
    If m_dictByName.Exists(name) Then
        GetStationCode = m_stations(CLng(m_dictByName(name))).Code
    Else
        GetStationCode = ""
    End If
End Function

Public Function GetStationName(ByVal code As String) As String
    Dim i As Long
    code = UCase(code)
    For i = 0 To m_stationCount - 1
        If UCase(m_stations(i).Code) = code Then
            GetStationName = m_stations(i).Name
            Exit Function
        End If
    Next
    GetStationName = ""
End Function

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_loaded
End Property

Public Property Get IsLoading() As Boolean
    IsLoading = m_loading
End Property

Public Sub SetLoading()
    m_loading = True
End Sub