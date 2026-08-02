VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_index As Long
Private m_source As String

Public Function Parse(ByVal jsonString As String) As Object
    On Error GoTo ErrHandler
    
    m_source = jsonString
    m_index = 1
    
    SkipWhitespace
    
    If PeekChar = "{" Then
        Set Parse = ParseObject
    ElseIf PeekChar = "[" Then
        Set Parse = ParseArray
    Else
        Err.Raise vbObjectError + 1001, "JSON.Parse", "Invalid JSON: expected object or array"
    End If
    
    Exit Function
ErrHandler:
    Set Parse = Nothing
End Function

Private Function ParseObject() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ConsumeChar "{"
    SkipWhitespace
    
    If PeekChar = "}" Then
        ConsumeChar "}"
        Set ParseObject = dict
        Exit Function
    End If
    
    Do
        SkipWhitespace
        Dim key As String
        key = ParseString
        SkipWhitespace
        ConsumeChar ":"
        SkipWhitespace
        
        Dim value As Variant
        ReadValueInto value
        
        dict.Add key, value
        
        SkipWhitespace
        
        If PeekChar = "," Then
            ConsumeChar ","
        ElseIf PeekChar = "}" Then
            Exit Do
        Else
            Err.Raise vbObjectError + 1002, "JSON.ParseObject", "Expected ',' or '}' at position " & m_index
        End If
    Loop
    
    ConsumeChar "}"
    Set ParseObject = dict
End Function

Private Function ParseArray() As Object
    Dim col As Collection
    Set col = New Collection
    
    ConsumeChar "["
    SkipWhitespace
    
    If PeekChar = "]" Then
        ConsumeChar "]"
        Set ParseArray = col
        Exit Function
    End If
    
    Do
        SkipWhitespace
        Dim value As Variant
        ReadValueInto value
        col.Add value
        SkipWhitespace
        
        If PeekChar = "," Then
            ConsumeChar ","
        ElseIf PeekChar = "]" Then
            Exit Do
        Else
            Err.Raise vbObjectError + 1003, "JSON.ParseArray", "Expected ',' or ']' at position " & m_index
        End If
    Loop
    
    ConsumeChar "]"
    Set ParseArray = col
End Function

Private Sub ReadValueInto(ByRef result As Variant)
    Dim ch As String
    ch = PeekChar
    
    Select Case ch
        Case "{"
            Dim obj As Object
            Set obj = ParseObject
            Set result = obj
        Case "["
            Dim arr As Object
            Set arr = ParseArray
            Set result = arr
        Case """"
            result = ParseString
        Case "t", "f"
            result = ParseBoolean
        Case "n"
            result = ParseNull
        Case "-", "0" To "9"
            result = ParseNumber
        Case Else
            Err.Raise vbObjectError + 1004, "JSON.ReadValueInto", "Unexpected character '" & ch & "' at position " & m_index
    End Select
End Sub

Private Function ParseValue() As Variant
    Dim tmp As Variant
    ReadValueInto tmp
    If IsObject(tmp) Then
        Set ParseValue = tmp
    Else
        ParseValue = tmp
    End If
End Function

Private Function ParseString() As String
    Dim result As String
    Dim ch As String
    
    ConsumeChar """"
    
    Do While m_index <= Len(m_source)
        ch = Mid$(m_source, m_index, 1)
        
        Select Case ch
            Case """"
                m_index = m_index + 1
                ParseString = result
                Exit Function
                
            Case "\"
                m_index = m_index + 1
                If m_index > Len(m_source) Then Exit Do
                
                ch = Mid$(m_source, m_index, 1)
                Select Case ch
                    Case """"
                        result = result & """"
                    Case "\"
                        result = result & "\"
                    Case "/"
                        result = result & "/"
                    Case "b"
                        result = result & vbBack
                    Case "f"
                        result = result & vbFormFeed
                    Case "n"
                        result = result & vbLf
                    Case "r"
                        result = result & vbCr
                    Case "t"
                        result = result & vbTab
                    Case "u"
                        m_index = m_index + 1
                        Dim hexStr As String
                        hexStr = Mid$(m_source, m_index, 4)
                        If Len(hexStr) < 4 Then
                            Err.Raise vbObjectError + 1005, "JSON.ParseString", "Invalid unicode escape"
                        End If
                        result = result & ChrW(CLng("&H" & hexStr))
                        m_index = m_index + 3
                    Case Else
                        result = result & ch
                End Select
                m_index = m_index + 1
                
            Case Else
                result = result & ch
                m_index = m_index + 1
        End Select
    Loop
    
    Err.Raise vbObjectError + 1006, "JSON.ParseString", "Unterminated string"
End Function

Private Function ParseBoolean() As Boolean
    Dim ch As String
    ch = PeekChar
    
    If ch = "t" Then
        If StrComp(Mid$(m_source, m_index, 4), "true", vbBinaryCompare) = 0 Then
            m_index = m_index + 4
            ParseBoolean = True
            Exit Function
        End If
    ElseIf ch = "f" Then
        If StrComp(Mid$(m_source, m_index, 5), "false", vbBinaryCompare) = 0 Then
            m_index = m_index + 5
            ParseBoolean = False
            Exit Function
        End If
    End If
    
    Err.Raise vbObjectError + 1007, "JSON.ParseBoolean", "Invalid boolean value"
End Function

Private Function ParseNull() As Variant
    If StrComp(Mid$(m_source, m_index, 4), "null", vbBinaryCompare) = 0 Then
        m_index = m_index + 4
        ParseNull = Null
        Exit Function
    End If
    
    Err.Raise vbObjectError + 1008, "JSON.ParseNull", "Invalid null value"
End Function

Private Function ParseNumber() As Double
    Dim startPos As Long
    startPos = m_index
    
    If Mid$(m_source, m_index, 1) = "-" Then
        m_index = m_index + 1
    End If
    
    Do While m_index <= Len(m_source)
        Dim ch As String
        ch = Mid$(m_source, m_index, 1)
        If ch >= "0" And ch <= "9" Then
            m_index = m_index + 1
        Else
            Exit Do
        End If
    Loop
    
    If m_index <= Len(m_source) Then
        ch = Mid$(m_source, m_index, 1)
        If ch = "." Then
            m_index = m_index + 1
            Do While m_index <= Len(m_source)
                ch = Mid$(m_source, m_index, 1)
                If ch >= "0" And ch <= "9" Then
                    m_index = m_index + 1
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    
    If m_index <= Len(m_source) Then
        ch = Mid$(m_source, m_index, 1)
        If ch = "e" Or ch = "E" Then
            m_index = m_index + 1
            If m_index <= Len(m_source) Then
                ch = Mid$(m_source, m_index, 1)
                If ch = "+" Or ch = "-" Then
                    m_index = m_index + 1
                End If
            End If
            Do While m_index <= Len(m_source)
                ch = Mid$(m_source, m_index, 1)
                If ch >= "0" And ch <= "9" Then
                    m_index = m_index + 1
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    
    ParseNumber = CDbl(Mid$(m_source, startPos, m_index - startPos))
End Function

Private Sub SkipWhitespace()
    Do While m_index <= Len(m_source)
        Dim ch As String
        ch = Mid$(m_source, m_index, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            m_index = m_index + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function PeekChar() As String
    If m_index <= Len(m_source) Then
        PeekChar = Mid$(m_source, m_index, 1)
    Else
        PeekChar = ""
    End If
End Function

Private Sub ConsumeChar(expected As String)
    If PeekChar <> expected Then
        Err.Raise vbObjectError + 1009, "JSON.ConsumeChar", "Expected '" & expected & "' at position " & m_index & ", got '" & PeekChar & "'"
    End If
    m_index = m_index + 1
End Sub
