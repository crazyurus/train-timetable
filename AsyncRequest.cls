VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_xmlhttp As Object
Private m_callbackTarget As Object
Private m_callbackMethod As String
Private m_aborted As Boolean

Public Sub GetRequest(ByVal url As String, ByRef callbackTarget As Object, ByVal callbackMethod As String)
    On Error GoTo ErrHandler
    
    Set m_callbackTarget = callbackTarget
    m_callbackMethod = callbackMethod
    m_aborted = False
    
    Set m_xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    On Error GoTo ErrHandler
    m_xmlhttp.Open "GET", url, True
    m_xmlhttp.onreadystatechange = Me
    m_xmlhttp.Send
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "AsyncRequest.GetRequest", Err.Description
End Sub

Public Sub PostRequest(ByVal url As String, ByVal postBody As String, ByRef callbackTarget As Object, ByVal callbackMethod As String, Optional ByRef extraHeaders As Object = Nothing)
    On Error GoTo ErrHandler
    
    Set m_callbackTarget = callbackTarget
    m_callbackMethod = callbackMethod
    m_aborted = False
    
    Set m_xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    m_xmlhttp.setOption 2, 13056
    
    On Error GoTo ErrHandler
    m_xmlhttp.Open "POST", url, True
    m_xmlhttp.onreadystatechange = Me
    m_xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    If Not extraHeaders Is Nothing Then
        Dim key As Variant
        For Each key In extraHeaders
            m_xmlhttp.setRequestHeader CStr(key), CStr(extraHeaders(key))
        Next
    End If
    
    m_xmlhttp.Send postBody
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "AsyncRequest.PostRequest", Err.Description
End Sub

Public Sub DefaultProc()
Attribute DefaultProc.VB_UserMemId = 0
    On Error Resume Next
    
    If m_aborted Then Exit Sub
    If m_xmlhttp Is Nothing Then Exit Sub
    If m_xmlhttp.readyState <> 4 Then Exit Sub
    
    If Not m_callbackTarget Is Nothing Then
        CallByName m_callbackTarget, m_callbackMethod, VbMethod, m_xmlhttp.status, m_xmlhttp.responseText
    End If
End Sub

Public Sub Abort()
    m_aborted = True
    If Not m_xmlhttp Is Nothing Then
        m_xmlhttp.Abort
        Set m_xmlhttp = Nothing
    End If
    Set m_callbackTarget = Nothing
End Sub

Private Sub Class_Terminate()
    If Not m_xmlhttp Is Nothing Then
        Set m_xmlhttp = Nothing
    End If
    Set m_callbackTarget = Nothing
End Sub
