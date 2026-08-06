VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.UserControl StationPicker 
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   1215
   ScaleWidth      =   3000
   Begin VBCCR18.ListBoxW lstSuggest 
      Height          =   570
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1005
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin VBCCR18.TextBoxW txtInput 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrDebounce 
      Interval        =   300
      Left            =   2280
      Top             =   360
   End
End
Attribute VB_Name = "StationPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event StationChanged(ByVal StationName As String, ByVal StationCode As String)

Private m_asyncLoad As AsyncRequest
Private m_strStationName As String
Private m_strStationCode As String
Private m_blnSkipChange As Boolean
Private m_normHeight As Long
Private m_arrSuggestCodes() As String

Private Sub UserControl_Initialize()
    m_normHeight = 375
    
    txtInput.Move 0, 0, UserControl.ScaleWidth, 375
    lstSuggest.Move 0, 360, UserControl.ScaleWidth, 0
    lstSuggest.Visible = False
    
    tmrDebounce.Enabled = False
End Sub

Private Sub UserControl_InitProperties()
    m_normHeight = 375
End Sub

Private Sub UserControl_Resize()
    txtInput.Move 0, 0, UserControl.ScaleWidth, 375
    If lstSuggest.Visible Then
        lstSuggest.Move 0, 360, UserControl.ScaleWidth, lstSuggest.Height
    Else
        m_normHeight = UserControl.Height
    End If
End Sub

Public Property Get Text() As String
    Text = txtInput.Text
End Property

Public Property Let Text(ByVal value As String)
    m_blnSkipChange = True
    txtInput.Text = value
    m_blnSkipChange = False
End Property

Public Property Get StationName() As String
    StationName = m_strStationName
End Property

Public Property Get StationCode() As String
    StationCode = m_strStationCode
End Property

Public Sub SetStation(ByVal name As String, ByVal code As String)
    m_strStationName = name
    m_strStationCode = code
    m_blnSkipChange = True
    txtInput.Text = name
    m_blnSkipChange = False
    HideSuggestions
End Sub

Private Sub txtInput_Change()
    If m_blnSkipChange Then Exit Sub
    m_strStationName = ""
    m_strStationCode = ""
    If Len(Trim(txtInput.Text)) >= 1 Then
        tmrDebounce.Enabled = False
        tmrDebounce.Enabled = True
    Else
        HideSuggestions
    End If
End Sub

Private Sub tmrDebounce_Timer()
    tmrDebounce.Enabled = False
    If Not StationData.IsLoaded Then
        If Not StationData.IsLoading Then
            StationData.SetLoading
            Set m_asyncLoad = New AsyncRequest
            m_asyncLoad.GetRequest "https://kyfw.12306.cn/otn/personalJS/core/common/station_name_new.js", Me, "OnStationDataLoaded"
        End If
        Exit Sub
    End If
    ShowFilteredStations
End Sub

Public Sub OnStationDataLoaded(ByVal status As Long, ByVal responseText As String)
    If status = 200 Then
        StationData.ParseStationList responseText
        ShowFilteredStations
    End If
    Set m_asyncLoad = Nothing
End Sub

Private Sub ShowFilteredStations()
    Dim keyword As String
    Dim results As Collection
    Dim entry As Object
    Dim i As Long
    Dim lngShow As Long
    
    keyword = Trim(txtInput.Text)
    If Len(keyword) = 0 Then
        HideSuggestions
        Exit Sub
    End If
    
    Set results = StationData.SearchStations(keyword)
    lstSuggest.Clear
    If results.Count = 0 Then
        HideSuggestions
        Exit Sub
    End If
    
    ReDim m_arrSuggestCodes(0 To results.Count - 1)
    For i = 1 To results.Count
        Set entry = results(i)
        lstSuggest.AddItem entry("name")
        m_arrSuggestCodes(i - 1) = entry("code")
    Next
    
    lstSuggest.ListIndex = 0
    If results.Count > 5 Then lngShow = 5 Else lngShow = results.Count
    lstSuggest.Move 0, 360, UserControl.ScaleWidth, lngShow * 375 + 60
    lstSuggest.Visible = True
    
    On Error Resume Next
    UserControl.Height = 375 + lstSuggest.Height + 60
    On Error GoTo 0
End Sub

Private Sub HideSuggestions()
    lstSuggest.Visible = False
    On Error Resume Next
    UserControl.Height = m_normHeight
    On Error GoTo 0
End Sub

Private Sub lstSuggest_Click()
    If lstSuggest.ListIndex < 0 Then Exit Sub
    
    m_strStationName = lstSuggest.List(lstSuggest.ListIndex)
    m_strStationCode = m_arrSuggestCodes(lstSuggest.ListIndex)
    
    m_blnSkipChange = True
    txtInput.Text = m_strStationName
    m_blnSkipChange = False
    
    HideSuggestions
    RaiseEvent StationChanged(m_strStationName, m_strStationCode)
End Sub

Private Sub lstSuggest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lstSuggest_Click
    If KeyCode = vbKeyEscape Then HideSuggestions
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown And lstSuggest.Visible Then
        If lstSuggest.ListCount > 0 Then
            lstSuggest.SetFocus
            lstSuggest.ListIndex = 0
        End If
    End If
    If KeyCode = vbKeyEscape Then HideSuggestions
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If Not m_asyncLoad Is Nothing Then
        m_asyncLoad.Abort
        Set m_asyncLoad = Nothing
    End If
End Sub
