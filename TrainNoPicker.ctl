VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.UserControl TrainNoPicker 
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   1575
   ScaleWidth      =   3000
   Begin VBCCR18.ListBoxW lstSuggest 
      Height          =   825
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1455
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
Attribute VB_Name = "TrainNoPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event TrainSelected(ByVal trainCode As String, ByVal trainNo As String)

Private Type TrainSearchResult
    train_no As String
    station_train_code As String
    from_station_name As String
    to_station_name As String
    start_time As String
    arrive_time As String
End Type

Private m_asyncSearch As AsyncRequest
Private m_arrResults() As TrainSearchResult
Private m_lngResultCount As Long
Private m_strTrainCode As String
Private m_strTrainNo As String
Private m_blnSkipChange As Boolean
Private m_strPendingKeyword As String
Private m_normHeight As Long

Private Sub UserControl_Initialize()
    m_normHeight = 375
    m_lngResultCount = 0
    
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

Public Property Get trainCode() As String
    trainCode = m_strTrainCode
End Property

Public Property Get trainNo() As String
    trainNo = m_strTrainNo
End Property

Public Sub SetTrainCode(ByVal trainCode As String, ByVal trainNo As String)
    m_strTrainCode = trainCode
    m_strTrainNo = trainNo
    m_blnSkipChange = True
    txtInput.Text = trainCode
    m_blnSkipChange = False
    HideSuggestions
End Sub

Private Sub txtInput_Change()
    If m_blnSkipChange Then Exit Sub
    If Len(Trim(txtInput.Text)) >= 1 Then
        tmrDebounce.Enabled = False
        tmrDebounce.Enabled = True
    Else
        HideSuggestions
    End If
End Sub

Private Sub tmrDebounce_Timer()
    tmrDebounce.Enabled = False
    SearchTrainSuggest
End Sub

Private Sub SearchTrainSuggest()
    Dim strKeyword As String
    Dim strUrl As String
    
    On Error GoTo ErrorHandler
    strKeyword = Trim(txtInput.Text)
    If Len(strKeyword) = 0 Then
        HideSuggestions
        Exit Sub
    End If
    
    On Error Resume Next
    If Not m_asyncSearch Is Nothing Then m_asyncSearch.Abort
    Set m_asyncSearch = Nothing
    On Error GoTo ErrorHandler
    
    m_strPendingKeyword = strKeyword
    strUrl = "https://search.12306.cn/search/v1/train/search?keyword=" & EncodeURL(strKeyword) & "&date=" & Format(Date, "yyyyMMdd")
    Set m_asyncSearch = New AsyncRequest
    m_asyncSearch.GetRequest strUrl, Me, "OnSuggestComplete"
    Exit Sub
    
ErrorHandler:
    HideSuggestions
End Sub

Public Sub OnSuggestComplete(ByVal status As Long, ByVal responseText As String)
    Dim i As Long
    Dim lngShow As Long
    Dim lngSelStart As Long
    Dim lngSelLen As Long
    
    On Error GoTo ErrHandler
    If Trim(txtInput.Text) <> m_strPendingKeyword Then Exit Sub
    If status <> 200 Then
        HideSuggestions
        Exit Sub
    End If
    
    ParseTrainSearchResult responseText
    
    lngSelStart = txtInput.SelStart
    lngSelLen = txtInput.SelLength
    lstSuggest.Clear
    
    If m_lngResultCount > 0 Then
        For i = 0 To m_lngResultCount - 1
            lstSuggest.AddItem m_arrResults(i).station_train_code
        Next i
        lstSuggest.ListIndex = 0
        If lstSuggest.ListCount > 5 Then lngShow = 5 Else lngShow = lstSuggest.ListCount
        lstSuggest.Move 0, 360, UserControl.ScaleWidth, lngShow * 375 + 60
        lstSuggest.Visible = True
        txtInput.SetFocus
        txtInput.SelStart = lngSelStart
        txtInput.SelLength = lngSelLen
    Else
        HideSuggestions
    End If
    Exit Sub
    
ErrHandler:
    HideSuggestions
End Sub

Private Sub ParseTrainSearchResult(ByVal strJson As String)
    Dim parser As JSON
    Dim root As Object
    Dim dataArr As Object
    Dim lngCount As Long
    Dim i As Long
    Dim itemObj As Object
    
    On Error GoTo ErrHandler
    m_lngResultCount = 0
    
    Set parser = New JSON
    Set root = parser.Parse(strJson)
    Set parser = Nothing
    If root Is Nothing Then Exit Sub
    
    If Not root.Exists("data") Then Exit Sub
    Set dataArr = root("data")
    If TypeName(dataArr) <> "Collection" Then Exit Sub
    lngCount = dataArr.Count
    If lngCount = 0 Then Exit Sub
    
    ReDim m_arrResults(0 To lngCount - 1)
    m_lngResultCount = lngCount
    
    For i = 1 To lngCount
        Set itemObj = dataArr(i)
        If TypeName(itemObj) = "Dictionary" Then
            With m_arrResults(i - 1)
                If itemObj.Exists("train_no") Then .train_no = CStr(itemObj("train_no")) Else .train_no = ""
                If itemObj.Exists("station_train_code") Then .station_train_code = CStr(itemObj("station_train_code")) Else .station_train_code = ""
                If itemObj.Exists("from_station") Then .from_station_name = CStr(itemObj("from_station")) Else .from_station_name = ""
                If itemObj.Exists("to_station") Then .to_station_name = CStr(itemObj("to_station")) Else .to_station_name = ""
                If itemObj.Exists("start_time") Then .start_time = CStr(itemObj("start_time")) Else .start_time = ""
                If itemObj.Exists("arrive_time") Then .arrive_time = CStr(itemObj("arrive_time")) Else .arrive_time = ""
            End With
        End If
    Next i
    Exit Sub
    
ErrHandler:
    m_lngResultCount = 0
End Sub

Public Function GetTrainNoByCode(ByVal trainCode As String) As String
    Dim i As Long
    For i = 0 To m_lngResultCount - 1
        If StrComp(m_arrResults(i).station_train_code, trainCode, vbTextCompare) = 0 Then
            GetTrainNoByCode = m_arrResults(i).train_no
            Exit Function
        End If
    Next
    GetTrainNoByCode = ""
End Function

Private Sub lstSuggest_Click()
    Dim strText As String
    Dim i As Long
    
    If lstSuggest.ListIndex < 0 Then Exit Sub
    strText = lstSuggest.List(lstSuggest.ListIndex)
    
    m_strTrainCode = strText
    m_strTrainNo = ""
    For i = 0 To m_lngResultCount - 1
        If StrComp(m_arrResults(i).station_train_code, strText, vbTextCompare) = 0 Then
            m_strTrainNo = m_arrResults(i).train_no
            Exit For
        End If
    Next
    
    m_blnSkipChange = True
    txtInput.Text = strText
    m_blnSkipChange = False
    
    HideSuggestions
    RaiseEvent TrainSelected(m_strTrainCode, m_strTrainNo)
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

Private Sub HideSuggestions()
    lstSuggest.Visible = False
    On Error Resume Next
    UserControl.Height = m_normHeight
    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If Not m_asyncSearch Is Nothing Then
        m_asyncSearch.Abort
        Set m_asyncSearch = Nothing
    End If
End Sub
