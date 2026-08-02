VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "列车时刻表查询"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleMode       =   0  'User
   ScaleWidth      =   9030
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   5055
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
            Text            =   "制造本程序：Cr4zy Uru5"
            TextSave        =   "制造本程序：Cr4zy Uru5"
            Object.ToolTipText     =   "Cr4zy Uru5"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2205
            MinWidth        =   2205
            Text            =   "No.2026001"
            TextSave        =   "No.2026001"
            Object.ToolTipText     =   "2026年第一个程序"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2734
            MinWidth        =   2734
            TextSave        =   "2026/8/2"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "22:05"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTrainNo 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.ListBox lstSuggest 
      Height          =   1590
      Left            =   4560
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   4215
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "站序"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车站"
         Object.Width           =   2645
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "出发时间"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "到达时间"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "历时"
         Object.Width           =   2645
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "备注"
         Object.Width           =   1763
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   149094401
      CurrentDate     =   36494
   End
   Begin VB.Timer tmrDebounce 
      Interval        =   300
      Left            =   7440
      Top             =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "车次："
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   300
      Width           =   540
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TrainSearchResult
    train_no As String
    station_train_code As String
    from_station_name As String
    to_station_name As String
    start_time As String
    arrive_time As String
End Type

Private m_arrTrainSearch() As TrainSearchResult
Private m_lngSearchCount As Long

Private Type StationInfo
    station_no As String
    station_name As String
    arrive_time As String
    start_time As String
    stopover_time As String
    arrive_day_str As String
    running_time As String
End Type

Private m_arrStations() As StationInfo
Private m_lngStationCount As Long
Private m_blnSkipChange As Boolean
Private m_asyncSuggest As AsyncRequest
Private m_strPendingKeyword As String

Private Sub Form_Load()
    dtpDate.Value = Date
    m_lngSearchCount = 0
    m_lngStationCount = 0
End Sub

Private Sub txtTrainNo_Change()
    If m_blnSkipChange = True Then Exit Sub
    If Len(Trim(txtTrainNo.Text)) >= 1 Then
        tmrDebounce.Enabled = False
        tmrDebounce.Enabled = True
    Else
        tmrDebounce.Enabled = False
        lstSuggest.Visible = False
    End If
End Sub

Private Sub tmrDebounce_Timer()
    tmrDebounce.Enabled = False
    SearchTrainSuggest
End Sub

Private Sub lstSuggest_Click()
    If lstSuggest.ListIndex >= 0 Then
        m_blnSkipChange = True
        txtTrainNo.Text = lstSuggest.List(lstSuggest.ListIndex)
        m_blnSkipChange = False
        lstSuggest.Visible = False
    End If
End Sub

Private Sub SearchTrainSuggest()
    Dim strKeyword As String
    Dim strDate As String
    Dim strUrl As String
    On Error GoTo ErrorHandler
    strKeyword = Trim(txtTrainNo.Text)
    If Len(strKeyword) = 0 Then
        lstSuggest.Visible = False
        Exit Sub
    End If
    On Error Resume Next
    If Not m_asyncSuggest Is Nothing Then
        m_asyncSuggest.Abort
    End If
    Set m_asyncSuggest = Nothing
    On Error GoTo ErrorHandler
    m_strPendingKeyword = strKeyword
    strDate = Format(dtpDate.Value, "yyyyMMdd")
    StatusBar.Panels(1).Text = "正在搜索车次..."
    strUrl = "https://search.12306.cn/search/v1/train/search?keyword=" & _
             EncodeURL(strKeyword) & "&date=" & strDate
    Set m_asyncSuggest = New AsyncRequest
    m_asyncSuggest.GetRequest strUrl, Me, "OnSuggestComplete"
    Exit Sub
ErrorHandler:
    StatusBar.Panels(1).Text = "搜索出错: " & Err.Description
    lstSuggest.Visible = False
End Sub

Public Sub OnSuggestComplete(ByVal status As Long, ByVal responseText As String)
    Dim i As Long
    Dim lngShow As Long
    Dim lngSelStart As Long
    Dim lngSelLen As Long
    On Error GoTo ErrHandler
    If Trim(txtTrainNo.Text) <> m_strPendingKeyword Then Exit Sub
    If status <> 200 Then
        StatusBar.Panels(1).Text = "搜索出错: HTTP status " & status
        lstSuggest.Visible = False
        Exit Sub
    End If
    ParseTrainSearchResult responseText
    lngSelStart = txtTrainNo.SelStart
    lngSelLen = txtTrainNo.SelLength
    lstSuggest.Clear
    If m_lngSearchCount > 0 Then
        For i = 0 To m_lngSearchCount - 1
            lstSuggest.AddItem m_arrTrainSearch(i).station_train_code
        Next i
        lstSuggest.ListIndex = 0
        If lstSuggest.ListCount > 5 Then
            lngShow = 5
        Else
            lngShow = lstSuggest.ListCount
        End If
        lstSuggest.Height = lngShow * 315 + 60
        lstSuggest.Visible = True
        txtTrainNo.SetFocus
        txtTrainNo.SelStart = lngSelStart
        txtTrainNo.SelLength = lngSelLen
    Else
        lstSuggest.Visible = False
    End If
    StatusBar.Panels(1).Text = "找到 " & m_lngSearchCount & " 个车次"
    Exit Sub
ErrHandler:
    StatusBar.Panels(1).Text = "搜索出错: " & Err.Description
    lstSuggest.Visible = False
End Sub

Private Sub ParseTrainSearchResult(ByVal strJson As String)
    Dim strData As String
    Dim strItems() As String
    Dim lngCount As Long
    Dim i As Long
    Dim strItem As String
    m_lngSearchCount = 0
    If InStr(strJson, """data"":") = 0 Then Exit Sub
    strData = Mid(strJson, InStr(strJson, """data"":") + 7)
    strData = Trim(strData)
    If InStr(strData, "[") > 0 Then
        strData = Mid(strData, InStr(strData, "["))
        strData = GetJsonArray(strData)
    End If
    If Len(strData) = 0 Then Exit Sub
    strItems = SplitJsonObject(strData)
    lngCount = UBound(strItems) + 1
    If lngCount = 0 Then Exit Sub
    ReDim m_arrTrainSearch(0 To lngCount - 1)
    m_lngSearchCount = lngCount
    For i = 0 To lngCount - 1
        strItem = strItems(i)
        With m_arrTrainSearch(i)
            .train_no = GetJsonValue(strItem, "train_no")
            .station_train_code = GetJsonValue(strItem, "station_train_code")
            .from_station_name = GetJsonValue(strItem, "from_station_name")
            .to_station_name = GetJsonValue(strItem, "to_station_name")
            .start_time = GetJsonValue(strItem, "start_time")
            .arrive_time = GetJsonValue(strItem, "arrive_time")
        End With
    Next i
End Sub

Private Sub cmdQuery_Click()
    Dim strTrainCode As String
    Dim strDate As String
    Dim strDateDash As String
    Dim strUrl As String
    Dim strResponse As String
    Dim i As Long
    Dim blnFound As Boolean
    Dim strTrainNoFull As String
    On Error GoTo ErrorHandler
    strTrainCode = Trim(txtTrainNo.Text)
    If Len(strTrainCode) = 0 Then
        MsgBox "请输入车次号", vbExclamation
        txtTrainNo.SetFocus
        Exit Sub
    End If
    If InStr(strTrainCode, " ") > 0 Then
        strTrainCode = Left(strTrainCode, InStr(strTrainCode, " ") - 1)
    End If
    StatusBar.Panels(1).Text = "正在查询列车时刻表..."
    cmdQuery.Enabled = False
    lstSuggest.Visible = False
    strDate = Format(dtpDate.Value, "yyyyMMdd")
    strDateDash = Format(dtpDate.Value, "yyyy-MM-dd")
    If m_lngSearchCount = 0 Or _
       (m_lngSearchCount > 0 And _
        StrComp(m_arrTrainSearch(0).station_train_code, strTrainCode, vbTextCompare) <> 0) Then
        strUrl = "https://search.12306.cn/search/v1/train/search?keyword=" & _
                 EncodeURL(strTrainCode) & "&date=" & strDate
        strResponse = HttpGet(strUrl)
        ParseTrainSearchResult strResponse
    End If
    blnFound = False
    strTrainNoFull = ""
    For i = 0 To m_lngSearchCount - 1
        If StrComp(m_arrTrainSearch(i).station_train_code, strTrainCode, vbTextCompare) = 0 Then
            strTrainNoFull = m_arrTrainSearch(i).train_no
            blnFound = True
            Exit For
        End If
    Next i
    If Not blnFound Then
        MsgBox "未找到车次 " & strTrainCode, vbExclamation
        StatusBar.Panels(1).Text = "未找到车次"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    strUrl = "https://kyfw.12306.cn/otn/queryTrainInfo/query?" & _
             "leftTicketDTO.train_no=" & EncodeURL(strTrainNoFull) & _
             "&leftTicketDTO.train_date=" & strDateDash & _
             "&rand_code="
    strResponse = HttpGet(strUrl)
    ParseStationInfo strResponse, strTrainCode
    ShowResults strTrainCode
    StatusBar.Panels(1).Text = "查询完成，全程共有 " & m_lngStationCount & " 个停靠站"
    cmdQuery.Enabled = True
    Exit Sub
ErrorHandler:
    MsgBox "查询出错: " & Err.Description, vbCritical
    StatusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub ParseStationInfo(ByVal strJson As String, ByVal strTrainCode As String)
    Dim strData As String
    Dim strItems() As String
    Dim lngCount As Long
    Dim i As Long
    Dim strItem As String
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    m_lngStationCount = 0
    lngPos1 = InStr(strJson, """data"":")
    If lngPos1 = 0 Then Exit Sub
    lngPos2 = InStr(lngPos1 + 10, strJson, """data"":")
    If lngPos2 = 0 Then
        strData = Mid(strJson, lngPos1 + 7)
    Else
        strData = Mid(strJson, lngPos2 + 7)
    End If
    strData = Trim(strData)
    If InStr(strData, "[") > 0 Then
        strData = Mid(strData, InStr(strData, "["))
        strData = GetJsonArray(strData)
    End If
    If Len(strData) = 0 Then Exit Sub
    strItems = SplitJsonObject(strData)
    lngCount = UBound(strItems) + 1
    If lngCount = 0 Then Exit Sub
    ReDim m_arrStations(0 To lngCount - 1)
    m_lngStationCount = lngCount
    For i = 0 To lngCount - 1
        strItem = strItems(i)
        With m_arrStations(i)
            .station_no = GetJsonValue(strItem, "station_no")
            .station_name = GetJsonValue(strItem, "station_name")
            .arrive_time = GetJsonValue(strItem, "arrive_time")
            .start_time = GetJsonValue(strItem, "start_time")
            .stopover_time = GetJsonValue(strItem, "stopover_time")
            .arrive_day_str = GetJsonValue(strItem, "arrive_day_str")
            .running_time = GetJsonValue(strItem, "running_time")
        End With
    Next i
End Sub

Private Sub ShowResults(ByVal strTrainCode As String)
    Dim i As Long
    Dim itm As ListItem
    Dim strLishi As String
    lvResult.ListItems.Clear
    For i = 0 To m_lngStationCount - 1
        Set itm = lvResult.ListItems.Add(, , m_arrStations(i).station_no)
        itm.SubItems(1) = m_arrStations(i).station_name
        itm.SubItems(2) = strTrainCode
        itm.SubItems(3) = m_arrStations(i).start_time
        itm.SubItems(4) = m_arrStations(i).arrive_time
        If i = 0 Then
            strLishi = "----"
        ElseIf Len(Trim(m_arrStations(i).running_time)) > 0 Then
            strLishi = m_arrStations(i).running_time
        Else
            strLishi = "----"
        End If
        itm.SubItems(5) = strLishi
        itm.SubItems(6) = m_arrStations(i).arrive_day_str
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_asyncSuggest Is Nothing Then
        m_asyncSuggest.Abort
        Set m_asyncSuggest = Nothing
    End If
End Sub
