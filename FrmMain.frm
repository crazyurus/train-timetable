VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "列车时刻表"
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
   Begin VBCCR18.ListView lvResult 
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7435
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      LabelWrap       =   0   'False
      HighlightColumnHeaders=   -1  'True
      TrackSizeColumnHeaders=   0   'False
      AutoSelectFirstItem=   0   'False
   End
   Begin VBCCR18.CommandButtonW cmdQuery 
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   241
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "查询"
   End
   Begin VBCCR18.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5055
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   661
      InitPanels      =   "FrmMain.frx":0000
   End
   Begin VBCCR18.DTPicker dtpDate 
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Value           =   32874
      CustomFormat    =   "FrmMain.frx":04F0
      AllowUserInput  =   -1  'True
   End
   Begin VB.TextBox txtTrainNo 
      Height          =   377
      Left            =   4560
      TabIndex        =   1
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
   Begin VB.Timer tmrDebounce 
      Interval        =   300
      Left            =   7320
      Top             =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "车次："
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   255
      Left            =   240
      TabIndex        =   4
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
Private m_asyncSearch As AsyncRequest
Private m_strPQ_TrainCode As String
Private m_strPQ_TrainNoFull As String
Private m_strPQ_Date As String
Private m_strPQ_DateDash As String
Private Sub Form_Load()
    dtpDate.value = Date
    m_lngSearchCount = 0
    m_lngStationCount = 0

    With lvResult.ColumnHeaders
        .Add , , "站序", 600
        .Add , , "车站", 1800
        .Add , , "车次", 800
        .Add , , "出发时间", 1000
        .Add , , "到达时间", 1000
        .Add , , "历时", 800
        .Add , , "备注", 1400
    End With
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
    strDate = Format(dtpDate.value, "yyyyMMdd")
    StatusBar.Panels(1).Text = "正在搜索车次..."
    strUrl = "https://search.12306.cn/search/v1/train/search?keyword=" & EncodeURL(strKeyword) & "&date=" & strDate
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
    Dim parser As JSON
    Dim root As Object
    Dim dataArr As Object
    Dim lngCount As Long
    Dim i As Long
    Dim itemObj As Object
    On Error GoTo ErrHandler
    m_lngSearchCount = 0
    Set parser = New JSON
    Set root = parser.Parse(strJson)
    Set parser = Nothing
    If root Is Nothing Then Exit Sub
    ' search.12306.cn returns: { data: [ {train_no, station_train_code, from_station, to_station}, ... ] }
    If Not root.Exists("data") Then Exit Sub
    Set dataArr = root("data")
    If TypeName(dataArr) <> "Collection" Then Exit Sub
    lngCount = dataArr.Count
    If lngCount = 0 Then Exit Sub
    ReDim m_arrTrainSearch(0 To lngCount - 1)
    m_lngSearchCount = lngCount
    For i = 1 To lngCount
        Set itemObj = dataArr(i)
        If TypeName(itemObj) = "Dictionary" Then
            With m_arrTrainSearch(i - 1)
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
    StatusBar.Panels(1).Text = "车次搜索失败: " & Err.Description
    m_lngSearchCount = 0
End Sub

Private Sub cmdQuery_Click()
    Dim strTrainCode As String
    Dim strDate As String
    Dim strDateDash As String
    Dim strUrl As String
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
    If Not m_asyncSearch Is Nothing Then
        m_asyncSearch.Abort
        Set m_asyncSearch = Nothing
    End If
    cmdQuery.Enabled = False
    lstSuggest.Visible = False
    m_strPQ_TrainCode = strTrainCode
    m_strPQ_TrainNoFull = ""
    m_strPQ_Date = Format(dtpDate.value, "yyyyMMdd")
    m_strPQ_DateDash = Format(dtpDate.value, "yyyy-MM-dd")
    If m_lngSearchCount = 0 Or _
       (m_lngSearchCount > 0 And _
        StrComp(m_arrTrainSearch(0).station_train_code, strTrainCode, vbTextCompare) <> 0) Then
        StatusBar.Panels(1).Text = "正在查询车次编号……"
        strUrl = "https://search.12306.cn/search/v1/train/search?keyword=" & EncodeURL(strTrainCode) & "&date=" & m_strPQ_Date
        Set m_asyncSearch = New AsyncRequest
        m_asyncSearch.GetRequest strUrl, Me, "OnPhase1Complete"
    Else
        Call DoPhase1Local
    End If
    Exit Sub
ErrorHandler:
    MsgBox "查询出错: " & Err.Description, vbCritical
    StatusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub DoPhase1Local()
    Dim i As Long
    Dim blnFound As Boolean
    blnFound = False
    For i = 0 To m_lngSearchCount - 1
        If StrComp(m_arrTrainSearch(i).station_train_code, m_strPQ_TrainCode, vbTextCompare) = 0 Then
            m_strPQ_TrainNoFull = m_arrTrainSearch(i).train_no
            blnFound = True
            Exit For
        End If
    Next i
    If Not blnFound Then
        MsgBox "未找到车次 " & m_strPQ_TrainCode, vbExclamation
        StatusBar.Panels(1).Text = "未找到车次"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    Call StartPhase2
End Sub

Public Sub OnPhase1Complete(ByVal status As Long, ByVal responseText As String)
    On Error GoTo ErrH
    If status <> 200 Then
        MsgBox "HTTP请求失败，状态码=" & status, vbCritical
        StatusBar.Panels(1).Text = "查询出错"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    ParseTrainSearchResult responseText
    Call DoPhase1Local
    Exit Sub
ErrH:
    MsgBox "查询出错: " & Err.Description, vbCritical
    StatusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub StartPhase2()
    Dim strUrl As String
    If Not m_asyncSearch Is Nothing Then
        m_asyncSearch.Abort
        Set m_asyncSearch = Nothing
    End If
    StatusBar.Panels(1).Text = "正在查询停靠站信息……"
    strUrl = "https://kyfw.12306.cn/otn/queryTrainInfo/query?" & "leftTicketDTO.train_no=" & EncodeURL(m_strPQ_TrainNoFull) & "&leftTicketDTO.train_date=" & m_strPQ_DateDash
    Set m_asyncSearch = New AsyncRequest
    m_asyncSearch.GetRequest strUrl, Me, "OnPhase2Complete"
End Sub

Public Sub OnPhase2Complete(ByVal status As Long, ByVal responseText As String)
    On Error GoTo ErrH
    If status <> 200 Then
        MsgBox "HTTP请求失败，状态码=" & status, vbCritical
        StatusBar.Panels(1).Text = "查询出错"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    ParseStationInfo responseText, m_strPQ_TrainCode
    ShowResults m_strPQ_TrainCode
    StatusBar.Panels(1).Text = "查询完成，全程共有 " & m_lngStationCount & " 个停靠站"
    cmdQuery.Enabled = True
    Set m_asyncSearch = Nothing
    Exit Sub
ErrH:
    MsgBox "查询出错: " & Err.Description, vbCritical
    StatusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub ParseStationInfo(ByVal strJson As String, ByVal strTrainCode As String)
    Dim parser As JSON
    Dim root As Object
    Dim dataObj As Object
    Dim dataArr As Object
    Dim lngCount As Long
    Dim i As Long
    Dim itemObj As Object
    On Error GoTo ErrHandler
    m_lngStationCount = 0
    Set parser = New JSON
    Set root = parser.Parse(strJson)
    Set parser = Nothing
    If root Is Nothing Then Exit Sub
    ' kyfw.12306.cn returns: { data: { data: [ {station_no, station_name, arrive_time, ...}, ... ] } }
    If Not root.Exists("data") Then Exit Sub
    Set dataObj = root("data")
    If dataObj Is Nothing Then Exit Sub
    If Not dataObj.Exists("data") Then Exit Sub
    Set dataArr = dataObj("data")
    If TypeName(dataArr) <> "Collection" Then Exit Sub
    lngCount = dataArr.Count
    If lngCount = 0 Then Exit Sub
    ReDim m_arrStations(0 To lngCount - 1)
    m_lngStationCount = lngCount
    For i = 1 To lngCount
        Set itemObj = dataArr(i)
        If TypeName(itemObj) = "Dictionary" Then
            With m_arrStations(i - 1)
                If itemObj.Exists("station_no") Then .station_no = CStr(itemObj("station_no")) Else .station_no = ""
                If itemObj.Exists("station_name") Then .station_name = CStr(itemObj("station_name")) Else .station_name = ""
                If itemObj.Exists("arrive_time") Then .arrive_time = CStr(itemObj("arrive_time")) Else .arrive_time = ""
                If itemObj.Exists("start_time") Then .start_time = CStr(itemObj("start_time")) Else .start_time = ""
                .stopover_time = ""
                If itemObj.Exists("arrive_day_str") Then .arrive_day_str = CStr(itemObj("arrive_day_str")) Else .arrive_day_str = ""
                If itemObj.Exists("running_time") Then .running_time = CStr(itemObj("running_time")) Else .running_time = ""
            End With
        End If
    Next i
    Exit Sub
ErrHandler:
    StatusBar.Panels(1).Text = "站点解析失败: " & Err.Description
    m_lngStationCount = 0
End Sub

Private Sub ShowResults(ByVal strTrainCode As String)
    Dim i As Long
    Dim item As Object
    Dim strLishi As String
    lvResult.ListItems.Clear
    For i = 0 To m_lngStationCount - 1
        Set item = lvResult.ListItems.Add(, , m_arrStations(i).station_no)
        item.SubItems(1) = m_arrStations(i).station_name
        item.SubItems(2) = strTrainCode
        item.SubItems(3) = m_arrStations(i).start_time
        item.SubItems(4) = m_arrStations(i).arrive_time
        If i = 0 Then
            strLishi = "----"
        ElseIf Len(Trim(m_arrStations(i).running_time)) > 0 Then
            strLishi = m_arrStations(i).running_time
        Else
            strLishi = "----"
        End If
        item.SubItems(5) = strLishi
        item.SubItems(6) = m_arrStations(i).arrive_day_str
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_asyncSearch Is Nothing Then
        m_asyncSearch.Abort
        Set m_asyncSearch = Nothing
    End If
    If Not m_asyncSuggest Is Nothing Then
        m_asyncSuggest.Abort
        Set m_asyncSuggest = Nothing
    End If
End Sub
