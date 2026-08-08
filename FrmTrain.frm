VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.Form frmTrain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "根据车次查询列车时刻表"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8910
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
   ScaleHeight     =   5565
   ScaleWidth      =   8910
   StartUpPosition =   2  '屏幕中心
   Begin 列车时刻表.TrainNoPicker tpTrainNo 
      Height          =   2295
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
   End
   Begin VBCCR18.ListView lsvResult 
      Height          =   4300
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7594
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
      Left            =   7560
      TabIndex        =   4
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "查询"
   End
   Begin VBCCR18.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5190
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   661
      InitPanels      =   "FrmTrain.frx":0000
   End
   Begin VBCCR18.DTPicker datePicker 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Value           =   32874
      CustomFormat    =   "FrmTrain.frx":04F0
      AllowUserInput  =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "车次："
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
End
Attribute VB_Name = "frmTrain"
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
Private m_asyncSearch As AsyncRequest
Private m_strPQ_TrainCode As String
Private m_strPQ_TrainNoFull As String
Private m_strPQ_Date As String

Private Sub Form_Load()
    datePicker.value = Date
    m_lngSearchCount = 0
    m_lngStationCount = 0

    With lsvResult.ColumnHeaders
        .Add , , "站序", 600
        .Add , , "车站", 1100
        .Add , , "车次", 800
        .Add , , "出发时间", 1000
        .Add , , "到达时间", 1000
        .Add , , "历时", 800
        .Add , , "备注", 1400
    End With
End Sub

Private Sub cmdQuery_Click()
    Dim strTrainCode As String
    Dim strDate As String
    Dim strUrl As String
    
    On Error GoTo ErrorHandler
    
    strTrainCode = Trim(tpTrainNo.Text)
    If Len(strTrainCode) = 0 Then
        MsgBox "请输入车次号", vbExclamation
        tpTrainNo.SetFocus
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
    m_strPQ_TrainCode = strTrainCode
    m_strPQ_TrainNoFull = tpTrainNo.GetTrainNoByCode(strTrainCode)
    m_strPQ_Date = Format(datePicker.value, "yyyy-MM-dd")
    
    If Len(m_strPQ_TrainNoFull) = 0 Then
        statusBar.Panels(1).Text = "正在查询车次编号……"
        strUrl = "https://search.12306.cn/search/v1/train/search?" & _
                 "keyword=" & EncodeURL(strTrainCode) & _
                 "&date=" & m_strPQ_Date
        
        Set m_asyncSearch = New AsyncRequest
        m_asyncSearch.GetRequest strUrl, Me, "OnPhase1Complete"
    Else
        Call StartPhase2
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "查询出错: " & Err.Description, vbCritical
    statusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Public Sub OnPhase1Complete(ByVal status As Long, ByVal responseText As String)
    On Error GoTo ErrH
    If status <> 200 Then
        MsgBox "HTTP请求失败，状态码=" & status, vbCritical
        statusBar.Panels(1).Text = "查询出错"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    ParseTrainSearchResult responseText
    m_strPQ_TrainNoFull = tpTrainNo.GetTrainNoByCode(m_strPQ_TrainCode)
    If Len(m_strPQ_TrainNoFull) = 0 Then
        MsgBox "未找到车次 " & m_strPQ_TrainCode, vbExclamation
        statusBar.Panels(1).Text = "未找到车次"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    Call StartPhase2
    Exit Sub
ErrH:
    MsgBox "查询出错: " & Err.Description, vbCritical
    statusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub StartPhase2()
    Dim strUrl As String
    If Not m_asyncSearch Is Nothing Then
        m_asyncSearch.Abort
        Set m_asyncSearch = Nothing
    End If
    statusBar.Panels(1).Text = "正在查询停靠站信息……"
    strUrl = "https://kyfw.12306.cn/otn/queryTrainInfo/query?leftTicketDTO.train_no=" & EncodeURL(m_strPQ_TrainNoFull) & "&leftTicketDTO.train_date=" & m_strPQ_Date
    Set m_asyncSearch = New AsyncRequest
    m_asyncSearch.GetRequest strUrl, Me, "OnPhase2Complete"
End Sub

Public Sub OnPhase2Complete(ByVal status As Long, ByVal responseText As String)
    On Error GoTo ErrH
    If status <> 200 Then
        MsgBox "HTTP请求失败，状态码=" & status, vbCritical
        statusBar.Panels(1).Text = "查询出错"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    ParseStationInfo responseText, m_strPQ_TrainCode
    ShowResults m_strPQ_TrainCode
    statusBar.Panels(1).Text = "查询完成，全程共有 " & m_lngStationCount & " 个停靠站"
    cmdQuery.Enabled = True
    Set m_asyncSearch = Nothing
    Exit Sub
ErrH:
    MsgBox "查询出错: " & Err.Description, vbCritical
    statusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Public Sub QueryTrainDirectly(ByVal trainCode As String, ByVal trainNo As String)
    m_strPQ_TrainCode = trainCode
    m_strPQ_TrainNoFull = trainNo
    m_strPQ_Date = Format(datePicker.value, "yyyy-MM-dd")
    tpTrainNo.Text = trainCode
    lsvResult.ListItems.Clear
    cmdQuery.Enabled = False
    StartPhase2
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
    m_lngSearchCount = 0
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
    m_lngStationCount = 0
End Sub

Private Sub ShowResults(ByVal strTrainCode As String)
    Dim i As Long
    Dim Item As Object
    Dim strLishi As String
    
    lsvResult.ListItems.Clear
    For i = 0 To m_lngStationCount - 1
        Set Item = lsvResult.ListItems.Add(, , m_arrStations(i).station_no)
        Item.SubItems(1) = m_arrStations(i).station_name
        Item.SubItems(2) = strTrainCode
        Item.SubItems(3) = m_arrStations(i).start_time
        Item.SubItems(4) = m_arrStations(i).arrive_time
        If i = 0 Then
            strLishi = "----"
        ElseIf Len(Trim(m_arrStations(i).running_time)) > 0 Then
            strLishi = m_arrStations(i).running_time
        Else
            strLishi = "----"
        End If
        Item.SubItems(5) = strLishi
        Item.SubItems(6) = m_arrStations(i).arrive_day_str
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_asyncSearch Is Nothing Then
        m_asyncSearch.Abort
        Set m_asyncSearch = Nothing
    End If
End Sub

