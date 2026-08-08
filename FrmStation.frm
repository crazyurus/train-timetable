VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.Form frmStation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "根据车站查询列车时刻表"
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
   Begin 列车时刻表.StationPicker spStation 
      Height          =   2295
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
   End
   Begin VBCCR18.CommandButtonW cmdQuery 
      Default         =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   0
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
      InitPanels      =   "FrmStation.frx":0000
   End
   Begin VBCCR18.DTPicker datePicker 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Value           =   32874
      CustomFormat    =   "FrmStation.frx":04F0
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.ListView lsvResult 
      Height          =   4300
      Left            =   240
      TabIndex        =   3
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "车站："
      Height          =   255
      Left            =   3840
      TabIndex        =   4
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
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type StationResult
    station_train_code As String
    train_class_name As String
    train_type_name As String
    start_station_name As String
    end_station_name As String
    start_time As String
    arrive_time As String
    running_time As String
    start_day_diff As String
    arrive_day_diff As String
    platform_no As String
    update_start_time As String
    update_arrive_time As String
    jiaolu_corporation_code As String
    jiaolu_train_style As String
    running_fig As String
    distance As String
    end_arrive_time As String
    speed As String
    stopover_time As String
    jiaolu_dept_train As String
    train_no As String
    station_no As String
End Type

Private m_arrResults() As StationResult
Private m_lngResultCount As Long
Private m_asyncQuery As AsyncRequest

Private Sub Form_Load()
    datePicker.value = Date
    m_lngResultCount = 0
    
    With lsvResult.ColumnHeaders
        .Add , , "车次", 800
        .Add , , "类型", 600
        .Add , , "到站时间", 900
        .Add , , "出站时间", 900
        .Add , , "停留时长", 900
        .Add , , "正晚点到达", 1100
        .Add , , "正晚点出发", 1100
        .Add , , "始发站", 1100
        .Add , , "终到站", 1100
        .Add , , "终到时间", 900
        .Add , , "历时", 1200
        .Add , , "距离", 1000
        .Add , , "站台", 1600
        .Add , , "担当路局", 1400
        .Add , , "车辆段", 1800
        .Add , , "车底类型", 1400
        .Add , , "速度", 1200
    End With
End Sub

Private Sub cmdQuery_Click()
    Dim strStationCode As String
    Dim strDate As String
    Dim strBody As String
    Dim headers As Object
    
    On Error GoTo ErrorHandler
    
    strStationCode = spStation.StationCode
    If Len(strStationCode) = 0 Then
        MsgBox "请选择车站", vbExclamation
        spStation.SetFocus
        Exit Sub
    End If
    
    If Not m_asyncQuery Is Nothing Then
        m_asyncQuery.Abort
        Set m_asyncQuery = Nothing
    End If
    
    cmdQuery.Enabled = False
    statusBar.Panels(1).Text = "正在查询列车信息……"
    lsvResult.ListItems.Clear
    
    strDate = Format(datePicker.value, "yyyy-mm-dd")
    strBody = "train_station_code=" & strStationCode & "&train_start_date=" & strDate
    
    Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "Referer", "https://servicewechat.com/"
    
    Set m_asyncQuery = New AsyncRequest
    m_asyncQuery.PostRequest "https://mobile.12306.cn/wxxcx/wechat/bigScreen/queryTrainByStation", strBody, Me, "OnQueryComplete", headers
    Exit Sub
    
ErrorHandler:
    MsgBox "查询出错: " & Err.Description, vbCritical
    statusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Public Sub OnQueryComplete(ByVal status As Long, ByVal responseText As String)
    On Error GoTo ErrHandler
    
    If status <> 200 Then
        MsgBox "HTTP请求失败，状态码=" & status, vbCritical
        statusBar.Panels(1).Text = "查询出错"
        cmdQuery.Enabled = True
        Exit Sub
    End If
    
    ParseStationResult responseText
    ShowResults
    statusBar.Panels(1).Text = "查询完成，共计 " & m_lngResultCount & " 个车次"
    cmdQuery.Enabled = True
    Set m_asyncQuery = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "查询出错: " & Err.Description, vbCritical
    statusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub ParseStationResult(ByVal strJson As String)
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
                If itemObj.Exists("station_train_code") Then .station_train_code = CStr(itemObj("station_train_code")) Else .station_train_code = ""
                If itemObj.Exists("train_class_name") Then .train_class_name = CStr(itemObj("train_class_name")) Else .train_class_name = ""
                If itemObj.Exists("train_type_name") Then .train_type_name = CStr(itemObj("train_type_name")) Else .train_type_name = ""
                If itemObj.Exists("start_station_name") Then .start_station_name = CStr(itemObj("start_station_name")) Else .start_station_name = ""
                If itemObj.Exists("end_station_name") Then .end_station_name = CStr(itemObj("end_station_name")) Else .end_station_name = ""
                If itemObj.Exists("start_time") Then .start_time = CStr(itemObj("start_time")) Else .start_time = ""
                If itemObj.Exists("arrive_time") Then .arrive_time = CStr(itemObj("arrive_time")) Else .arrive_time = ""
                If itemObj.Exists("running_time") Then .running_time = CStr(itemObj("running_time")) Else .running_time = ""
                If itemObj.Exists("start_day_diff") Then .start_day_diff = CStr(itemObj("start_day_diff")) Else .start_day_diff = ""
                If itemObj.Exists("arrive_day_diff") Then .arrive_day_diff = CStr(itemObj("arrive_day_diff")) Else .arrive_day_diff = ""
                If itemObj.Exists("platform_no") Then .platform_no = CStr(itemObj("platform_no")) Else .platform_no = ""
                If itemObj.Exists("update_start_time") Then .update_start_time = FormatUpdateTime(CStr(itemObj("update_start_time"))) Else .update_start_time = ""
                If itemObj.Exists("update_arrive_time") Then .update_arrive_time = FormatUpdateTime(CStr(itemObj("update_arrive_time"))) Else .update_arrive_time = ""
                If itemObj.Exists("jiaolu_corporation_code") Then .jiaolu_corporation_code = CStr(itemObj("jiaolu_corporation_code")) Else .jiaolu_corporation_code = ""
                If itemObj.Exists("jiaolu_train_style") Then .jiaolu_train_style = CStr(itemObj("jiaolu_train_style")) Else .jiaolu_train_style = ""
                If itemObj.Exists("running_fig") Then .running_fig = CStr(itemObj("running_fig")) Else .running_fig = ""
                If itemObj.Exists("distance") Then .distance = CStr(itemObj("distance")) Else .distance = ""
                If itemObj.Exists("end_arrive_time") Then .end_arrive_time = CStr(itemObj("end_arrive_time")) Else .end_arrive_time = ""
                If itemObj.Exists("speed") Then .speed = CStr(itemObj("speed")) Else .speed = ""
                If itemObj.Exists("stopover_time") Then .stopover_time = CStr(itemObj("stopover_time")) Else .stopover_time = ""
                If itemObj.Exists("jiaolu_dept_train") Then .jiaolu_dept_train = CStr(itemObj("jiaolu_dept_train")) Else .jiaolu_dept_train = ""
                If itemObj.Exists("train_no") Then .train_no = CStr(itemObj("train_no")) Else .train_no = ""
                If itemObj.Exists("station_no") Then .station_no = CStr(itemObj("station_no")) Else .station_no = ""
            End With
        End If
    Next i
    Exit Sub
    
ErrHandler:
    m_lngResultCount = 0
End Sub

Private Function FormatUpdateTime(ByVal timeStr As String) As String
    If Len(timeStr) = 4 Then
        FormatUpdateTime = Left$(timeStr, 2) & ":" & Right$(timeStr, 2)
    ElseIf Len(timeStr) = 0 Or timeStr = "----" Then
        FormatUpdateTime = ""
    Else
        FormatUpdateTime = timeStr
    End If
End Function

Private Sub ShowResults()
    Dim i As Long
    Dim Item As Object
    Dim strStartTime As String
    Dim strArriveTime As String
    
    lsvResult.ListItems.Clear
    For i = 0 To m_lngResultCount - 1
        Set Item = lsvResult.ListItems.Add(, , m_arrResults(i).station_train_code)
        Item.SubItems(1) = m_arrResults(i).train_class_name
        Item.SubItems(2) = m_arrResults(i).arrive_time
        Item.SubItems(3) = m_arrResults(i).start_time
        If m_arrResults(i).stopover_time = "0" Then Item.SubItems(4) = "----" Else Item.SubItems(4) = m_arrResults(i).stopover_time + " 分"
        Item.SubItems(5) = m_arrResults(i).update_arrive_time
        Item.SubItems(6) = m_arrResults(i).update_start_time
        Item.SubItems(7) = m_arrResults(i).start_station_name
        Item.SubItems(8) = m_arrResults(i).end_station_name
        Item.SubItems(9) = m_arrResults(i).end_arrive_time
        Item.SubItems(10) = m_arrResults(i).running_time
        Item.SubItems(11) = m_arrResults(i).distance + " 公里"
        Item.SubItems(12) = m_arrResults(i).platform_no
        Item.SubItems(13) = m_arrResults(i).jiaolu_corporation_code
        Item.SubItems(14) = m_arrResults(i).jiaolu_dept_train
        Item.SubItems(15) = m_arrResults(i).jiaolu_train_style
        Item.SubItems(16) = m_arrResults(i).speed + " 公里/时"
    Next i
End Sub

Private Sub lsvResult_ItemDblClick(ByVal Item As VBCCR18.LvwListItem, ByVal Button As Integer)
    Dim strTrainCode As String
    Dim strTrainNo As String
    Dim i As Long
    
    strTrainCode = Item.Text
    strTrainNo = ""
    
    For i = 0 To m_lngResultCount - 1
        If StrComp(m_arrResults(i).station_train_code, strTrainCode, vbTextCompare) = 0 Then
            strTrainNo = m_arrResults(i).train_no
            Exit For
        End If
    Next
    
    If Len(strTrainCode) > 0 Then
        frmTrain.QueryTrainDirectly strTrainCode, strTrainNo
        frmTrain.Show
        frmTrain.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_asyncQuery Is Nothing Then
        m_asyncQuery.Abort
        Set m_asyncQuery = Nothing
    End If
End Sub

