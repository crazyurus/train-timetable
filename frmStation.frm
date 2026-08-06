VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.Form frmStation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "根据始发站查询列车时刻表"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10230
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
   ScaleHeight     =   5535
   ScaleWidth      =   10230
   StartUpPosition =   2  '屏幕中心
   Begin 列车时刻表.StationPicker spTo 
      Height          =   2295
      Left            =   6600
      TabIndex        =   7
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4048
   End
   Begin 列车时刻表.StationPicker spFrom 
      Height          =   2295
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4048
   End
   Begin VBCCR18.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5160
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   661
      InitPanels      =   "frmStation.frx":0000
   End
   Begin VBCCR18.CommandButtonW cmdQuery 
      Default         =   -1  'True
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "查询"
   End
   Begin VBCCR18.DTPicker datePicker 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Value           =   32874
      CustomFormat    =   "frmStation.frx":04F0
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.ListView lsvResult 
      Height          =   4215
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   9720
      _ExtentX        =   17145
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "到达站："
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   300
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "出发站："
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   300
      Width           =   720
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

Private Type TicketResult
    train_no As String
    station_train_code As String
    train_class_name As String
    from_station_name As String
    to_station_name As String
    start_time As String
    arrive_time As String
    lishi As String
    day_difference As String
    swz_price As String
    zy_price As String
    ze_price As String
End Type

Private m_arrTickets() As TicketResult
Private m_lngTicketCount As Long
Private m_asyncQuery As AsyncRequest

Private Sub Form_Load()
    datePicker.value = Date
    m_lngTicketCount = 0
    
    With lsvResult.ColumnHeaders
        .Add , , "车次", 900
        .Add , , "类型", 700
        .Add , , "出发站", 1200
        .Add , , "到达站", 1200
        .Add , , "出发时间", 1000
        .Add , , "到达时间", 1000
        .Add , , "历时", 800
        .Add , , "商务座", 800
        .Add , , "一等座", 800
        .Add , , "二等座", 800
    End With
End Sub

Private Sub cmdQuery_Click()
    Dim strFromCode As String
    Dim strToCode As String
    Dim strDate As String
    Dim strUrl As String
    
    On Error GoTo ErrorHandler
    
    strFromCode = spFrom.StationCode
    If Len(strFromCode) = 0 Then
        MsgBox "请选择出发站", vbExclamation
        spFrom.SetFocus
        Exit Sub
    End If
    
    strToCode = spTo.StationCode
    If Len(strToCode) = 0 Then
        MsgBox "请选择到达站", vbExclamation
        spTo.SetFocus
        Exit Sub
    End If
    
    If Not m_asyncQuery Is Nothing Then
        m_asyncQuery.Abort
        Set m_asyncQuery = Nothing
    End If
    
    cmdQuery.Enabled = False
    statusBar.Panels(1).Text = "正在查询列车信息……"
    lsvResult.ListItems.Clear
    
    strDate = Format(datePicker.value, "yyyy-MM-dd")
    strUrl = "https://kyfw.12306.cn/otn/leftTicketPrice/queryAllPublicPrice?" & _
             "leftTicketDTO.train_date=" & strDate & _
             "&leftTicketDTO.from_station=" & strFromCode & _
             "&leftTicketDTO.to_station=" & strToCode & _
             "&purpose_codes=ADULT"
    
    Set m_asyncQuery = New AsyncRequest
    m_asyncQuery.GetRequest strUrl, Me, "OnQueryComplete"
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
    
    ParseTicketResult responseText
    ShowResults
    statusBar.Panels(1).Text = "查询完成，共找到 " & m_lngTicketCount & " 趟列车"
    cmdQuery.Enabled = True
    Set m_asyncQuery = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "查询出错: " & Err.Description, vbCritical
    statusBar.Panels(1).Text = "查询出错"
    cmdQuery.Enabled = True
End Sub

Private Sub ParseTicketResult(ByVal strJson As String)
    Dim parser As JSON
    Dim root As Object
    Dim dataArr As Object
    Dim lngCount As Long
    Dim i As Long
    Dim itemObj As Object
    Dim dto As Object
    
    On Error GoTo ErrHandler
    m_lngTicketCount = 0
    
    Set parser = New JSON
    Set root = parser.Parse(strJson)
    Set parser = Nothing
    If root Is Nothing Then Exit Sub
    
    If Not root.Exists("data") Then Exit Sub
    Set dataArr = root("data")
    If TypeName(dataArr) <> "Collection" Then Exit Sub
    lngCount = dataArr.Count
    If lngCount = 0 Then Exit Sub
    
    ReDim m_arrTickets(0 To lngCount - 1)
    m_lngTicketCount = lngCount
    
    For i = 1 To lngCount
        Set itemObj = dataArr(i)
        If TypeName(itemObj) = "Dictionary" Then
            If itemObj.Exists("queryLeftNewDTO") Then
                Set dto = itemObj("queryLeftNewDTO")
                If TypeName(dto) = "Dictionary" Then
                    With m_arrTickets(i - 1)
                        If dto.Exists("train_no") Then .train_no = CStr(dto("train_no")) Else .train_no = ""
                        If dto.Exists("station_train_code") Then .station_train_code = CStr(dto("station_train_code")) Else .station_train_code = ""
                        If dto.Exists("train_class_name") Then .train_class_name = CStr(dto("train_class_name")) Else .train_class_name = ""
                        If dto.Exists("from_station_name") Then .from_station_name = CStr(dto("from_station_name")) Else .from_station_name = ""
                        If dto.Exists("to_station_name") Then .to_station_name = CStr(dto("to_station_name")) Else .to_station_name = ""
                        If dto.Exists("start_time") Then .start_time = CStr(dto("start_time")) Else .start_time = ""
                        If dto.Exists("arrive_time") Then .arrive_time = CStr(dto("arrive_time")) Else .arrive_time = ""
                        If dto.Exists("lishi") Then .lishi = CStr(dto("lishi")) Else .lishi = ""
                        If dto.Exists("day_difference") Then .day_difference = CStr(dto("day_difference")) Else .day_difference = ""
                        If dto.Exists("swz_price") Then .swz_price = FormatPrice(CStr(dto("swz_price"))) Else .swz_price = ""
                        If dto.Exists("zy_price") Then .zy_price = FormatPrice(CStr(dto("zy_price"))) Else .zy_price = ""
                        If dto.Exists("ze_price") Then .ze_price = FormatPrice(CStr(dto("ze_price"))) Else .ze_price = ""
                    End With
                End If
            End If
        End If
    Next i
    Exit Sub
    
ErrHandler:
    m_lngTicketCount = 0
End Sub

Private Function FormatPrice(ByVal priceStr As String) As String
    Dim d As Double
    On Error Resume Next
    d = CDbl(priceStr) / 100
    If d > 0 Then
        FormatPrice = ChrW(165) & Format(d, "0.0")
    Else
        FormatPrice = ""
    End If
End Function

Private Sub ShowResults()
    Dim i As Long
    Dim Item As Object
    
    lsvResult.ListItems.Clear
    For i = 0 To m_lngTicketCount - 1
        Set Item = lsvResult.ListItems.Add(, , m_arrTickets(i).station_train_code)
        Item.SubItems(1) = m_arrTickets(i).train_class_name
        Item.SubItems(2) = m_arrTickets(i).from_station_name
        Item.SubItems(3) = m_arrTickets(i).to_station_name
        Item.SubItems(4) = m_arrTickets(i).start_time
        Item.SubItems(5) = m_arrTickets(i).arrive_time
        Item.SubItems(6) = m_arrTickets(i).lishi
        Item.SubItems(7) = m_arrTickets(i).swz_price
        Item.SubItems(8) = m_arrTickets(i).zy_price
        Item.SubItems(9) = m_arrTickets(i).ze_price
    Next i
End Sub

Private Sub lsvResult_ItemDblClick(ByVal Item As VBCCR18.LvwListItem, ByVal Button As Integer)
    Dim strTrainCode As String
    Dim strTrainNo As String
    Dim i As Long
    
    strTrainCode = Item.Text
    strTrainNo = ""
    
    For i = 0 To m_lngTicketCount - 1
        If StrComp(m_arrTickets(i).station_train_code, strTrainCode, vbTextCompare) = 0 Then
            strTrainNo = m_arrTickets(i).train_no
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
