VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{9DB12C0E-8736-49E6-9CA1-896A1E67D6F3}#1.0#0"; "vsflex8n.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   13590
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   720
      Top             =   5520
   End
   Begin VB.Frame Frame2 
      Caption         =   "产品信息"
      Height          =   5175
      Left            =   3400
      TabIndex        =   20
      Top             =   360
      Width           =   6015
      Begin VB.TextBox TextPhone2 
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Text            =   "电话号码"
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton btAphone2 
         Caption         =   "增加"
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton btDphone2 
         Caption         =   "删除"
         Height          =   375
         Left            =   4680
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TextAddre 
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Text            =   "分机位置"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CommandButton btAadd 
         Caption         =   "增加"
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton btDadd 
         Caption         =   "删除"
         Height          =   375
         Left            =   4680
         TabIndex        =   21
         Top             =   2160
         Width           =   975
      End
      Begin VSFlex8NCtl.VSFlexGrid xGrid2 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
         _cx             =   4471
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8NCtl.VSFlexGrid xGrid3 
         Height          =   2055
         Left            =   240
         TabIndex        =   0
         Top             =   2880
         Width           =   5535
         _cx             =   9763
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame3 
      Caption         =   "分机信息"
      Height          =   5175
      Left            =   9720
      TabIndex        =   16
      Top             =   360
      Width           =   3615
      Begin VB.CommandButton btMsg 
         Caption         =   "短信测试"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   800
         Width           =   1600
      End
      Begin VB.TextBox Text1 
         Height          =   3615
         Left            =   360
         TabIndex        =   18
         Text            =   "三相电压实时监测调试"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton btTest 
         Caption         =   "分机调试"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1600
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton bttime 
      Caption         =   "时间同步"
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "维护信息"
      Height          =   5175
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   2775
      Begin VB.TextBox TextSA 
         Height          =   350
         Left            =   1080
         TabIndex        =   31
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox TextSC 
         Height          =   350
         Left            =   1080
         TabIndex        =   30
         Top             =   840
         Width           =   1500
      End
      Begin VB.TextBox TextSN 
         Height          =   350
         Left            =   1080
         TabIndex        =   29
         Top             =   360
         Width           =   1500
      End
      Begin VB.CommandButton btDphone1 
         Caption         =   "删除"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton btAphone1 
         Caption         =   "增加"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox TextPhone1 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Text            =   "电话号码"
         Top             =   4200
         Width           =   2295
      End
      Begin VSFlex8NCtl.VSFlexGrid xGrid1 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   2535
         _cx             =   4471
         _cy             =   2990
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbSupplierA 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1360
         Width           =   1000
      End
      Begin VB.Label lbSupplierC 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   880
         Width           =   1000
      End
      Begin VB.Label lbSupplierN 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   380
         Width           =   1000
      End
   End
   Begin VB.CommandButton btWriteDate 
      Caption         =   "写入"
      Height          =   500
      Left            =   8520
      TabIndex        =   8
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton btWriteXls 
      Caption         =   "导出数据"
      Height          =   500
      Left            =   5640
      TabIndex        =   7
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton btReadXls 
      Caption         =   "导入数据"
      Height          =   500
      Left            =   4200
      TabIndex        =   6
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton btRS232 
      Caption         =   "打开串口"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.ComboBox comPort 
      Height          =   300
      ItemData        =   "frmMain.frx":0000
      Left            =   1080
      List            =   "frmMain.frx":001C
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   6000
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StbShow 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6525
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lbRS232 
      Caption         =   "串口号："
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim comState As String
    
Private Sub waittime(delay As Single)
Dim starttime As Single
starttime = Timer
Do Until (Timer - starttime) > delay
DoEvents
Loop
End Sub
Public Function DecToHex(ByVal DecNumber As Byte) As String
    If DecNumber <= 15 Then
        DecToHex = "0" & Hex(DecNumber)
    Else: DecToHex = Hex(DecNumber)
    End If
End Function
Private Sub btAadd_Click()
    If xGrid3.TextMatrix(1, 1) = "" Then
        xGrid3.TextMatrix(1, 0) = 1
        xGrid3.TextMatrix(1, 1) = 1
        xGrid3.TextMatrix(1, 2) = TextAddre.Text
        xGrid3.TextMatrix(1, 3) = 0
    Else
        xGrid3.Rows = xGrid3.Rows + 1
        xGrid3.TextMatrix(xGrid3.Rows - 1, 0) = xGrid3.Rows - 1
        xGrid3.TextMatrix(xGrid3.Rows - 1, 1) = xGrid3.Rows - 1
        xGrid3.TextMatrix(xGrid3.Rows - 1, 2) = TextAddre.Text
        xGrid3.TextMatrix(xGrid3.Rows - 1, 3) = 0
    End If
    
    If xGrid3.ScrollBars = flexScrollBarHorizontal Then
        xGrid3.ColWidth(2) = xGrid3.Width - 1920
    End If
    
End Sub

Private Sub btAphone1_Click()
    If xGrid1.TextMatrix(1, 1) = "" Then
        xGrid1.TextMatrix(1, 0) = 1
        xGrid1.TextMatrix(1, 1) = TextPhone1.Text
    Else
        xGrid1.Rows = xGrid1.Rows + 1
        xGrid1.TextMatrix(xGrid1.Rows - 1, 0) = xGrid1.Rows - 1
        xGrid1.TextMatrix(xGrid1.Rows - 1, 1) = TextPhone1.Text
    End If
End Sub

Private Sub btAphone2_Click()
    If xGrid2.TextMatrix(1, 1) = "" Then
        xGrid2.TextMatrix(1, 0) = 1
        xGrid2.TextMatrix(1, 1) = TextPhone2.Text
    Else
        xGrid2.Rows = xGrid2.Rows + 1
        xGrid2.TextMatrix(xGrid2.Rows - 1, 0) = xGrid2.Rows - 1
        xGrid2.TextMatrix(xGrid2.Rows - 1, 1) = TextPhone2.Text
    End If
    
End Sub

Private Sub btDadd_Click()
    If xGrid3.Rows < 3 Then
        xGrid3.TextMatrix(1, 0) = ""
        xGrid3.TextMatrix(1, 1) = ""
        xGrid3.TextMatrix(1, 2) = ""
        xGrid3.TextMatrix(1, 3) = ""
    Else
        xGrid3.Rows = xGrid3.Rows - 1
    End If
End Sub

Private Sub btDphone1_Click()
    If xGrid1.Rows < 3 Then
        xGrid1.TextMatrix(1, 0) = ""
        xGrid1.TextMatrix(1, 1) = ""
    Else
        xGrid1.Rows = xGrid1.Rows - 1
    End If
    
End Sub

Private Sub btDphone2_Click()
    If xGrid2.Rows < 3 Then
        xGrid2.TextMatrix(1, 0) = ""
        xGrid2.TextMatrix(1, 1) = ""
    Else
        xGrid2.Rows = xGrid2.Rows - 1
    End If
End Sub

Private Sub btMsg_Click()

    On Error Resume Next
        MSComm1.Output = Trim(("+G " + Chr(10)))
        If Err.Number = 8018 Then
            MsgBox "" & "串口未打开"
        End If

End Sub

Private Sub btReadXls_Click()
Dim A(101) As String
Dim i As Integer
Dim xlsApp As Excel.Application
Dim xlsWorkbook As Excel.Workbook
Dim xlssheet As Excel.Worksheet
   Set xlsApp = CreateObject("Excel.Application")
   Set xlsWorkbook = xlsApp.Workbooks.Open("D:\my.xls")
   xlsApp.Visible = False
   Set xlssheet = xlsApp.Worksheets("Sheet1")
   'xlsApp.Range("A1").Select
   
    xGrid1.Rows = 2
    xGrid1.TextMatrix(1, 0) = ""
    xGrid1.TextMatrix(1, 1) = ""
    xGrid2.Rows = 2
    xGrid2.TextMatrix(1, 0) = ""
    xGrid2.TextMatrix(1, 1) = ""
    xGrid3.Rows = 2
    xGrid3.TextMatrix(1, 0) = ""
    xGrid3.TextMatrix(1, 1) = ""
    xGrid3.TextMatrix(1, 2) = ""
    xGrid3.TextMatrix(1, 3) = 0
   
For i = 1 To 15
    A(i) = xlssheet.Cells(i + 4, 4).Value
    If (A(i)) <> ("") Then
        If xGrid1.TextMatrix(1, 1) = "" Then
            xGrid1.TextMatrix(1, 0) = 1
            xGrid1.TextMatrix(1, 1) = A(i)
        Else
            xGrid1.Rows = xGrid1.Rows + 1
            xGrid1.TextMatrix(xGrid1.Rows - 1, 0) = xGrid1.Rows - 1
            xGrid1.TextMatrix(xGrid1.Rows - 1, 1) = A(i)
            'xGrid1.Col
        End If
    
    End If
Next i

For i = 1 To 15
    A(i) = xlssheet.Cells(i + 4, 5).Value
    If (A(i)) <> ("") Then
        If xGrid2.TextMatrix(1, 1) = "" Then
            xGrid2.TextMatrix(1, 0) = 1
            xGrid2.TextMatrix(1, 1) = A(i)
        Else
            xGrid2.Rows = xGrid2.Rows + 1
            xGrid2.TextMatrix(xGrid2.Rows - 1, 0) = xGrid2.Rows - 1
            xGrid2.TextMatrix(xGrid2.Rows - 1, 1) = A(i)
        End If
    
    End If
Next i


For i = 1 To 15
    A(i) = xlssheet.Cells(i + 4, 10).Value
    If (A(i)) <> ("") Then
        If xGrid3.TextMatrix(1, 1) = "" Then
        xGrid3.TextMatrix(1, 0) = 1
        xGrid3.TextMatrix(1, 1) = 1
        xGrid3.TextMatrix(1, 2) = A(i)
        xGrid3.TextMatrix(1, 3) = 0
    Else
        xGrid3.Rows = xGrid3.Rows + 1
        xGrid3.TextMatrix(xGrid3.Rows - 1, 0) = xGrid3.Rows - 1
        xGrid3.TextMatrix(xGrid3.Rows - 1, 1) = xGrid3.Rows - 1
        xGrid3.TextMatrix(xGrid3.Rows - 1, 2) = A(i)
        xGrid3.TextMatrix(xGrid3.Rows - 1, 3) = 0
    End If
    
    End If
Next i

A(1) = xlssheet.Cells(22, 5).Value
A(2) = xlssheet.Cells(23, 5).Value
A(3) = xlssheet.Cells(24, 5).Value
TextSN.Text = A(1)
TextSC.Text = A(2)
TextSA.Text = A(3)

xlsWorkbook.Close
xlsApp.Quit
Set xlssheet = Nothing
Set xlsWorkbook = Nothing
Set xlsApp = Nothing

End Sub

Private Sub btRS232_Click()
    On Error Resume Next
    Dim baud As String, parity As String, stopbit As String, numBit, comPortNum As Integer
    If btRS232.Caption = "打开串口" Then
        '串口选择
        Select Case comPort.ListIndex
            Case -1
                comPortNum = 0
            Case Else
                comPortNum = comPort.ItemData(comPort.ListIndex)
        End Select
        '波特率选择
                baud = "9600"
        '校验方式选择
                parity = "N"
        '停止位选择
                stopbit = "1"
        '数据位选择
                numBit = "8"
        '检查输入是否有误
        If comPortNum = 0 Then
            MsgBox "请选择串口"
        End If
        If comPortNum <> 0 And baud <> "" And parity <> "" And stopbit <> "" And numBit <> "" Then
            MSComm1.CommPort = comPortNum
            MSComm1.Settings = baud & "," & parity & "," & numBit & "," & stopbit
            MSComm1.InputMode = comInputModeText
            MSComm1.RThreshold = 1  '接收缓冲区每收到一个字符触发事件onComm()
            MSComm1.PortOpen = True
            If Err.Number = 8002 Then
                MsgBox "无效的端口号", vbOKOnly, "信息"
                GoTo Label1
            ElseIf Err.Number = 8005 Then
                MsgBox "端口被占用", vbOKOnly, "信息"
                GoTo Label1
            Else
                StbShow.Panels.Item(2).Text = "串口状态：" & "连接中 COM" & MSComm1.CommPort
                btRS232.Caption = "关闭串口"
            End If
        Else
            MsgBox "通信参数不完整", vbOKOnly, "信息"
            GoTo Label1
        End If
    ElseIf btRS232.Caption = "关闭串口" Then
        MSComm1.PortOpen = False
        btRS232.Caption = "打开串口"
        StbShow.Panels.Item(2).Text = "串口状态：" & "未连接"
    End If
Label1:
End Sub

Private Sub btTest_Click()
    
    If btTest.Caption = "分机调试" Then
        btTest.Caption = "退出调试"
        Shape1.FillColor = &HFF00&
        Timer1.Enabled = True
    Else
        Shape1.FillColor = &HFF&
        btTest.Caption = "分机调试"
        Timer1.Enabled = False
    On Error Resume Next
        MSComm1.Output = Trim(("+H Q" + Chr(10)))
        
        If Err.Number = 8018 Then
            MsgBox "" & "串口未打开"
        End If
    
    End If
        
End Sub
Private Sub bttime_Click()
    Dim timeNow() As Byte
    Dim timestr As String

    timestr = Date + Time
    timeNow = StrConv(timestr, vbFromUnicode)
        On Error Resume Next
            MSComm1.Output = Trim("+I " & (timeNow(2) - 48) & (timeNow(3) - 48) & (timeNow(5) - 48) & _
            (timeNow(6) - 48) & (timeNow(8) - 48) & (timeNow(9) - 48) & (timeNow(11) - 48) & _
            (timeNow(12) - 48) & (timeNow(14) - 48) & (timeNow(15) - 48) & (timeNow(17) - 48) & _
            (timeNow(18) - 48) & Chr(10))
        If Err.Number = 8018 Then
            MsgBox "" & "串口未打开"
        End If
    
End Sub

Private Sub btWriteDate_Click()
    If MsgBox("写入数据确认", vbYesNo, "提示") = vbYes Then
    
        Dim n
        Dim strm() As Byte
        Dim strms As String
        Dim strms1 As String
        On Error Resume Next
            strm() = ""
            strms = ""
            strm() = TextSC.Text
            For n = 0 To Len(TextSC.Text)
                strms = strms + DecToHex(strm(n * 2 + 1)) + DecToHex(strm(n * 2))
            Next n
        
            MSComm1.Output = Trim(("+F " + TextSN.Text + strms + Chr(10)))
            
            waittime (1)
        If Err.Number = 8018 Then
            MsgBox "" & "串口未打开"
        End If
      
        On Error Resume Next
            strm() = ""
            strms1 = ""
            strm() = TextSA.Text
            For n = 0 To Len(TextSA.Text)
                strms1 = strms1 + DecToHex(strm(n * 2 + 1)) + DecToHex(strm(n * 2))
            Next n
        
            MSComm1.Output = Trim(("+J " + strms1 + Chr(10)))
            waittime (1)
        If Err.Number = 8018 Then
            MsgBox "" & "串口未打开"
        End If
        
        Dim stt() As Byte
        Dim sttt As String
        Dim m
        Dim i
        For i = 1 To (xGrid3.Rows - 1)
            stt() = ""
            sttt = ""
            On Error Resume Next
                stt() = xGrid3.TextMatrix(i, 1) + xGrid3.TextMatrix(i, 2)
            For m = 0 To Len(xGrid3.TextMatrix(i, 1) + xGrid3.TextMatrix(i, 2))
                sttt = sttt + DecToHex(stt(m * 2 + 1)) + DecToHex(stt(m * 2))
            Next m
            MSComm1.Output = Trim(("+D " + sttt + Chr(10)))
            waittime (1)
            If Err.Number = 8018 Then
                MsgBox "" & "串口未打开"
            End If
        Next i
        
        Dim stss() As Byte
        Dim stst As String
        Dim j
        For j = 1 To (xGrid1.Rows - 1)
            stss() = ""
            stst = ""
            On Error Resume Next
                stss() = xGrid1.TextMatrix(j, 1) + "F"
            For m = 0 To 5
                stst = stst + Chr(stss(m * 4 + 2)) + Chr(stss(m * 4))
            Next m
            MSComm1.Output = Trim(("+B " + stst + Chr(10)))
            waittime (1)
            If Err.Number = 8018 Then
                MsgBox "" & "串口未打开"
            End If
        Next j
        
        For j = 1 To (xGrid2.Rows - 1)
            stss() = ""
            stst = ""
            On Error Resume Next
                stss() = xGrid2.TextMatrix(j, 1) + "F"
            For m = 0 To 5
                stst = stst + Chr(stss(m * 4 + 2)) + Chr(stss(m * 4))
            Next m
            MSComm1.Output = Trim(("+K " + stst + Chr(10)))
            waittime (1)
            If Err.Number = 8018 Then
                MsgBox "" & "串口未打开"
            End If
        Next j
        
        
        
    End If
End Sub

Private Sub btWriteXls_Click()
    If MsgBox("导出数据确认", vbYesNo, "提示") = vbYes Then
        Dim A(101) As String
        Dim i As Integer
        Dim xlsApp As Excel.Application
        Dim xlsWorkbook As Excel.Workbook
        Dim xlssheet As Excel.Worksheet
        Set xlsApp = CreateObject("Excel.Application")
        Set xlsWorkbook = xlsApp.Workbooks.Open("D:\my.xls")
        xlsApp.Visible = False
        Set xlssheet = xlsApp.Worksheets("Sheet1")
   
        xlssheet.Cells(22, 5).Value = TextSN.Text
        xlssheet.Cells(23, 5).Value = TextSC.Text
        xlssheet.Cells(24, 5).Value = TextSA.Text
   
   
        For i = 1 To 15
            xlssheet.Cells(i + 4, 4).Value = ""
        Next i
        For i = 1 To 15
            xlssheet.Cells(i + 4, 5).Value = ""
        Next i
        For i = 1 To 50
            xlssheet.Cells(i + 4, 9).Value = ""
            xlssheet.Cells(i + 4, 10).Value = ""
        Next i
        
        For i = 1 To xGrid1.Rows - 1
                xlssheet.Cells(i + 4, 4).Value = xGrid1.TextMatrix(i, 1)
        Next i


        For i = 1 To xGrid2.Rows - 1
                xlssheet.Cells(i + 4, 5).Value = xGrid2.TextMatrix(i, 1)
        Next i

        For i = 1 To xGrid3.Rows - 1
                xlssheet.Cells(i + 4, 9).Value = xGrid3.TextMatrix(i, 1)
                xlssheet.Cells(i + 4, 10).Value = xGrid3.TextMatrix(i, 2)
        Next i
    xlsWorkbook.Close
    xlsApp.Quit
    Set xlssheet = Nothing
    Set xlsWorkbook = Nothing
    Set xlsApp = Nothing
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
MsgBox "这是check1(0)"
Case 1
MsgBox "这是check1(1)"
Case 2
MsgBox "这是check1(2)"
Case 3
MsgBox "这是check1(3)"
Case 4
MsgBox "这是check1(4)"
Case 5
MsgBox "这是check1(5)"
Case 6
MsgBox "这是check1(6)"

End Select
End Sub

Private Sub Form_Load()
    Dim guserName As String
    Dim guserKind As String

    
    Me.BorderStyle = 1
    Me.Caption = "三相电主机读写上位机V2.0"
    
    guserName = frmLogin.Fusername
    guserKind = "管理员"
    comState = "未连接"
    
    Shape1.FillColor = &HFF&
    
    Text1.Text = "A相：" & 0 & " B相：" & 0 & " C相：" & 0 & " 电池：" & 0
    
    With StbShow
    .Panels.Add (1)
    .Panels.Item(1).Width = Me.Width - 5200
    .Panels.Add (2)
    .Panels.Item(2).Width = 3000
    .Panels.Add (3)
    .Panels.Item(3).Width = 1200
    .Panels.Item(4).Width = 1000
    .Panels.Item(1).Style = sbrText
    .Panels.Item(2).Style = sbrText
    .Panels.Item(2).Alignment = sbrRight
    .Panels.Item(3).Style = sbrDate
    .Panels.Item(4).Style = sbrTime
    .Panels.Item(1).Text = "当前用户：" & guserName & "    权限：" & guserKind
    .Panels.Item(2).Text = "串口状态：" & comState
    End With
    
    lbSupplierN.Caption = "主机编号:"
    lbSupplierC.Caption = "公司名称:"
    lbSupplierA.Caption = "公司地址:"
    
    With xGrid1
    .BorderStyle = flexBorderSingle
    .FormatString = "|^联系方式1"
    .ColWidth(0) = 300
    .ColWidth(1) = xGrid1.Width - 380
    .RowHeight(0) = 400
    .Row = 0
    .Col = 0
    End With

    With xGrid2
    .FormatString = "|^联系方式2"
    .ColWidth(0) = 300
    .ColWidth(1) = xGrid2.Width - 380
    .RowHeight(0) = 400
    .Row = 0
    .Col = 0
    End With

    With xGrid3
    .FormatString = "|^电柜编号|^电柜地址 "
    .ColWidth(0) = 300
    .ColWidth(1) = 1000
    .ColWidth(2) = xGrid3.Width - 1680
    .ColWidth(3) = 300
    .RowHeight(0) = 400
    .ColDataType(3) = flexDTBoolean
    .Row = 0
    .Col = 0
    End With
    
    
    With comPort
    .ItemData(0) = 1
    .ItemData(1) = 2
    .ItemData(2) = 3
    .ItemData(3) = 4
    .ItemData(4) = 5
    .ItemData(5) = 6
    .ItemData(6) = 7
    .ItemData(7) = 8
    .ListIndex = 0
    End With
    
End Sub

Private Sub MSComm1_OnComm()
    Dim buf As String
    Dim InData As Variant ' 变体变量
    Dim Arr() As Byte    ' 接收字节数组
    Dim s10 As String
    Dim s11 As String
    Dim s12 As String
    MSComm1.InputMode = comInputModeBinary
    Select Case MSComm1.CommEvent
            Case comEvReceive ' 触发接收事件
            waittime (1)
            InData = Trim(MSComm1.Input) ' 接收数据
    End Select
    Arr = InData

    If Len(InData) > 0 Then
    Text1.Text = "A相：" & Val(Arr(1)) & " B相：" & Val(Arr(2)) & " C相：" & Val(Arr(3))
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
        MSComm1.Output = Trim(("+H 1" + Chr(10)))
    If Err.Number = 8018 Then
        MsgBox "" & "串口未打开"
    End If
End Sub

Private Sub xGrid3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    MsgBox "" & xGrid3.RowSel
    'Text1.Text = ""
    If xGrid3.TextMatrix(xGrid3.RowSel, 3) <> "" Then
        If xGrid3.TextMatrix(xGrid3.RowSel, 3) = 0 Then
            xGrid3.TextMatrix(xGrid3.RowSel, 3) = 1
        Else
            xGrid3.TextMatrix(xGrid3.RowSel, 3) = 0
        End If
    End If
End If
End Sub


