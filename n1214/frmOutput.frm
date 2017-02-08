VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3165
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A(101) As String
Dim i As Integer
Dim xlsApp As Excel.Application
Dim xlsWorkbook As Excel.Workbook
Dim xlssheet As Excel.Worksheet
   Set xlsApp = CreateObject("Excel.Application")
   Set xlsWorkbook = xlsApp.Workbooks.Open("D:\my.xls")
   xlsApp.Visible = False
   Set xlssheet = xlsApp.Worksheets("Sheet1")
   
   
For i = 1 To 15

    If (MSHFlexGrid1.TextMatrix(i, 1)) <> ("") Then
        xlssheet.Cells(i + 4, 4).Value = MSHFlexGrid1.TextMatrix(i, 1)
    End If
Next i


For i = 1 To 15

    If (MSHFlexGrid2.TextMatrix(i, 1)) <> ("") Then
        xlssheet.Cells(i + 4, 5).Value = MSHFlexGrid2.TextMatrix(i, 1)
    End If
Next i

xlsWorkbook.Close
xlsApp.Quit
Set xlssheet = Nothing
Set xlsWorkbook = Nothing
Set xlsApp = Nothing

Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "确认导出"
End Sub
