VERSION 5.00
Begin VB.Form mainForm 
   Caption         =   "pingTest - 主窗口"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9735
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8520
      Top             =   3240
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim host, lstr As String
    host = Text2.Text
    Shell "cmd.exe /c ping " & host & " -n 1 > d:\ping.dat", vbHide
    
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    Open "d:\ping.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, lstr
        If lstr <> "" Then
            Text1.Text = Text1.Text & lstr & vbCrLf
            Text1.SelStart = Len(Text1)
            
            '检查结果中是否含义“out”字符串，如果有就代表网络不通，否则就是网络通
            If InStr(1, lstr, "out") > 0 Then
                Label1.Caption = "不通"
                Label1.BackColor = vbRed
                Exit Do
            Else
                Label1.Caption = "通"
                Label1.BackColor = vbGreen
            End If
        End If
    Loop
    Close #1
    Timer1.Enabled = False
    
End Sub
