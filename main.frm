VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "sshTest - 主窗口"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9015
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1300
      Left            =   7800
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   8400
      Top             =   3600
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function HideCaret Lib "User32.dll" (ByVal hwnd As Long) As Boolean
Public ssh_str As String
Public myCon As New ADODB.Connection
Public myRs As New ADODB.Recordset
Public conStr, rsStr As String



Private Sub ssh_link()
    Shell "cmd.exe /c plink.exe -l root -pw 11 -m d:\Project\sshTest\cmd.txt 192.168.0.1 > d:\Project\sshTest\1.txt", vbHide
End Sub

Private Sub ssh_deal()
    Dim lstr As String                  '存储数据文件内的所有信息
    Dim mac_str(), name_str() As String                '数组存储每一个主机的地址信息
    Dim hostCount As Integer            '主机个数
    Dim i, j As Integer                   'FOR循环的临时变量
    
    '打开数据文件
    Open "d:\Project\sshTest\1.txt" For Input As #1
    
    '将数据文件中的内容赋值给lstr
    Do While Not EOF(1)
        Line Input #1, lstr
    Loop
    
    '关闭数据文件
    Close #1
    
    '根据数据长度获取主机数量
    hostCount = Len(lstr) / 18
    
    '根据主机数量重新定义数组
    ReDim mac_str(hostCount)
    ReDim name_str(hostCount)
    
    '给每个数组元素赋值
    For i = 0 To hostCount - 1
        mac_str(i) = Mid(lstr, (1 + i * 18), 17)
    Next i
    
    '根据获取的MAC地址到数据库里检索对应的名称
    For i = 0 To hostCount - 1
        rsStr = "select * from base_info where mac_addr = '" & mac_str(i) & "'"
        myRs.Open rsStr, myCon, 1, 3
        name_str(i) = myRs.Fields("name")
        myRs.Close
        ssh_str = ssh_str & mac_str(i) & "  " & name_str(i) & vbCrLf
    Next i
    
    '显示主机信息
    If ssh_str <> "" Then
        Text1.Text = ssh_str
    End If
    Timer1.Enabled = False
End Sub

'退出按钮
Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    '连接数据库
    conStr = "driver={sql server};server=xp1;uid=sa;pwd=11;Database=sshTest"
    myCon.Open conStr
End Sub

Private Sub Text1_GotFocus()
    HideCaret Text1.hwnd
End Sub

Private Sub Timer1_Timer()
    ssh_str = ""
    Call ssh_deal
End Sub

Private Sub Timer2_Timer()
    Call ssh_link
    Timer1.Enabled = True
End Sub
