VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "sshTest - ������"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9015
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "�˳�"
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
    Dim lstr As String                  '�洢�����ļ��ڵ�������Ϣ
    Dim mac_str(), name_str() As String                '����洢ÿһ�������ĵ�ַ��Ϣ
    Dim hostCount As Integer            '��������
    Dim i, j As Integer                   'FORѭ������ʱ����
    
    '�������ļ�
    Open "d:\Project\sshTest\1.txt" For Input As #1
    
    '�������ļ��е����ݸ�ֵ��lstr
    Do While Not EOF(1)
        Line Input #1, lstr
    Loop
    
    '�ر������ļ�
    Close #1
    
    '�������ݳ��Ȼ�ȡ��������
    hostCount = Len(lstr) / 18
    
    '���������������¶�������
    ReDim mac_str(hostCount)
    ReDim name_str(hostCount)
    
    '��ÿ������Ԫ�ظ�ֵ
    For i = 0 To hostCount - 1
        mac_str(i) = Mid(lstr, (1 + i * 18), 17)
    Next i
    
    '���ݻ�ȡ��MAC��ַ�����ݿ��������Ӧ������
    For i = 0 To hostCount - 1
        rsStr = "select * from base_info where mac_addr = '" & mac_str(i) & "'"
        myRs.Open rsStr, myCon, 1, 3
        name_str(i) = myRs.Fields("name")
        myRs.Close
        ssh_str = ssh_str & mac_str(i) & "  " & name_str(i) & vbCrLf
    Next i
    
    '��ʾ������Ϣ
    If ssh_str <> "" Then
        Text1.Text = ssh_str
    End If
    Timer1.Enabled = False
End Sub

'�˳���ť
Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    '�������ݿ�
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
