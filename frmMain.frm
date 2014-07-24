VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "质数判断"
   ClientHeight    =   2250
   ClientLeft      =   2910
   ClientTop       =   3465
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3135
   Begin VB.TextBox txtResult 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtN 
      Height          =   270
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "By: 闪闪的星"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2080
      Width           =   975
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "输入后按回车："
      Height          =   255
      Left            =   160
      TabIndex        =   0
      Top             =   165
      Width           =   2415
   End
   Begin VB.Shape shpProcess 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub startProcess()
    On Error GoTo errShow

    If txtN = "0" Or txtN = "1" Then
        txtResult = "发生了错误：0和1不在我们讨论范围内。"
        Exit Sub
    ElseIf txtN = "" Then
        txtResult = "发生了错误：没有输入任何数。"
        Exit Sub
    End If
    
    txtN = Format(Abs(Int(txtN)))
    Call txtN_Change
    txtN.Enabled = False
    txtResult.Enabled = False
    
    Dim n, a As Long
    n = Int(txtN)
    a = 2
    
    While a <= Sqr(n)
        shpProcess.Width = (a - 1) / Sqr(n) * Me.Width
        DoEvents
        If n / a = Int(n / a) Then   'If n Mod a = 0 Then
            txtResult = txtResult & Format(a) & "×" & Format(n / a) & ", "
        End If
        a = a + 1
    Wend
    shpProcess.Width = Me.Width
    
    If txtResult = "" Then
        txtResult = n & "为质数，因数只有1与它本身。"
    Else
        txtResult = n & "为合数，除1与其本身之外的全部因数列举如下：" & vbCr & vbLf & Left(txtResult, Len(txtResult) - 2)
    End If
    
errShow:
    Select Case Err.Number
        Case 0
            
        'Case 6
        '    If txtResult = "" Then
        '        txtResult = "发生了错误：溢出，数目已超出本程序的处理范围。（代号：6）" & vbCr & vbLf & "程序也无法肯定" & n & "是质数还是合数。"
        '    Else
        '        txtResult = "发生了错误：溢出，数目已超出本程序的处理范围。（代号：6）" & vbCr & vbLf & "但程序肯定" & n & "为合数，除1与其本身之外的部分因数列举如下：" & Left(txtResult, Len(txtResult) - 2)
        '    End If
        Case Else
            txtResult = "发生了错误：" & Err.Description & "（代号：" & Err.Number & "）。"
    End Select
    
    txtN.Enabled = True
    txtResult.Enabled = True
End Sub

Private Sub Form_Load()
    Call txtN_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub txtN_Change()
    txtResult = ""
    shpProcess.Width = 0
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8, 48 To 57 '若不是输入的数字或退格
        
    Case 13
        Call startProcess
    Case Else
        KeyAscii = 0 '不让输入的键生效
    End Select
End Sub
