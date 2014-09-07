VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "质数判断"
   ClientHeight    =   1860
   ClientLeft      =   2910
   ClientTop       =   3465
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   2325
   Begin VB.PictureBox picProcess 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   140
      ScaleHeight     =   255
      ScaleWidth      =   2040
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   2035
      Begin VB.Label lblProcess 
         BackStyle       =   0  'Transparent
         Caption         =   "输入一个正整数，再按回车"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10
         TabIndex        =   5
         Top             =   40
         Width           =   2055
      End
      Begin VB.Shape shpProcess 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.TextBox txtResult 
      Height          =   855
      Left            =   140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtN 
      Height          =   270
      Left            =   720
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
      Left            =   1250
      TabIndex        =   3
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "输入："
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   160
      Width           =   2415
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

    If txtN = "0" Then
        lblProcess = "0不在我们讨论范围内！"
        Exit Sub
    ElseIf txtN = "" Then
        lblProcess = "没有输入任何数！"
        Exit Sub
    ElseIf txtN = "1" Then
        shpProcess.Width = picProcess.Width
        lblProcess = "1既不是质数，也不是合数。"
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
        Dim processWidth  As Integer
        processWidth = (a - 1) / Sqr(n) * picProcess.Width
        processWidth = Int((processWidth - shpProcess.Width) / (0.01 * picProcess.Width))
        If processWidth > 0 Then shpProcess.Width = shpProcess.Width + processWidth * 0.01 * picProcess.Width
        lblProcess = "正在分解" & n
        DoEvents
        If n / a = Int(n / a) Then   'If n Mod a = 0 Then
            txtResult = txtResult & Format(a) & "×"
            n = n / a
            a = 2
        Else
            a = a + 1
        End If
    Wend
    shpProcess.Width = picProcess.Width
    
    If txtResult = "" Then
        lblProcess = Int(txtN) & "为质数。"
    Else
        lblProcess = Int(txtN) & "为合数。"
    End If
    txtResult = Int(txtN) & "＝" & txtResult & Format(n)
    
errShow:
    Select Case Err.Number
        Case 0
            
        Case Else
            txtResult = "错误：" & Err.Description & "（代号：" & Err.Number & "）。"
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
    lblProcess = "输入一个正整数，再按回车"
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
