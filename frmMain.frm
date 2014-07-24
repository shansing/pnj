VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "质数判断"
   ClientHeight    =   2175
   ClientLeft      =   2910
   ClientTop       =   3465
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   2880
   Begin VB.CommandButton cmdGo 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtResult 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtAB 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "2×2"
      Top             =   550
      Width           =   1095
   End
   Begin VB.TextBox txtN 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblCopyright 
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
      Left            =   1800
      TabIndex        =   5
      Top             =   2000
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Caption         =   "请输入一个正整数："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
    On Error GoTo errShow

    If txtN = "0" Or txtN = "1" Then
        txtResult = "过程未开始，因为发生了错误：0和1不在我们讨论范围内。"
        Exit Sub
    ElseIf txtN = "" Then
        txtResult = "过程未开始，因为发生了错误：没有输入任何数。"
        Exit Sub
    End If
    
    Call txtN_Change
    txtN.Enabled = False
    txtResult.Enabled = False
    cmdGo.Enabled = False
    
    Dim n, a, b As Integer
    n = Int(txtN)
    a = 2
    b = 2
    While a <= Sqr(n)
        While b <= n / a
            txtAB = Format(a) & "×" & Format(b)
            DoEvents
            
            If a * b = n Then
                txtResult = txtResult & Format(a) & "×" & Format(b) & ", "
            End If
            b = b + 1
        Wend
        b = 2
        a = a + 1
    Wend
    
    If txtResult = "" Then
        txtResult = "过程完毕。" & n & "为质数，因数只有1与它本身。"
    Else
        txtResult = "过程完毕。" & n & "为合数，除1与其本身之外的全部因数列举如下：" & Left(txtResult, Len(txtResult) - 2)
    End If
    
errShow:
    Select Case Err.Number
        Case 0
            
        Case 6
            txtResult = "过程被终止，因为发生了错误：溢出，数目已超出本程序的处理范围。（代号：6）"
        Case Else
            txtResult = "过程被终止，因为发生了错误：未预料到的错误，描述为“" & Err.Description & "”（代号：" & Err.Number & "）。"
    End Select
    
    txtN.Enabled = True
    txtResult.Enabled = True
    cmdGo.Enabled = True
End Sub

Private Sub txtN_Change()
    txtAB = "2×2"
    txtResult = ""
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8, 48 To 57 '若不是输入的数字或退格
        
    Case Else
        KeyAscii = 0 '不让输入的键生效
    End Select
End Sub
