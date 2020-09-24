VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form eroror 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generator"
   ClientHeight    =   540
   ClientLeft      =   2670
   ClientTop       =   4215
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2025
   Begin VB.CommandButton Command1 
      Caption         =   "&Go!"
      Height          =   255
      Left            =   10800
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   2
      Text            =   "1"
      Top             =   4560
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2295
      Left            =   12360
      TabIndex        =   0
      Top             =   5280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"errrr.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Goto line:"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   4560
      Width           =   2295
   End
End
Attribute VB_Name = "eroror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

SetCursorAtLine Val(Text1), RichTextBox1

RichTextBox1.SetFocus
End Sub


Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, _
    (Screen.Height - Height) / 2

For x = 1 To 10
    RichTextBox1.text = RichTextBox1.text & "Line " & x & vbCrLf
Next

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

    Select Case Chr(KeyAscii)
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
    Case Else
        If KeyAscii = 13 Then
            Command1.Value = True
        ElseIf KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End Select
End Sub


