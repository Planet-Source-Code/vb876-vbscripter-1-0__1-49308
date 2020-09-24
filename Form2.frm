VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Clipboard"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6870
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKeyWords 
      Height          =   1365
      Left            =   9000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form2.frx":038A
      Top             =   3600
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6360
      Top             =   480
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4185
      Left            =   0
      TabIndex        =   0
      Top             =   320
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7382
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form2.frx":0663
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " You can edit My Clipboard contents by editing this text:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5
      Width           =   6855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.SetText ""
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Clipboard.SetText Text1.Text
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = Clipboard.GetText
Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Text1.Text = Clipboard.GetText
  KeyWords = txtKeyWords.Text
      Text1.SelStart = Len(Text1.Text)

Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack

End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End Sub

Private Sub Text1_Change()
Clipboard.SetText Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
End Sub
