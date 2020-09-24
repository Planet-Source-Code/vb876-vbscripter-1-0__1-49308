VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add script"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6870
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Paste"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Text            =   "Script"
      Top             =   0
      Width           =   5655
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUB"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.TextBox txtKeyWords 
      Height          =   1365
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form3.frx":038A
      Top             =   6120
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6360
      Top             =   475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   4675
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4675
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4675
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4065
      Left            =   0
      TabIndex        =   3
      Top             =   435
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7170
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form3.frx":0663
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
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   70
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
If Check1.Value = 1 Then
mainForm.Text1.SelText = Chr$(13) + Chr$(10) + "" + Chr$(13) + Chr$(10) + "'Start of " + Text2.Text + Chr$(13) + Chr$(10) + "sub " + Text2.Text + Chr$(13) + Chr$(10) + Text1.Text + Chr$(13) + Chr$(10) + "end sub" + Chr$(13) + Chr$(10) + "" + "'End of " + Text2.Text + Chr$(13) + Chr$(10)

Unload Me
Else
mainForm.Text1.SelText = "" + Chr$(13) + Chr$(10) + "'Start of " + Text2.Text + Chr$(13) + Chr$(10) + Text1.Text + Chr$(13) + Chr$(10) + "" + "'End of " + Text2.Text + Chr$(13) + Chr$(10)

Unload Me
End If

      mainForm.Text1.SelStart = Len(Text1.Text)
Colorize mainForm.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
 
 mainForm.Text1.MousePointer = 0
End Sub

Private Sub Command3_Click()
Unload Me
 mainForm.Text1.MousePointer = 0
End Sub

Private Sub Command4_Click()
Text1.SelText = Clipboard.GetText
Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
End Sub

Private Sub Form_Load()
Timer1.Enabled = True

  KeyWords = txtKeyWords.Text
      Text1.SelStart = Len(Text1.Text)

Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack

End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
End Sub

