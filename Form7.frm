VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plugins"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5415
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   8
      Text            =   "plugin"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4800
      Picture         =   "Form7.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Refresh"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4920
      Picture         =   "Form7.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add script"
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4920
      Picture         =   "Form7.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Clear output"
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4920
      Picture         =   "Form7.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "My Clipboard"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Run Plugin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   4200
      Picture         =   "Form7.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Paste Output && Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Archive         =   0   'False
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2190
      Left            =   120
      Pattern         =   "*.EXE"
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mainForm.Text1.SelText = Clipboard.GetText
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Shell App.Path + "\" + Text1.Text + "\" + File1.FileName + "", vbHide
End Sub

Private Sub Command4_Click()
outputman.Show
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
Clipboard.SetText ""
End Sub

Private Sub Command7_Click()
Form3.Show
Form3.Text1.Text = Clipboard.GetText
Unload Me
End Sub

Private Sub Command8_Click()
File1.Path = App.Path + "\" + Text1.Text + "\"
File1.Refresh
End Sub

Private Sub Command9_Click()

End Sub

Private Sub File1_DblClick()
Shell App.Path + "\" + Text1.Text + "\" + File1.FileName + "", vbHide
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\" + Text1.Text + "\"
End Sub

