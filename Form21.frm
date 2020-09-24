VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "VBScripter"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   2535
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   1440
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   5640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   5805
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Verifying VBScripter version 1.0..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   5805
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   0
      Picture         =   "Form21.frx":0000
      Top             =   240
      Width           =   5805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If App.CompanyName = "Avtor: Jan Robas (JanSoft)" Then
Else
MsgBox "Wrong version!", vbCritical, "VBSCRIPTER"
End
End If

If App.ProductName = "VBSCRIPTER" Then
Else
MsgBox "Wrong version!", vbCritical, "VBSCRIPTER"
End
End If

If App.Comments = "s040290308" Then
Else
MsgBox "Wrong version!", vbCritical, "VBSCRIPTER"
End
End If


Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Unload Me
mainForm.Show
End Sub
