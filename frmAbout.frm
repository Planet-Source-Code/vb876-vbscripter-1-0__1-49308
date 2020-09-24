VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VBScripter"
   ClientHeight    =   3345
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6015
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2308.778
   ScaleMode       =   0  'User
   ScaleWidth      =   5648.396
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ANTI encryption"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7200
      TabIndex        =   3
      Top             =   2400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   585
      Left            =   2505
      Picture         =   "frmAbout.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1905
      Left            =   105
      Picture         =   "frmAbout.frx":0396
      Top             =   60
      Width           =   5805
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Author: Jan Robas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   2220
      Width           =   5925
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0 professional"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   1980
      Width           =   5805
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
If Check1.Value = 0 Then
mainForm.toolsDecode.Visible = False
mainForm.toolsEncode.Visible = False
Else
mainForm.toolsDecode.Visible = True
mainForm.toolsEncode.Visible = True
End If
  'uc
  Unload Me
End Sub

