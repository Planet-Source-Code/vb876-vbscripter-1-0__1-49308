VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Syntax Sample"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtKeyWords 
      Height          =   1365
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   2
      Text            =   "frmMain.frx":0000
      Top             =   5040
      Width           =   9135
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":02D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "KeyWord variables:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  KeyWords = txtKeyWords.Text
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
 '// Only Colorize-Keywords, when the user press the ReturnKey
 '// --> So whe save time
 If KeyAscii = vbKeyReturn Then Colorize frmMain.Text1, vbGreen, vbRed, vbBlue
 Text1.SelColor = vbBlack
End Sub
Private Sub txtKeyWords_Change()
 KeyWords = txtKeyWords.Text
End Sub


