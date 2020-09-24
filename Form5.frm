VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test & Compile"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5670
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":1C32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   128
      TabIndex        =   11
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2143
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2108
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Duble-click on any icon!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tip: HTA file acts like EXE, but its in HTML language. If you want to compile HTA, use vbscript and document.write!"
      Height          =   495
      Left            =   330
      TabIndex        =   12
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Compile hta.hta"
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
      Left            =   330
      TabIndex        =   10
      Top             =   2655
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Compile HTA file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   9
      Top             =   2655
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2010
      TabIndex        =   8
      Top             =   2655
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   975
      Left            =   10800
      Picture         =   "Form5.frx":2886
      Top             =   960
      Width           =   810
   End
   Begin VB.Image Image7 
      Height          =   960
      Left            =   9960
      Picture         =   "Form5.frx":526C
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   6
      Top             =   1815
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2010
      TabIndex        =   5
      Top             =   2415
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2010
      TabIndex        =   4
      Top             =   2175
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Compile VBS file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2325
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Compile vbs.vbs"
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
      Left            =   330
      TabIndex        =   2
      Top             =   2415
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Simulate compiled vbs script"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2325
      TabIndex        =   1
      Top             =   2175
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test vbs.vbs"
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
      Left            =   330
      TabIndex        =   0
      Top             =   2175
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   9840
      Picture         =   "Form5.frx":7BAE
      Top             =   1560
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   8760
      Picture         =   "Form5.frx":A0B0
      Top             =   1560
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   870
      Left            =   10560
      Picture         =   "Form5.frx":C5B2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9360
      Picture         =   "Form5.frx":FB6C
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
ListView1.ListItems.Add , , "Test vbs.vbs", 1
ListView1.ListItems.Add , , "Compile vbs.vbs", 2
ListView1.ListItems.Add , , "Compile hta.hta", 3
End Sub

Private Sub Form_LostFocus()
Form5.Show
End Sub

Private Sub Image3_Click()
Image3.Picture = Image2.Picture
Image6.Picture = Image4.Picture
End Sub

Private Sub Image3_DblClick()
Dim tempscript

    On Error Resume Next
    tempscript = "C:\tempscript.vbs"
    Open tempscript For Output As #1
        Print #1, mainForm.Text1.text
    Close #1
    nResult = Shell("start.exe " & tempscript, vbHide)
End Sub

Private Sub Image6_Click()
Image6.Picture = Image5.Picture
Image3.Picture = Image1.Picture
Image9.Picture = Image8.Picture
End Sub

Private Sub Image6_DblClick()
    SaveAs
End Sub

Private Sub Image9_Click()
Image9.Picture = Image7.Picture
Image3.Picture = Image1.Picture
Image6.Picture = Image4.Picture
End Sub

Private Sub Image9_DblClick()
sfile = ShowSave
mainForm.stringbox = sfile
mainForm.stringbox2 = mainForm.stringbox + mainForm.stringbox3

Dim txtHTML, tempHTML
txtHTML = "<HTML>" & vbCrLf
txtHTML = txtHTML & "<HTML>" & vbCrLf
txtHTML = txtHTML & "<TITLE>HTA file</TITLE>" & vbCrLf
txtHTML = txtHTML & "<HEAD>" & vbCrLf
txtHTML = txtHTML & "<SCRIPT LANGUAGE=VBSCRIPT>" & vbCrLf
txtHTML = txtHTML & mainForm.Text1.text
txtHTML = txtHTML & "</SCRIPT>" & vbCrLf
txtHTML = txtHTML & "</HEAD>" & vbCrLf
txtHTML = txtHTML & "</HTML>" & vbCrLf
    On Error Resume Next
    tempHTML = mainForm.stringbox2
    Open tempHTML For Output As #1
        Print #1, txtHTML
    Close #1
End Sub

Private Sub ListView1_DblClick()
Select Case ListView1.SelectedItem.Index

Case 1
Dim tempscript

    On Error Resume Next
    tempscript = "C:\tempscript.vbs"
    Open tempscript For Output As #1
        Print #1, mainForm.Text1.text
    Close #1
    nResult = Shell("start.exe " & tempscript, vbHide)
    
    
Case 2
    SaveAs
    
Case 3
sfile = ShowSave
mainForm.stringbox = sfile
mainForm.stringbox2 = mainForm.stringbox + mainForm.stringbox3

Dim txtHTML, tempHTML
txtHTML = "<HTML>" & vbCrLf
txtHTML = txtHTML & "<HTML>" & vbCrLf
txtHTML = txtHTML & "<TITLE>HTA file</TITLE>" & vbCrLf
txtHTML = txtHTML & "<HEAD>" & vbCrLf
txtHTML = txtHTML & "<SCRIPT LANGUAGE=VBSCRIPT>" & vbCrLf
txtHTML = txtHTML & mainForm.Text1.text
txtHTML = txtHTML & "</SCRIPT>" & vbCrLf
txtHTML = txtHTML & "</HEAD>" & vbCrLf
txtHTML = txtHTML & "</HTML>" & vbCrLf
    On Error Resume Next
    tempHTML = mainForm.stringbox2
    Open tempHTML For Output As #1
        Print #1, txtHTML
    Close #1
    
    End Select
End Sub

Private Sub Picture1_Click()
Image3.Picture = Image1.Picture
Image6.Picture = Image4.Picture
Image9.Picture = Image8.Picture
End Sub

