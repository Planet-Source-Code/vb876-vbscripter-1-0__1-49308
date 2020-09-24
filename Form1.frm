VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainForm 
   Caption         =   "VBSCRIPTER PRO"
   ClientHeight    =   6045
   ClientLeft      =   3750
   ClientTop       =   3150
   ClientWidth     =   10275
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":038A
   ScaleHeight     =   6045
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1920
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1160
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1868
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2678
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3110
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3494
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3818
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2430
      ItemData        =   "Form1.frx":42A4
      Left            =   2160
      List            =   "Form1.frx":42A6
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2730
      ItemData        =   "Form1.frx":42A8
      Left            =   2160
      List            =   "Form1.frx":42AF
      TabIndex        =   8
      Top             =   45
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   360
      Width           =   495
      Begin VB.Shape Shp 
         BorderWidth     =   2
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.TextBox stringbox3 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Text            =   ".hta"
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox stringbox2 
      Height          =   405
      Left            =   6240
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox stringbox 
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox end 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Text            =   $"Form1.frx":42BE
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox start 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Text            =   $"Form1.frx":42CD
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   1440
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4920
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox txtKeyWords 
      Height          =   1965
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":42ED
      Top             =   2400
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   2760
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   661
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "New File"
            Object.ToolTipText     =   "New File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Open File"
            Object.ToolTipText     =   "Open File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Save File"
            Object.ToolTipText     =   "Save File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Test && Compile"
            Object.ToolTipText     =   "Test && Compile"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "My Clipboard"
            Object.ToolTipText     =   "My Clipboard"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Plugins"
            Object.ToolTipText     =   "Plugins"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Insert Output"
            Object.ToolTipText     =   "Insert Output"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Tools"
            Object.ToolTipText     =   "Tools"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Error Log"
            Object.ToolTipText     =   "Error Log"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Run"
            Object.ToolTipText     =   "Run"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Test && Compile"
            Object.ToolTipText     =   "Test && Compile"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "***Customize Toolbar***"
            Object.ToolTipText     =   "Customize Toolbar"
            ImageIndex      =   16
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin RichTextLib.RichTextBox Text1 
      DragIcon        =   "Form1.frx":45C6
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":4950
      MouseIcon       =   "Form1.frx":4A3C
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
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2880
      Picture         =   "Form1.frx":4D56
      Top             =   4800
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   """"
      Height          =   855
      Left            =   1200
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label rpl2 
      Caption         =   """)"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label rpl1 
      Caption         =   "("""
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   4440
      Width           =   615
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu fileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu fileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu fileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu fileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu fileSaveas 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu hhhhhhhhhhh 
         Caption         =   "ghg"
         Visible         =   0   'False
      End
      Begin VB.Menu ruinex 
         Caption         =   "&Test && Compile [F6]"
      End
      Begin VB.Menu fileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu fileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&My clipboard"
      Begin VB.Menu vbsclip 
         Caption         =   "Open My Clipboard"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep99 
         Caption         =   "-"
      End
      Begin VB.Menu editUndo 
         Caption         =   "VBScripter: &Undo"
         Shortcut        =   {F4}
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu editCut 
         Caption         =   "VBScripter: Cu&t"
         Shortcut        =   ^F
      End
      Begin VB.Menu editCopy 
         Caption         =   "VBScripter: &Copy"
         Shortcut        =   ^G
      End
      Begin VB.Menu editPaste 
         Caption         =   "VBScripter: &Paste"
         Shortcut        =   ^H
      End
      Begin VB.Menu editSep2 
         Caption         =   "-"
      End
      Begin VB.Menu editSelectAll 
         Caption         =   "VBScripter: &Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu toolsRun 
         Caption         =   "&Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ruine 
         Caption         =   "&Test && Compile"
         Shortcut        =   {F6}
      End
      Begin VB.Menu debug 
         Caption         =   "&Debug"
         Begin VB.Menu rss 
            Caption         =   "&Run And Debug Selected Text"
         End
         Begin VB.Menu rada 
            Caption         =   "&Run And Debug All"
         End
      End
      Begin VB.Menu addscr 
         Caption         =   "Add &Script"
      End
      Begin VB.Menu gtli 
         Caption         =   "Goto &Line"
         Shortcut        =   ^Y
      End
      Begin VB.Menu toolsruninIE 
         Caption         =   "Compile &HTA"
      End
      Begin VB.Menu vbscomp 
         Caption         =   "Compile &VBS"
      End
      Begin VB.Menu cf 
         Caption         =   "&Clear Error Log"
      End
      Begin VB.Menu dddddddddddddd 
         Caption         =   "-"
      End
      Begin VB.Menu csztm 
         Caption         =   "&Customize Toolbar"
      End
      Begin VB.Menu fvdfg 
         Caption         =   "-"
      End
      Begin VB.Menu dsdsdscall 
         Caption         =   "Call Script"
         Shortcut        =   {F2}
      End
      Begin VB.Menu toolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu toolsEncode 
         Caption         =   "&Encode"
         Shortcut        =   ^E
      End
      Begin VB.Menu toolsDecode 
         Caption         =   "&Decode"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu werdg 
      Caption         =   "&Error Log"
   End
   Begin VB.Menu beat 
      Caption         =   "&Plugins"
      Begin VB.Menu enaku 
         Caption         =   "Plugins..."
      End
      Begin VB.Menu dvaka 
         Caption         =   "Insert Output"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu debmso 
         Caption         =   "&Mouse Commands"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GEFIRSTVISIBLELINE = &HCE

Dim BL As Long
Dim B  As Long
Private Sub about_Click()
    frmAbout.Show 1
End Sub
Private Sub encode_Click()
    Text1.text = Encode(Text1.text)
End Sub

Private Sub Command2_Click()
On Error GoTo errrrrr
ScriptControl1.AddCode Text1.text
Exit Sub


errrrrr:

SetCursorAtLine Val(ScriptControl1.Error.Line), Text1
Text1.SetFocus
Text1.SelColor = vbRed
Text1.SelText = "'" + Err.Description + " > "


End Sub

Private Sub dva_Click()


SetCursorAtLine Val(ScriptControl1.Error.Line), Text1
Text1.SetFocus
Text1.SelColor = vbRed
Text1.SelText = "'" + Err.Description + " > "

End Sub

Private Sub addscr_Click()
Form3.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cf_Click()
List1.Clear
End Sub

Private Sub csztm_Click()
Toolbar1.Customize
End Sub

Private Sub debmso_Click()
Dialog.Show
End Sub

Private Sub dsdsdscall_Click()
Dim xcv

xcv = InputBox("Enter the name!", "Call")
Text1.SelText = Chr$(13) + Chr$(10) + "call " + xcv + Chr$(13) + Chr$(10)

Colorize Me.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
End Sub

Private Sub dvaka_Click()
editPaste_Click
End Sub

Private Sub editCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
End Sub
Private Sub editCut_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
End Sub
Private Sub editPaste_Click()
    Text1.SelText = Clipboard.GetText
    
    Colorize mainForm.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
End Sub

Private Sub editSelectAll_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.text)
End Sub

Private Sub editUndo_Click()
    UNDO
End Sub

Private Sub ena_Click()
On Error GoTo errrrrr
ScriptControl1.AddCode Text1.SelText
Exit Sub


errrrrr:

SetCursorAtLine Val(ScriptControl1.Error.Line), Text1
Text1.SetFocus
Text1.SelColor = vbRed
Text1.SelText = "'" + Err.Description + " > "

End Sub

Private Sub errlog_Click()
Form4.Show
End Sub

Private Sub enaku_Click()
Form7.Show
End Sub

Private Sub fileExit_Click()
    fileNew_Click
Unload Me
End Sub

Private Sub fileNew_Click()
On Error Resume Next
Dim tempscript, yn
tempscript = Environ("temp") & "\tempscript.vbs"
Kill tempscript
If txtChange = True Then
    yn = MsgBox("Save changes?", vbInformation + vbYesNo, "-VBScripter-")
    If yn = 6 Then
        Save
    End If
End If
    Text1.text = ""
    FileName = ""
    txtChange = ""
    Me.Caption = "VBScripter"
End Sub

Private Sub fileOpen_Click()
    Me.MousePointer = 11
    OpenFile
    Me.MousePointer = 0
    Colorize mainForm.Text1, &H8000&, vbRed, vbBlue
End Sub

Private Sub fileSave_Click()
    Save
End Sub

Private Sub fileSaveas_Click()
    SaveAs
End Sub
Function Encode(ArgChr)
Dim i, var1, var2

On Error Resume Next
For i = 1 To Len(ArgChr)
    var1 = Mid(ArgChr, i, 1)
    If Asc(var1) = 13 Then
        var2 = 15
    Else
        var2 = Asc(var1) + 100
    End If
    Encode = Encode & Chr(var2)
Next
End Function

Function Decode(ArgChr)
Dim num, i, var1, var2

On Error Resume Next
num = Val(Chr(49) & Chr(48) & Chr(48))
For i = 1 To Len(ArgChr)
    var1 = Mid(ArgChr, i, 1)
    If Asc(var1) = 15 Then
        Decode = Decode & Chr(13)
    Else
        var2 = Asc(var1) - num
        Decode = Decode & Chr(var2)
    End If
Next
End Function

Private Sub Form_Click()
List2.Visible = False
List1.Visible = False
End Sub

Private Sub Form_Load()


On Error Resume Next
If App.PrevInstance = True Then End
'If LCase(Command) <> "vbse007" Then End
fileNew_Click
  KeyWords = txtKeyWords.text
  
  
Set Pic.Font = Text1.Font
Pic.Print "a"
Pic.Print "a"

BL = Pic.CurrentY - Pic.TextHeight("a")
Shp.Height = Pic.TextHeight("a") / 1.5
B = Shp.Height / 2
Shp.Top = BL

Text1_SelChange
Pic.Cls
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Width = Me.Width - 610
Text1.Height = Me.Height - 1050
Toolbar1.Width = Me.Width - 140
Pic.Height = Me.Height - 1050

If mainForm.Width <= 10155 Then
mainForm.Width = 10155
End If

If mainForm.Height <= 3585 Then
mainForm.Height = 3585
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    fileExit_Click
    End
End Sub

Private Sub gtli_Click()
Dim LineL
LineL = InputBox("Enter the line number!", "Goto Line")
On Error Resume Next
SetCursorAtLine Val(LineL), Text1
Text1.SetFocus
End Sub

Private Sub hhhhhhhhhhh_Click()
SaveHtml
End Sub

Private Sub List1_DblClick()
If List1.Width <= 3615 Then
Timer3.Enabled = True
Else
List1.Width = 3615
End If

If List2.Width <= 3615 Then
Timer3.Enabled = True
Else
List2.Width = 3615
End If
End Sub

Private Sub List1_LostFocus()
List2.Visible = False
List1.Visible = False
End Sub

Private Sub rada_Click()
On Error GoTo errrrrr
ScriptControl1.AddCode Text1.text
Exit Sub


errrrrr:

SetCursorAtLine Val(ScriptControl1.Error.Line), Text1
Text1.SetFocus
Text1.SelColor = vbRed
Text1.SelText = "'" + Err.Description + " > "

List1.AddItem Err.Description

End Sub

Private Sub rss_Click()
On Error GoTo errrrrr
ScriptControl1.AddCode Text1.SelText
Exit Sub


errrrrr:

SetCursorAtLine Val(ScriptControl1.Error.Line), Text1
Text1.SetFocus
Text1.SelColor = vbRed
Text1.SelText = "'" + Err.Description + " > "

List1.AddItem Err.Description

End Sub

Private Sub ruine_Click()
'    On Error Resume Next
 '   tempscript = Environ("temp") & "\tempscript.vbs"
  '  Open tempscript For Output As #1
   '     Print #1, Text1.Text
    'Close #1
     'nResult = Shell("start.exe " & tempscript, vbHide)
     
     Form5.Show
End Sub

Private Sub ruinex_Click()
'    On Error Resume Next
 '   tempscript = Environ("temp") & "\tempscript.vbs"
  '  Open tempscript For Output As #1
   '     Print #1, Text1.Text
    'Close #1
     'nResult = Shell("start.exe " & tempscript, vbHide)
     
     Form5.Show
End Sub

Private Sub Text1_Change()
    txtChange = True
List2.Visible = False
List1.Visible = False
End Sub

Private Sub Text1_Click()
List2.Visible = False
List1.Visible = False
End Sub

Private Sub Text1_DblClick()
If Text1.text = "666" Then
Text1.text = "999"
Else
If Text1.text = "999" Then
Text1.text = "666"
End If
End If

If Text1.text = "App.Title" Then
Text1.text = "App.Title = " + Label1 + "VBScripter" + Label1
End If

If Text1.text = "App.WhoMadeThis" Then
Text1.text = "App.WhoMadeThis = " + Label1 + "Jan Robas" + Label1
End If

If Text1.text = "Crash" Then
Text1.text = Label1.Caption
Colorize mainForm.Text1, &H8000&, vbRed, &HFF0000
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

 If KeyAscii = vbKeyReturn Then Colorize mainForm.Text1, &H8000&, vbRed, &HFF0000
 Text1.SelColor = vbBlack
  
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim text, txt



If Button = 2 Then
If Text1.SelLength = 0 Then
Form3.Show
Else
rss_Click
End If
End If
End Sub

Private Sub Text1_SelChange()
Dim CurrIndex As Long
Dim indx As Long

CurrIndex = SendMessage(Text1.hWnd, EM_GEFIRSTVISIBLELINE, 0, 0)
indx = SendMessage(Text1.hWnd, EM_LINEFROMCHAR, -1, 0&)
Shp.Top = BL * (indx - CurrIndex) + B
Pic.CurrentY = 0
Pic.Cls
End Sub

Private Sub Timer1_Timer()
If InStr(1, Text1.text, "ahahahahah") > 1 Then
    toolsEncode.Enabled = True
    toolsDecode.Enabled = True
Else
    toolsEncode.Enabled = True
    toolsDecode.Enabled = True
End If
If Len(Text1.SelText) > 0 Then
    editCut.Enabled = True
    editCopy.Enabled = True
Else
    editCut.Enabled = False
    editCopy.Enabled = False
End If
If Len(Text1.text) > 0 Then
    tools.Enabled = True
Else
    tools.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If List1.Top <= 290 Then
List1.Top = List1.Top + 280
Else
Timer1.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If List1.Width < 7335 Then
List1.Width = List1.Width + 100
Else
Timer3.Enabled = False
End If

If List2.Width < 7335 Then
List2.Width = List1.Width + 100
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
List1.Visible = False
Timer4.Enabled = False
Timer4.Interval = 2000
End Sub

Private Sub Timer5_Timer()
Text1.text = Text1.text + "oOo"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
fileNew_Click

Case 2
fileOpen_Click

Case 3
fileSave_Click

Case "5", "20"
ruinex_Click

Case 7
editUndo_Click

Case 8
editCut_Click

Case 9
editCopy_Click

Case 10
editPaste_Click

Case 11
vbsclip_Click



Case 13
Form7.Show

Case 14
    editPaste_Click
    
    Case 16
    PopupMenu tools
    
    Case 18
    werdg_Click
    
    Case 19
    toolsRun_Click
    
    Case 24
    csztm_Click
End Select
End Sub

Private Sub toolsDecode_Click()
Dim ab, ba, aa

toolsRun.Enabled = True
On Error GoTo errmsg
    ab = InStr(1, Text1.text, "ahahahahah(") + 12
    ba = InStr(1, Text1.text, ")")
    aa = Mid(Text1.text, ab, (ba - ab))
    Text1.text = Decode(aa)
    Exit Sub
errmsg:
    MsgBox "Nothing to decode!", vbExclamation, "-VBScripter-"
    

End Sub
Private Sub toolsEncode_Click()
MsgBox "If you use encryption, then the RUN command won't work. To test click on 'Test & Compile' and then click on 'test vbs.vbs'! To compile click on 'compile vbs.vbs' or save file!", vbInformation, "Encryption"

Dim var_stg

    var_stg = "Function ahahahahah(ahahahah): On Error Resume Next: ahahah = Chr(49) & Chr(48) & Chr(48):For I = 1 To Len(ahahahah):ahah = Mid(ahahahah, I, 1):If Asc(ahah) = 15 Then:ahahahahah = ahahahahah & Chr(13):Else:ah = Asc(ahah) - ahahah:ahahahahah = ahahahahah & Chr(ah):End If:Next:End Function"
    Text1.text = "Execute ahahahahah(" & Chr(34) & Encode(Text1.text) & Chr(34) & ")" & vbCrLf & var_stg


End Sub
Private Sub toolsRun_Click()
Dim erline
'    On Error Resume Next
'    tempscript = Environ("temp") & "\tempscript.vbs"
'    Open tempscript For Output As #1
'        Print #1, Text1.Text
'    Close #1
'    nResult = Shell("start.exe " & tempscript, vbHide)

On Error GoTo errrrrr
ScriptControl1.AddCode Text1.text
Exit Sub


errrrrr:
'napaka:___start_of_error_script__
erline = ScriptControl1.Error.Line
SetCursorAtLine Val(ScriptControl1.Error.Line), Text1
Text1.SetFocus
Text1.SelColor = vbRed
Text1.SelText = "'" + Err.Description + " > "

List1.AddItem Err.Description

End Sub

Private Sub toolsruninIE_Click()
Dim sfile
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

Private Sub vbsclip_Click()
Form2.Show
End Sub

Private Sub vbscomp_Click()
    SaveAs
End Sub

Private Sub werdg_Click()
List2.Visible = True
List1.Visible = True
List1.Top = -3000
Timer1.Enabled = True
End Sub
