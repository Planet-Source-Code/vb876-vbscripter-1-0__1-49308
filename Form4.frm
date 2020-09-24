VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Error log"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1440
      Top             =   2520
   End
   Begin VB.ListBox List1l 
      Height          =   1815
      ItemData        =   "Form4.frx":0000
      Left            =   0
      List            =   "Form4.frx":0007
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form4.List1l.List = mainForm.List1.List
End Sub

Private Sub Form_Resize()
Form4.List1l.Width = Me.Width
Form4.List1l.Height = Me.Height
End Sub

Private Sub List1l_Click()

End Sub

Private Sub Timer1_Timer()
Form4.List1l.List = mainForm.List1.List
End Sub
