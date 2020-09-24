Attribute VB_Name = "Module1"
Const EM_UNDO = &HC7
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public FileName, txtChange
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Dim OFName As OPENFILENAME
Public Function ShowOpen() As String
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = mainForm.hWnd
    'OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Visual Basic Script (*.VBS)" + Chr$(0) + "*.VBS" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
Public Function ShowSave() As String
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = mainForm.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Type the file name without extension! (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    'OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Save File"
    OFName.flags = 0
    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function
Sub OpenFile()
    sfile = ShowOpen
    If sfile <> "" Then
    On Error Resume Next
    FileName = Mid(sfile, 1, Len(sfile) - 1)
        If Len(FileName) > 0 Then
        mainForm.Caption = "VBScripter PRO - " & FileName
        Open FileName For Input As #1
            mainForm.Text1.Text = ""
            Do While Not EOF(1)
                Dim MyString
                Line Input #1, MyString
                MyString = MyString & vbCrLf
                mainForm.Text1.Text = mainForm.Text1.Text & MyString
            Loop
        Close #1
        FileName = sfile
        mainForm.Text1.SelStart = Len(mainForm.Text1.Text)
        End If
    End If
End Sub
Sub SaveAs()
On Error Resume Next
sfile = ShowSave
If Len(sfile) > 0 Then
sfile = Mid(sfile, 1, Len(sfile) - 1)
If LCase(Right(sfile, 4)) <> ".vbs" Then sfile = sfile & ".vbs"
    If Len(Dir(sfile)) > 1 Then
        yn = MsgBox("That file already exist. Overwrite it!", vbInformation + vbYesNoCancel, "-VBScripter-")
            If yn = 6 Then
                FileName = sfile
                mainForm.Caption = "VBScripter PRO - " & FileName
                txtChange = ""
                    Open FileName For Output As #1
                        Print #1, mainForm.Text1.Text
                    Close #1
                Exit Sub
            ElseIf yn = 7 Then
                SaveAs
            End If
    Exit Sub
    End If
    FileName = sfile
    mainForm.Caption = "VBScripter PRO - " & FileName
    txtChange = ""
    Open FileName For Output As #1
        Print #1, mainForm.Text1.Text
    Close #1
End If
End Sub

Sub SaveHtml()
On Error Resume Next
sfile = ShowSave
If Len(sfile) > 0 Then
sfile = Mid(sfile, 1, Len(sfile) - 1)
If LCase(Right(sfile, 4)) <> ".vbs" Then sfile = sfile & ".vbs"
    If Len(Dir(sfile)) > 1 Then
        yn = MsgBox("That file already exist. Overwrite it!", vbInformation + vbYesNoCancel, "-VBScripter-")
            If yn = 6 Then
                FileName = sfile
                mainForm.Caption = "VBScripter PRO - " & FileName
                txtChange = ""
                    Open FileName For Output As #1
                        Print #1, mainForm.start.Text + mainForm.Text1.Text + mainForm.end.Text
                    Close #1
                Exit Sub
            ElseIf yn = 7 Then
                SaveAs
            End If
    Exit Sub
    End If
    FileName = sfile
    mainForm.Caption = "VBScripter PRO - " & FileName
    txtChange = ""
    Open FileName For Output As #1
        Print #1, Form.start.Text + mainForm.Text1.Text + mainForm.end.Text
    Close #1
End If
End Sub


Sub UNDO()
    SendMessage mainForm.Text1.hWnd, EM_UNDO, 0&, 0&
End Sub
Sub Save()
If Len(FileName) > 0 Then
    txtChange = ""
    Open FileName For Output As #1
        Print #1, mainForm.Text1.Text
    Close #1
Else
    SaveAs
End If
End Sub
