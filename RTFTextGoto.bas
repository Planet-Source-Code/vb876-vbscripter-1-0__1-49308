Attribute VB_Name = "Module0"



Public Sub SetCursorAtLine(WhichLine As Long, WhichRTFText As RichTextBox)

Dim Estimate As Long, StartP As Long, EndP As Long
Dim NumChars As Long



With WhichRTFText

    NumChars = Len(.Text)

   
    If NumChars = 0 Then
        Exit Sub
    End If
    
    If WhichLine <= 1 Then
        .SelStart = 0
        .SelLength = 0
        Exit Sub
    ElseIf WhichLine > (.GetLineFromChar(NumChars) + 1) Then
        .SelStart = NumChars
        .SelLength = 0
        Exit Sub
    End If
        
    Estimate = Int(NumChars / 2)
    StartP = 1
    EndP = NumChars

    Dim Finalised As Long
    Do
        If WhichLine < (.GetLineFromChar(Estimate) + 1) Then
     
            StartP = StartP
            EndP = Estimate
            Estimate = StartP + Int((EndP - StartP) / 2)
        ElseIf WhichLine > (.GetLineFromChar(Estimate) + 1) Then
 
            StartP = Estimate
            EndP = EndP
            Estimate = StartP + Int((EndP - StartP) / 2)
        Else
            Finalised = Estimate
   
            Do
                Finalised = Finalised - 1
                If Finalised = 0 Then
                
                    .SelStart = Finalised
                    .SelLength = 0
                    Exit Do
                Else
                    If (.GetLineFromChar(Finalised) + 1) < WhichLine Then
                        Finalised = Finalised + 1
                        .SelStart = Finalised
                        .SelLength = 0
                        Exit Do
                    End If
                End If
            Loop
            Exit Do
        End If
    Loop
End With
End Sub
