Sub NamesSearch()

Dim a, b, c, d, e, q, s, t, u, v, w, x, y, z, FirstNameRow1, FirstNameRow2, LastNameRow1, LastNameRow2 As Long
Dim FNPopScore, LNPopScore, Score1, Score2, FirstNameScore, LastNameScore, AvgScore As Long
Dim Data, Term, Alphabet, Breaks, Letter, Name1, Name2, SingleName, FullName As String
Dim Words() As Variant
Dim NameIndex() As Variant
Dim CapitalCheck As Boolean

Sheets("Main").Select
Alphabet = "abcdefghijklmnopqrstuvwxyz"
q = """"
Breaks = ".?!,'" & q
CapitalCheck = True

Sheets("Main").Range("A2:A" & Rows.Count).Font.Bold = False
Sheets("Main").Range("A2:A" & Rows.Count).Font.ColorIndex = xlAutomatic
Sheets("Main").Range("A2:A" & Rows.Count).Font.TintAndShade = 0
Sheets("Main").Range("B2:B" & Rows.Count).ClearContents

ThisWorkbook.Application.ScreenUpdating = False

'Index
ReDim NameIndex(26, 4)
For x = 1 To 26
    Letter = Mid(Alphabet, x, 1)
    NameIndex(x, 1) = Application.WorksheetFunction.Match(Letter & "1", Sheets("Names Data").Range("D:D"), 0)
    NameIndex(x, 2) = Application.WorksheetFunction.Match(Letter & "2", Sheets("Names Data").Range("D:D"), 0)
    NameIndex(x, 3) = Application.WorksheetFunction.Match(Letter & "1", Sheets("Names Data").Range("I:I"), 0)
    NameIndex(x, 4) = Application.WorksheetFunction.Match(Letter & "1", Sheets("Names Data").Range("I:I"), 0)
Next x

y = Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
If y > 501 Then
    y = 501
End If
For x = 2 To y
    'Data Breakdown
    Data = Sheets("Main").Range("A" & x).Value & " "
    t = Len(Data)
    
    ReDim Words(3, 1000)
    a = 1
    b = 1
    c = Len(Data)
    d = 0

    Do Until a > c
        Do Until InStr(1, Alphabet, LCase(Mid(Data, a, 1))) > 0 Or a > c
            If InStrRev(Breaks, LCase(Mid(Data, a, 1))) > 0 Then
                d = d + 1
                Words(1, d) = "xx"
            End If
            a = a + 1
        Loop
        b = a
        Do Until InStr(1, Alphabet, LCase(Mid(Data, b + 1, 1))) = 0 Or b > c
            b = b + 1
        Loop
        
        Term = Trim(Mid(Data, a, b - a + 1))
        If Len(Term) > 0 And Term <> " " Then
            d = d + 1
            Words(1, d) = Term
        End If
        
        a = b
        a = a + 1
    Loop

    'Data Name Lookup
    For e = 1 To d
        Term = Words(1, e)
        'Sheets("Main").Range("E" & e + 1).Value = Term
        Letter = Left(Term, 1)
        If InStr(1, Alphabet, LCase(Letter)) > 0 Then
            FirstNameRow1 = Application.WorksheetFunction.Match(Letter & "1", Sheets("Names Data").Range("D:D"), 0)
            FirstNameRow2 = Application.WorksheetFunction.Match(Letter & "2", Sheets("Names Data").Range("D:D"), 0)
            LastNameRow1 = Application.WorksheetFunction.Match(Letter & "1", Sheets("Names Data").Range("I:I"), 0)
            LastNameRow2 = Application.WorksheetFunction.Match(Letter & "2", Sheets("Names Data").Range("I:I"), 0)
            
            If IsError(Application.Match(LCase(Term), Sheets("Names Data").Range("A" & FirstNameRow1 & ":A" & FirstNameRow2), 0)) = False And _
                IsError(Application.Match(LCase(Term), Sheets("Names Data").Range("L:L"), 0)) = True Then
                    FNPopScore = Application.WorksheetFunction.VLookup(LCase(Term), Sheets("Names Data").Range("A" & FirstNameRow1 & ":C" & FirstNameRow2), 3, False)
            Else
                FNPopScore = 0
            End If
            If CapitalCheck = True And Letter <> UCase(Letter) Then
                FNPopScore = 0
            End If
            'Sheets("Main").Range("F" & e + 1).Value = FNPopScore
            
            If IsError(Application.Match(LCase(Term), Sheets("Names Data").Range("F" & LastNameRow1 & ":F" & LastNameRow2), 0)) = False And _
                IsError(Application.Match(LCase(Term), Sheets("Names Data").Range("L:L"), 0)) = True Then
                    LNPopScore = Application.WorksheetFunction.VLookup(LCase(Term), Sheets("Names Data").Range("F" & LastNameRow1 & ":H" & LastNameRow2), 3, False)
            Else
                LNPopScore = 0
            End If
            If CapitalCheck = True And Letter <> UCase(Letter) Then
                LNPopScore = 0
            End If
            'Sheets("Main").Range("G" & e + 1).Value = LNPopScore
            
            If FNPopScore > 0 Then
                Words(2, e) = FNPopScore
                Words(3, e) = "X"
            End If
            
            If LNPopScore > 0 Then
                If Len(Words(2, e)) = 0 Or LNPopScore < FNPopScore Then
                    Words(2, e) = LNPopScore
                End If
            End If
        End If
    Next e
    
    'Check for Full Names
    'PopScore Max
    FirstNameScore = 80
    LastNameScore = 80
    
    For e = 1 To d - 1
        Name1 = Words(1, e)
        Name2 = Words(1, e + 1)
        Score1 = Words(2, e)
        Score2 = Words(2, e + 1)
                
        'FirstName LastName
        If Score1 > 0 And Score2 > 0 And Words(3, e) = "X" And Len(Name2) > 2 Then
            If Score1 <= FirstNameScore Or Score2 <= LastNameScore Then
                FullName = Name1 & " " & Name2
                AvgScore = 100 - Round((Score1 + Score2) / 2, 1)
                If IsEmpty(Sheets("Main").Range("B" & x)) = True Then
                    Sheets("Main").Range("B" & x).Value = FullName & " (" & AvgScore & "%)"
                Else
                    Sheets("Main").Range("B" & x).Value = Sheets("Main").Range("B" & x).Value & ", " & FullName & " (" & AvgScore & "%)"
                End If

                Call ColorName(x, FullName, vbRed)
            End If
        End If
    Next e

    'Counter
    Sheets("Main").Range("D1").Value = x & " of " & y
    If x Mod 500 = 0 Then
        ThisWorkbook.Application.ScreenUpdating = True
        ThisWorkbook.Application.ScreenUpdating = False
    End If
Next x
Sheets("Main").Range("D1").ClearContents
MsgBox "Done!", vbInformation, "Done!"

End Sub

Function ColorName(NameRow, NameText, NameColor As Long)

Dim w, x, y, z As Long
Dim Data As String

Data = Sheets("Main").Range("A" & NameRow).Value
y = Len(Data)
For x = 1 To y
    If Mid(LCase(Data), x, Len(NameText)) = LCase(NameText) Then
        With Sheets("Main").Range("A" & NameRow).Characters(Start:=x, Length:=Len(NameText)).Font
            .FontStyle = "Bold"
            .Color = NameColor
            .TintAndShade = 0
        End With
    End If
Next x

End Function
