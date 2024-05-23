Sub CompareTexts()
    
Dim a, b, c, s, t, v, w, x, y, z, ColumnNumber, TermsCount1, TermsCount2 As Long
Dim Data1, Data2, Word, TermsMatches1, TermsMatches2 As String
Dim Terms1() As Variant
Dim Terms1x() As Variant
Dim Terms2() As Variant
Dim Terms2x() As Variant
Dim Matches() As Variant

Sheets("Main").Select
Sheets("Main").Range("B2:C" & Rows.Count).ClearContents
Sheets("Main").Range("A2:A" & Rows.Count).Font.ColorIndex = xlAutomatic
Sheets("Main").Range("A2:A" & Rows.Count).Font.TintAndShade = 0
Sheets("Main").Range("A2:A" & Rows.Count).Font.Bold = False
Sheets("Main").Range("E2:F" & Rows.Count).ClearContents
Sheets("Main").Range("D2:D" & Rows.Count).Font.ColorIndex = xlAutomatic
Sheets("Main").Range("D2:D" & Rows.Count).Font.TintAndShade = 0
Sheets("Main").Range("D2:D" & Rows.Count).Font.Bold = False
ThisWorkbook.Application.ScreenUpdating = False

y = Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row

For x = 2 To y
    Data1 = Trim(Sheets("Main").Range("A" & x).Value) & " "
    Data2 = Trim(Sheets("Main").Range("D" & x).Value) & " "

    'Breakdown Data
    ReDim Terms1(500)
    ReDim Terms2(500)
    Terms1 = DataBreakdown(Data1)
    Terms2 = DataBreakdown(Data2)
    
    'Expand Terms
    ReDim Terms1x(5000)
    ReDim Terms2x(5000)
    Terms1x = DataExpand(Terms1)
    Terms2x = DataExpand(Terms2)

    'Compare Terms
    ReDim Matches(2, 1000)
    Matches = CompareData(Terms1x, Terms2x)
    
    TermsMatches1 = ""
    TermsMatches2 = ""
    TermsCount1 = 0
    TermsCount2 = 0
    z = Matches(0, 0)
    For w = 1 To z
        ColumnNumber = Matches(1, w)
        If ColumnNumber = 1 Then
            TermsCount1 = TermsCount1 + 1
            Word = Terms1(Matches(2, w))
        ElseIf ColumnNumber = 2 Then
            TermsCount2 = TermsCount2 + 1
            Word = Terms2(Matches(2, w))
        End If
        
        If Len(ColumnNumber) > 0 Then
            If ColumnNumber = 1 Then
                If Len(TermsMatches1) = 0 Then
                    TermsMatches1 = Word
                Else
                    TermsMatches1 = TermsMatches1 & ", " & Word
                End If
                Call Highlight(Word, Sheets("Main").Range("A" & x), vbRed)
            ElseIf ColumnNumber = 2 Then
                If Len(TermsMatches2) = 0 Then
                    TermsMatches2 = Word
                Else
                    TermsMatches2 = TermsMatches2 & ", " & Word
                End If
                Call Highlight(Word, Sheets("Main").Range("D" & x), vbRed)
            End If
            Sheets("Main").Range("B" & x).Value = TermsMatches1
            Sheets("Main").Range("E" & x).Value = TermsMatches2
        End If
    Next w
    
    'Score
    Sheets("Main").Range("C" & x).Value = TermsCount1 / Terms1(0)
    Sheets("Main").Range("F" & x).Value = TermsCount2 / Terms2(0)

    'Counter
    Sheets("Main").Range("G1").Value = x & " of " & y
    If x Mod 1000 = 0 Then
        ThisWorkbook.Application.ScreenUpdating = True
        ThisWorkbook.Application.ScreenUpdating = False
        'ThisWorkbook.Save
    End If
Next x
Sheets("Main").Range("G1").ClearContents
ThisWorkbook.Application.ScreenUpdating = True
MsgBox "Done!", vbInformation, "Done!"

End Sub

Function DataBreakdown(Data) As Variant

Dim a, b, c, d As Long
Dim Alpha, Term As String
Dim Words() As Variant

ReDim Words(500)

Alpha = "abcdefghijklmnopqrstuvwxyz0123456789"
a = 1
b = 1
c = Len(Data)
d = 1
Do Until a > c
    Do Until InStr(1, Alpha, LCase(Mid(Data, a, 1))) > 0 Or a > c
        a = a + 1
    Loop

    b = a
    Do Until InStr(1, Alpha, LCase(Mid(Data, b, 1))) = 0 Or b > c
        b = b + 1
    Loop
    
    Term = Trim(Mid(Data, a, b - a))
    If Len(Term) > 0 And Term <> " " Then
        Words(d) = Term
        d = d + 1
    End If
    a = b
    a = a + 1
Loop
Words(0) = d - 1
DataBreakdown = Words

End Function

Function DataExpand(Terms) As Variant

Dim w, x, y, z As Long
Dim LastLetter As String
Dim Term As Variant
Dim ExpandedTerms() As Variant

ReDim ExpandedTerms(2, 5000)
y = Terms(0)
z = 0
For x = 1 To y
    Term = LCase(Terms(x))
    LastLetter = Right(Term, 1)
    
    z = z + 1
    ExpandedTerms(1, z) = x
    ExpandedTerms(2, z) = Term
    z = z + 1
    ExpandedTerms(1, z) = x
    ExpandedTerms(2, z) = Term & "s"
    z = z + 1
    ExpandedTerms(1, z) = x
    ExpandedTerms(2, z) = Term & "es"
    
    If Len(Term) > 2 And Right(Term, 1) = "y" Then
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Left(Term, Len(Term) - 1) & "ies"
    End If
    If Len(Term) > 5 And Right(Term, 3) = "ies" Then
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Left(Term, Len(Term) - 3) & "y"
    End If
    If Len(Term) > 2 And Right(Term, 1) = "e" Then
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Left(Term, Len(Term) - 1) & "ing"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "d"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "r"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "rs"
    Else
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "ing"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & LastLetter & "ing"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "ed"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & LastLetter & "ed"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "er"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & LastLetter & "er"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & "ers"
        z = z + 1
        ExpandedTerms(1, z) = x
        ExpandedTerms(2, z) = Term & LastLetter & "ers"
    End If
Next x
ExpandedTerms(0, 0) = z
DataExpand = ExpandedTerms

End Function

Function CompareData(Terms1x, Terms2x) As Variant

Dim u, v, w, x, y, z As Long
Dim Term1, Term2 As Variant
Dim Terms1() As Variant
Dim Terms2() As Variant
Dim Matches() As Variant

ReDim Terms1(500)
ReDim Terms2(500)
ReDim Matches(500)

u = 0
y = Terms1x(0, 0)
z = Terms2x(0, 0)
For x = 1 To y
    Term1 = Terms1x(2, x)
    For w = 1 To z
        Term2 = Terms2x(2, w)
        If Term1 = Term2 Then
            u = u + 1
            Terms1(u) = Terms1x(1, x)
            Terms2(u) = Terms2x(1, w)
            Exit For
        End If
    Next w
Next x
Terms1 = DataDedupe(Terms1)
Terms2 = DataDedupe(Terms2)

ReDim Matches(2, 1000)
w = 1
y = UBound(Terms1)
For x = 0 To y
    If Len(Terms1(x)) > 0 Then
        Matches(1, w) = 1
        Matches(2, w) = Terms1(x)
        w = w + 1
    End If
Next x
y = UBound(Terms2)
For x = 0 To y
    If Len(Terms2(x)) > 0 Then
        Matches(1, w) = 2
        Matches(2, w) = Terms2(x)

        w = w + 1
    End If
Next x
Matches(0, 0) = w
CompareData = Matches

End Function

Function DataDedupe(Terms) As Variant

Dim w, x, y, z As Long
Dim Dict As Scripting.Dictionary

Set Dict = New Scripting.Dictionary

For x = 1 To UBound(Terms)
    If IsMissing(Terms(x)) = False Then
        Dict.Item(Terms(x)) = 1
    End If
Next x
DataDedupe = Dict.Keys


End Function

Function Highlight(Term, HRange As Range, HighlightColor As Long)

Dim w, x, y, z As Long
Dim Keyword As String

Alpha = "abcdefghijklmnopqrstuvwxyz0123456789"
Keyword = " " & LCase(Trim(HRange.Value)) & " "
Term = LCase(Term)

z = Len(Keyword)
For x = 2 To z
    If Term = Mid(Keyword, x, Len(Term)) And InStrRev(Alpha, Mid(Keyword, x - 1, 1)) = 0 And InStrRev(Alpha, Mid(Keyword, x + Len(Term), 1)) = 0 Then
        HRange.Characters(Start:=x - 1, Length:=Len(Term)).Font.Bold = True
        HRange.Characters(Start:=x - 1, Length:=Len(Term)).Font.Color = HighlightColor
    End If
Next x

End Function


