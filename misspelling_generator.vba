Sub GenerateMisspellings()

Dim w, x, y, z As Long
Dim Phrase, Alphabet As String
Dim MissedKeys() As String

Alphabet = "abcdefghijklmnopqrstuvwxyz"
ReDim MissedKeys(26)
MissedKeys(1) = "qwsxz"
MissedKeys(2) = "vghn"
MissedKeys(3) = "xdfv"
MissedKeys(4) = "serfcx"
MissedKeys(5) = "w34rfds"
MissedKeys(6) = "drtgvc"
MissedKeys(7) = "ftyhbv"
MissedKeys(8) = "gyujnb"
MissedKeys(9) = "u89olkj"
MissedKeys(10) = "huikmn"
MissedKeys(11) = "jiolm"
MissedKeys(12) = "kop"
MissedKeys(13) = "njk"
MissedKeys(14) = "bhjm"
MissedKeys(15) = "i90plk"
MissedKeys(16) = "o0l"
MissedKeys(17) = "12wsa"
MissedKeys(18) = "e45tgfd"
MissedKeys(19) = "aqwdxz"
MissedKeys(20) = "r56yhgf"
MissedKeys(21) = "y78ikjh"
MissedKeys(22) = "cfgb"
MissedKeys(23) = "q23edsa"
MissedKeys(24) = "zsdc"
MissedKeys(25) = "t67ujhg"
MissedKeys(26) = "asx"

Sheets("List").Range("A2:D" & Rows.Count).Delete
y = Sheets("Main").Range("D" & Rows.Count).End(xlUp).Row

For x = 2 To y
    Phrase = Trim(Sheets("Main").Range("D" & x).Value)
    v = Sheets("List").Range("A" & Rows.Count).End(xlUp).Row + 1
    z = Len(Phrase)
    If Sheets("Main").Range("B2").Value = True Then
        Call SkippedLetters(Phrase, z, v)
    End If
    If Sheets("Main").Range("B3").Value = True Then
        Call DoubleLetters(Phrase, z, v)
    End If
    If Sheets("Main").Range("B4").Value = True Then
        Call ReverseLetters(Phrase, z, v)
    End If
    If Sheets("Main").Range("B5").Value = True Then
        Call SkipSpaces(Phrase, z, v)
    End If
    If Sheets("Main").Range("B6").Value = True Then
        Call MissedKey(Phrase, z, v, Alphabet, MissedKeys)
    End If
    If Sheets("Main").Range("B7").Value = True Then
        Call InsertedKey(Phrase, z, v, Alphabet, MissedKeys)
    End If
Next x
Sheets("List").Range("A:D").RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes

Sheets("Main").Range("A10").Value = "Spellchecking..."
y = Sheets("List").Range("A" & Rows.Count).End(xlUp).Row
For x = 2 To y
    If Application.CheckSpelling(Sheets("List").Range("C" & x).Value) = False Then
        Sheets("List").Range("D" & x).Value = "Misspelled"
    End If
Next x
Sheets("Main").Range("A10").ClearContents
Sheets("List").Select

End Sub

Function SkippedLetters(Phrase, z, v)

Dim w, x, y As Long
Dim Misspelling As String

Sheets("Main").Range("A10").Value = "Skipped Letters Generating..."
For w = 1 To z
    Sheets("List").Range("A" & v).Value = Phrase
    Sheets("List").Range("B" & v).Value = "Skipped Letters"
    Sheets("List").Range("C" & v).Value = Left(Phrase, w - 1) & Right(Phrase, Len(Phrase) - w)
    v = v + 1
Next w
Sheets("Main").Range("A10").ClearContents

End Function

Function DoubleLetters(Phrase, z, v)

Dim w, x, y As Long
Dim Letter As String

Sheets("Main").Range("A10").Value = "Double Letters Generating..."
For w = 1 To z
    Letter = Mid(Phrase, w, 1)
    If Letter <> " " Then
        Sheets("List").Range("A" & v).Value = Phrase
        Sheets("List").Range("B" & v).Value = "Double Letters"
        Sheets("List").Range("C" & v).Value = Left(Phrase, w) & Letter & Right(Phrase, Len(Phrase) - w)
        v = v + 1
    End If
Next w
Sheets("Main").Range("A10").ClearContents

End Function

Function ReverseLetters(Phrase, z, v)

Dim w, x, y As Long
Dim Letter1, Letter2 As String

Sheets("Main").Range("A10").Value = "Reverse Letters Generating..."
For w = 1 To z - 1
    Letter1 = Mid(Phrase, w, 1)
    Letter2 = Mid(Phrase, w + 1, 1)
    If Letter1 <> Letter2 Then
        Sheets("List").Range("A" & v).Value = Phrase
        Sheets("List").Range("B" & v).Value = "Reverse Letters"
        Sheets("List").Range("C" & v).Value = Left(Phrase, w - 1) & Letter2 & Letter1 & Right(Phrase, Len(Phrase) - (w + 1))
        v = v + 1
    End If
Next w
Sheets("Main").Range("A10").ClearContents

End Function

Function SkipSpaces(Phrase, z, v)

Dim w, x, y As Long
Dim Letter As String

Sheets("Main").Range("A10").Value = "Skip Spaces Generating..."
For w = 1 To z
    Letter = Mid(Phrase, w, 1)
    If Letter = " " Then
        Sheets("List").Range("A" & v).Value = Phrase
        Sheets("List").Range("B" & v).Value = "Skip Spaces"
        Sheets("List").Range("C" & v).Value = Left(Phrase, w - 1) & Right(Phrase, z - w)
        v = v + 1
    End If
Next w
Sheets("Main").Range("A10").ClearContents

End Function

Function MissedKey(Phrase, z, v, Alphabet, MissedKeys)

Dim w, x, y As Long
Dim SelectedMissedKey As String

Sheets("Main").Range("A10").Value = "Missed Key Generating..."
For w = 1 To z
    If InStrRev(Alphabet, LCase(Mid(Phrase, w, 1))) > 0 Then
        SelectedMissedKey = MissedKeys(InStrRev(Alphabet, LCase(Mid(Phrase, w, 1))))
        y = Len(SelectedMissedKey)
        For x = 1 To y
            Sheets("List").Range("A" & v).Value = Phrase
            Sheets("List").Range("B" & v).Value = "Missed Key"
            Sheets("List").Range("C" & v).Value = Left(Phrase, w - 1) & Mid(SelectedMissedKey, x, 1) & Right(Phrase, Len(Phrase) - w)
            v = v + 1
        Next x
    End If
Next w
Sheets("Main").Range("A10").ClearContents

End Function

Function InsertedKey(Phrase, z, v, Alphabet, MissedKeys)

Dim w, x, y As Long
Dim SelectedMissedKey As String

Sheets("Main").Range("A10").Value = "Inserted Key Generating..."
For w = 1 To z
    If InStrRev(Alphabet, LCase(Mid(Phrase, w, 1))) > 0 Then
        SelectedMissedKey = MissedKeys(InStrRev(Alphabet, LCase(Mid(Phrase, w, 1))))
        y = Len(SelectedMissedKey)
        For x = 1 To y
            Sheets("List").Range("A" & v).Value = Phrase
            Sheets("List").Range("B" & v).Value = "Inserted Key"
            Sheets("List").Range("C" & v).Value = Left(Phrase, w - 1) & Mid(SelectedMissedKey, x, 1) & Right(Phrase, (Len(Phrase) - w) + 1)
            v = v + 1
            Sheets("List").Range("A" & v).Value = Phrase
            Sheets("List").Range("B" & v).Value = "Inserted Key"
            Sheets("List").Range("C" & v).Value = Left(Phrase, w) & Mid(SelectedMissedKey, x, 1) & Right(Phrase, (Len(Phrase) - w) + 2)
            v = v + 1
        Next x
    End If
Next w
Sheets("Main").Range("A10").ClearContents

End Function
