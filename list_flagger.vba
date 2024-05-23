Sub FlagList()

Dim w, x, y, z As Long
Dim Flag, FoundPhrase As String
Dim AllFound() As String
Dim Found, ExactMatch, WriteFlag As Boolean

ThisWorkbook.Application.ScreenUpdating = False
Sheets("Main").Range("G2:G" & Rows.Count).ClearContents
If Sheets("Main").Range("B2").Value = 1 Then
    ExactMatch = False
Else
    ExactMatch = True
End If
y = Sheets("Main").Range("D" & Rows.Count).End(xlUp).Row

For x = 2 To y
    Flag = Trim(LCase(Sheets("Main").Range("D" & x).Value))
    Erase AllFound
    Found = FindAll(Flag, Sheets("Main"), "F:F", AllFound(), 2)
    If Found = True Then
        For w = 1 To UBound(AllFound)
            WriteFlag = False
            If ExactMatch = False Then
                WriteFlag = True
            Else
                FoundPhrase = Trim(LCase(Sheets("Main").Range(AllFound(w)).Value))
                If FoundPhrase = Flag Or _
                    Left(FoundPhrase, Len(Flag)) & " " = Flag & " " Or _
                    " " & Right(FoundPhrase, Len(Flag)) = " " & Flag Or _
                    InStr(1, FoundPhrase, " " & Flag & " ") > 0 Then
                        WriteFlag = True
                Else
                    WriteFlag = False
                End If
            End If
            
            If WriteFlag = True Then
                If IsEmpty(Sheets("Main").Range(AllFound(w)).Offset(, 1)) = True Then
                    Sheets("Main").Range(AllFound(w)).Offset(, 1).Value = Flag
                Else
                    Sheets("Main").Range(AllFound(w)).Offset(, 1).Value = Sheets("Main").Range(AllFound(w)).Offset(, 1).Value & ", " & Flag
                End If
            End If
        Next w
    End If
Next x
ThisWorkbook.Application.ScreenUpdating = True
MsgBox "Done!", vbInformation, "Done!"

End Sub

Function FindAll(ByVal SearchText As String, ByRef SheetName As Worksheet, ByRef SearchRange As String, ByRef Matches() As String, SearchType As Integer) As Boolean

Dim FoundRange As Range
Dim x As Integer
Dim FirstAddress

On Error GoTo Error_Trap

x = 0
Erase Matches
Set FoundRange = SheetName.Range(SearchRange).Find(What:=SearchText, LookIn:=xlValues, LookAt:=SearchType)
If Not FoundRange Is Nothing Then
    FirstAddress = FoundRange.Address
    Do Until FoundRange Is Nothing
        x = x + 1
        ReDim Preserve Matches(x)
        Matches(x) = FoundRange.Address
        Set FoundRange = SheetName.Range(SearchRange).FindNext(FoundRange)
        If FoundRange.Address = FirstAddress Then
            Exit Do
        End If
    Loop
    FindAll = True
Else
    FindAll = False
End If

Error_Trap:
If Err <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbInformation, "Find All"
    Err.Clear
    FindAll = False
    Exit Function
End If

End Function
