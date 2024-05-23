Sub CountWords()

Dim v, w, x, y, z, Col, NumWords, MinCount, MaxTimer As Long
Dim Alphabet, Symbols, Term, NewTerm, Letter, Term2, Current, CountFileLine, Line, Folder As String
Dim Terms() As String
Dim fso As New FileSystemObject
Dim CountFile As Integer
Dim Stream1, Stream2, Stream3, Stream4, Stream5, Stream6, Stream7 As TextStream

MinCount = 2

'Create folder
x = 1
Do Until Len(Dir(ThisWorkbook.Path & "\vba" & x, vbDirectory)) = 0
    x = x + 1
Loop
Folder = ThisWorkbook.Path & "\vba" & x
MkDir Folder

Set Stream1 = fso.CreateTextFile(Folder & "\count1.txt", True)
Set Stream2 = fso.CreateTextFile(Folder & "\count2.txt", True)
Set Stream3 = fso.CreateTextFile(Folder & "\count3.txt", True)
Set Stream4 = fso.CreateTextFile(Folder & "\count4.txt", True)
Set Stream5 = fso.CreateTextFile(Folder & "\count5.txt", True)
Set Stream6 = fso.CreateTextFile(Folder & "\count6.txt", True)

Sheets("Counts").Range("A2:XFD" & Rows.Count).ClearContents
Sheets("Main").Select
Sheets("Main").Range("D1").Value = "Step 1 - Breakdown Keywords"
Application.ScreenUpdating = False

Alphabet = "abcdefghijklmnopqrstuvwxyz0123456789@ "
Symbols = "-'"
y = Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row

For x = 2 To y
    'Breakdown Keyword
    Term = Trim(LCase(Sheets("Main").Range("A" & x).Text))
    NewTerm = ""
    z = Len(Term)
    For w = 1 To z
        Letter = Mid(Term, w, 1)
        If InStrRev(Alphabet, Letter) > 0 Or InStrRev(Symbols, Letter) > 0 Then
            NewTerm = NewTerm & Letter
        End If
    Next w
    NewTerm = Trim(NewTerm)
    NumWords = Len(NewTerm) - Len(Application.WorksheetFunction.Substitute(NewTerm, " ", ""))
    If Len(NewTerm) > 0 And NewTerm <> " " Then
        ReDim Terms(NumWords)
        Terms() = Split(NewTerm, " ")

        'Clean Breakdowns
        For w = 0 To NumWords
            Term = Terms(w)
            z = 1
            Do Until InStrRev(Symbols, Mid(Term, z, 1)) = 0 Or z > Len(Term)
                z = z + 1
            Loop

            If z > Len(Term) Then
                Terms(w) = ""
            Else
                Term = Right(Term, Len(Term) - z + 1)
                z = Len(Term)
                Do Until InStrRev(Symbols, Mid(Term, z, 1)) = 0
                    z = z - 1
                Loop

                If z <= Len(Term) Then
                    Terms(w) = Left(Term, z)
                End If
            End If
        Next w
        
        'Bucket Breakdowns
        For w = 0 To NumWords
            Term2 = Terms(w)
            Stream1.WriteLine Term2
            If NumWords - w >= 1 Then
                Term2 = Term2 & " " & Terms(w + 1)
                Stream2.WriteLine Term2
            End If
            If NumWords - w >= 2 Then
                Term2 = Term2 & " " & Terms(w + 2)
                Stream3.WriteLine Term2
            End If
            If NumWords - w >= 3 Then
                Term2 = Term2 & " " & Terms(w + 3)
                Stream4.WriteLine Term2
            End If
            If NumWords - w >= 4 Then
                Term2 = Term2 & " " & Terms(w + 4)
                Stream5.WriteLine Term2
            End If
            If NumWords - w >= 5 Then
                Term2 = Term2 & " " & Terms(w + 5)
                Stream6.WriteLine Term2
            End If
        Next w
    End If

    If x Mod 10000 = 0 Then
        Application.ScreenUpdating = True
        Sheets("Main").Range("D1").Value = "Step 1 - Breakdown Keywords - " & x & " of " & y
        Application.ScreenUpdating = False
    End If
Next x
Stream1.WriteLine "zzzzzzzz"
Stream2.WriteLine "zzzzzzzz"
Stream3.WriteLine "zzzzzzzz"
Stream4.WriteLine "zzzzzzzz"
Stream5.WriteLine "zzzzzzzz"
Stream6.WriteLine "zzzzzzzz"
Stream1.Close
Stream2.Close
Stream3.Close
Stream4.Close
Stream5.Close
Stream6.Close

'Sorting Breakdowns
Sheets("Main").Select
Application.ScreenUpdating = True
ChDir Folder
MaxTimer = 300
For x = 1 To 6
    Shell "Sort count" & x & ".txt /O count" & x & "_sorted.txt", vbHide
    z = 1
    Do Until Dir("count" & x & "_sorted.txt") <> ""
        Application.Wait (Now + TimeValue("0:00:01"))
        Sheets("Main").Range("D1").Value = "Step 2 - Sorting Breakdowns - File " & x & " - Waiting " & z & " seconds..."
        z = z + 1
            
        If z > MaxTimer Then
            MsgBox "This is taking too long. Press 'OK' to Exit.", vbCritical, "Error"
            Exit Sub
        End If
    Loop
Next x

'Counting Breakdowns
Application.ScreenUpdating = False
Set Stream1 = fso.CreateTextFile(Folder & "\count1_counted.txt", True)
Set Stream2 = fso.CreateTextFile(Folder & "\count2_counted.txt", True)
Set Stream3 = fso.CreateTextFile(Folder & "\count3_counted.txt", True)
Set Stream4 = fso.CreateTextFile(Folder & "\count4_counted.txt", True)
Set Stream5 = fso.CreateTextFile(Folder & "\count5_counted.txt", True)
Set Stream6 = fso.CreateTextFile(Folder & "\count6_counted.txt", True)
'Set Stream7 = fso.CreateTextFile(ThisWorkbook.Path & "\countall.txt", True)

For x = 1 To 6
    Sheets("Main").Range("D1").Value = "Step 3 - Counting Breakdowns - " & x & " of 6"
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    
    z = 1
    Current = ""
    CountFile = FreeFile()
    Open Folder & "\count" & x & "_sorted.txt" For Input As #CountFile
    While Not EOF(CountFile)
        Line Input #CountFile, CountFileLine
    
        If CountFileLine <> Current Then
            If Current <> "" And Len(Application.WorksheetFunction.Substitute(Current, " ", "")) > 0 And z >= MinCount Then
                If x = 1 Then
                    Stream1.WriteLine Current & vbTab & z
                ElseIf x = 2 Then
                    Stream2.WriteLine Current & vbTab & z
                ElseIf x = 3 Then
                    Stream3.WriteLine Current & vbTab & z
                ElseIf x = 4 Then
                    Stream4.WriteLine Current & vbTab & z
                ElseIf x = 5 Then
                    Stream5.WriteLine Current & vbTab & z
                ElseIf x = 6 Then
                    Stream6.WriteLine Current & vbTab & z
                End If
                'Stream7.WriteLine Current & vbTab & z
            End If
            Current = CountFileLine
            z = 1
        Else
            z = z + 1
        End If
    Wend
    Close CountFile
Next x
Stream1.Close
Stream2.Close
Stream3.Close
Stream4.Close
Stream5.Close
Stream6.Close
'Stream7.Close

'Importing Counts
For x = 1 To 6
    Sheets("Main").Select
    Sheets("Main").Range("D1").Value = "Step 4 - Importing Counts - " & x & " of 6"
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Col = x * 2 + (x - 2)
    
    If FileLen(Folder & "\count" & x & "_counted.txt") > 0 Then
        Sheets("Counts").Select
        With Sheets("Counts").QueryTables.Add(Connection:="TEXT;" & Folder & "\count" & x & "_counted.txt", Destination:=Cells(2, Col))
            .Name = "Count" & x
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 65001
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        ActiveWorkbook.Connections(1).Delete
    
        Sheets("Counts").Sort.SortFields.Clear
        Sheets("Counts").Sort.SortFields.Add Key:=Columns(Col + 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Counts").Sort
            .SetRange Range(Columns(Col), Columns(Col + 1))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
Next x

Sheets("Main").Range("D1").ClearContents
Sheets("Counts").Select
Sheets("Counts").Range("A:XFD").EntireColumn.AutoFit
Application.ScreenUpdating = True
ChDir ThisWorkbook.Path
Kill Folder & "\*.*"
RmDir Folder & "\"

End Sub

