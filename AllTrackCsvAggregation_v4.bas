Attribute VB_Name = "AllTrackCsvAggregation"
Option Explicit

Private Const CSV_ROOT_FOLDER As String = "csv"
Private Const SORTED_ROOT_FOLDER As String = "Sorted_Data"

Private Const CONTROL_SHEET As String = "Control"
Private Const HEADER_ROW As Long = 3
Private Const KEY_SEPARATOR As String = "|||"
Private Const EXAM_SEPARATOR As String = "###"
Private Const SHEET_SUFFIX_USER As String = "_user"
Private Const SHEET_SUFFIX_PASSED As String = "_passed"

Private gCurrentStep As String
Private gMatchedExamRows As Long
Private gUnmatchedCourseRows As Long
Private gCourseTitleDebug As Object

' ============================================================
' Setup
' ============================================================

Public Sub SetupCsvAggregation()

    Dim ws As Worksheet
    Dim tracks As Variant
    Dim i As Long

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Set ws = GetOrCreateSheet(CONTROL_SHEET)
    ws.Cells.Clear

    ws.Range("C1").Value = "Start Folder"
    ws.Range("E1").Value = "End Folder"

    ws.Range("D1").NumberFormat = "@"
    ws.Range("F1").NumberFormat = "@"

    ws.Range("C1:F1").Font.Bold = True

    CreateRunButton ws

    tracks = GetTrackNames()

    For i = LBound(tracks) To UBound(tracks)
        Set ws = GetOrCreateSheet(CStr(tracks(i)) & SHEET_SUFFIX_USER)
        ClearOutputArea ws
        WriteUserSummary ws, 0, 0
        WriteUserHeaders ws, 0
        FormatSheetAtRow ws, HEADER_ROW

        Set ws = GetOrCreateSheet(CStr(tracks(i)) & SHEET_SUFFIX_PASSED)
        ClearOutputArea ws
        WritePassedSummaryEmpty ws, GetMaxExamCountForTrack(CStr(tracks(i)))
        WritePassedHeaders ws, GetMaxExamCountForTrack(CStr(tracks(i))), _
            GetPassedHeaderRow(GetMaxExamCountForTrack(CStr(tracks(i))))
        FormatSheetAtRow ws, GetPassedHeaderRow(GetMaxExamCountForTrack(CStr(tracks(i))))
    Next i

    Application.ScreenUpdating = True

    MsgBox _
        "Setup completed." & vbCrLf & vbCrLf & _
        "Enter the start folder in D1 and end folder in F1 on the Control sheet." & vbCrLf & _
        "Examples: 20260113 or 2026_1_13" & vbCrLf & _
        "Leave either cell blank to include the full available range.", _
        vbInformation

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Setup error." & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
End Sub

' ============================================================
' Main
' ============================================================

Public Sub RunCsvAggregation()

    Dim basePath As String
    Dim csvRootPath As String
    Dim tracks As Variant
    Dim courseMap As Object

    Dim userDataByTrack As Object
    Dim userOrderByTrack As Object
    Dim userCoursesByTrack As Object

    Dim passedDataByTrack As Object
    Dim passedOrderByTrack As Object
    Dim passedExamDataByTrack As Object
    Dim passedExamOrderByTrack As Object

    Dim dateFolders As Variant
    Dim csvFiles As Variant

    Dim startKey As String
    Dim endKey As String
    Dim folderKey As String

    Dim i As Long
    Dim j As Long
    Dim t As Long

    Dim dateFolderName As String
    Dim dateFolderPath As String
    Dim trackName As String
    Dim trackFolderPath As String
    Dim csvFilePath As String

    Dim processedFileCount As Long
    Dim skippedFileCount As Long
    Dim selectedFolderCount As Long
    Dim foundCsvFileCount As Long

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gCurrentStep = "Resolving workbook folder"

    basePath = ResolveWorkbookLocalPath(ThisWorkbook.Path)

    If Len(basePath) = 0 Then
        Err.Raise vbObjectError + 901, , _
            "Could not resolve the local folder for this workbook."
    End If

    gCurrentStep = "Resolving CSV root folder"

    csvRootPath = ResolveCsvRoot(basePath)

    If Len(csvRootPath) = 0 Then
        Err.Raise vbObjectError + 902, , _
            "No CSV root folder was found. Expected csv, Sorted_Data, or date folders beside the macro workbook."
    End If

    EnsureControlSheet

    gCurrentStep = "Reading aggregation period"

    startKey = ParsePeriodCell(ThisWorkbook.Worksheets(CONTROL_SHEET).Range("D1").Value)
    endKey = ParsePeriodCell(ThisWorkbook.Worksheets(CONTROL_SHEET).Range("F1").Value)

    If Len(startKey) > 0 And Len(endKey) > 0 Then
        If CLng(startKey) > CLng(endKey) Then
            Err.Raise vbObjectError + 903, , _
                "Start Folder is later than End Folder."
        End If
    End If

    tracks = GetTrackNames()
    Set courseMap = BuildCourseToTrackMap()

    gMatchedExamRows = 0
    gUnmatchedCourseRows = 0
    Set gCourseTitleDebug = NewDictionary()

    Set userDataByTrack = NewDictionary()
    Set userOrderByTrack = NewDictionary()
    Set userCoursesByTrack = NewDictionary()

    Set passedDataByTrack = NewDictionary()
    Set passedOrderByTrack = NewDictionary()
    Set passedExamDataByTrack = NewDictionary()
    Set passedExamOrderByTrack = NewDictionary()

    InitializeTrackContainers _
        tracks, _
        userDataByTrack, _
        userOrderByTrack, _
        userCoursesByTrack, _
        passedDataByTrack, _
        passedOrderByTrack, _
        passedExamDataByTrack, _
        passedExamOrderByTrack

    gCurrentStep = "Reading date folders"

    dateFolders = GetSortedDateFolders(csvRootPath)

    If IsEmpty(dateFolders) Then
        Err.Raise vbObjectError + 904, , _
            "No valid date folders were found under: " & csvRootPath
    End If

    For i = LBound(dateFolders) To UBound(dateFolders)

        dateFolderName = CStr(dateFolders(i))
        folderKey = NormalizeDateFolderName(dateFolderName)

        If IsFolderInPeriod(folderKey, startKey, endKey) Then

            selectedFolderCount = selectedFolderCount + 1
            dateFolderPath = CombinePath(csvRootPath, dateFolderName)

            For t = LBound(tracks) To UBound(tracks)

                trackName = CStr(tracks(t))
                trackFolderPath = ResolveTrackFolder(dateFolderPath, trackName)

                If Len(trackFolderPath) > 0 Then

                    csvFiles = GetSortedCsvFiles(trackFolderPath)

                    If Not IsEmpty(csvFiles) Then

                        For j = LBound(csvFiles) To UBound(csvFiles)

                            csvFilePath = CombinePath(trackFolderPath, CStr(csvFiles(j)))
                            foundCsvFileCount = foundCsvFileCount + 1
                            gCurrentStep = "Processing CSV: " & csvFilePath

                            If ProcessCsvFile( _
                                csvFilePath, _
                                trackName, _
                                courseMap, _
                                userDataByTrack, _
                                userOrderByTrack, _
                                userCoursesByTrack, _
                                passedDataByTrack, _
                                passedOrderByTrack, _
                                passedExamDataByTrack, _
                                passedExamOrderByTrack) Then

                                processedFileCount = processedFileCount + 1
                            Else
                                skippedFileCount = skippedFileCount + 1
                            End If

                        Next j

                    End If

                End If

            Next t

        End If

    Next i

    If selectedFolderCount = 0 Then
        Err.Raise vbObjectError + 905, , _
            "No date folders matched the period specified in D1 and F1."
    End If

    gCurrentStep = "Writing aggregation sheets"

    OutputAllTracks _
        tracks, _
        userDataByTrack, _
        userOrderByTrack, _
        userCoursesByTrack, _
        passedDataByTrack, _
        passedOrderByTrack, _
        passedExamDataByTrack, _
        passedExamOrderByTrack

    gCurrentStep = "Writing CSV aggregation debug sheet"

    WriteCsvAggregationDebug courseMap

    gCurrentStep = "Exporting track workbooks"

    ExportTrackWorkbooks tracks, basePath

    gCurrentStep = "Completed"

    MsgBox _
        "CSV aggregation completed." & vbCrLf & vbCrLf & _
        "CSV root: " & csvRootPath & vbCrLf & _
        "Selected date folders: " & selectedFolderCount & vbCrLf & _
        "Found CSV files: " & foundCsvFileCount & vbCrLf & _
        "Processed CSV files: " & processedFileCount & vbCrLf & _
        "Skipped CSV files: " & skippedFileCount & vbCrLf & _
        "Matched Exam Rows: " & gMatchedExamRows & vbCrLf & _
        "Unmatched Course Rows: " & gUnmatchedCourseRows & vbCrLf & vbCrLf & _
        "See CsvAggregation_Debug for actual Course title values.", _
        vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox _
        "CSV aggregation error." & vbCrLf & vbCrLf & _
        "Step: " & gCurrentStep & vbCrLf & _
        "Error number: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description, _
        vbCritical
    Resume SafeExit
End Sub

' ============================================================
' CSV processing
' ============================================================

Private Function ProcessCsvFile( _
    ByVal filePath As String, _
    ByVal expectedTrack As String, _
    ByVal courseMap As Object, _
    ByVal userDataByTrack As Object, _
    ByVal userOrderByTrack As Object, _
    ByVal userCoursesByTrack As Object, _
    ByVal passedDataByTrack As Object, _
    ByVal passedOrderByTrack As Object, _
    ByVal passedExamDataByTrack As Object, _
    ByVal passedExamOrderByTrack As Object) As Boolean

    Dim fileNo As Integer
    Dim lineText As String
    Dim fields As Variant
    Dim headers As Object

    Dim username As String
    Dim email As String
    Dim fullName As String
    Dim courseTitle As String
    Dim courseStatus As String
    Dim enrollmentDate As Variant
    Dim enrollmentEndDate As Variant
    Dim finalScore As Variant

    Dim trackName As String
    Dim userKey As String
    Dim examKey As String

    Dim users As Object
    Dim userOrder As Collection
    Dim userCourses As Object

    Dim passedUsers As Object
    Dim passedOrder As Collection
    Dim passedExamData As Object
    Dim passedExamOrder As Object

    Dim userInfo As Variant
    Dim passedInfo As Variant
    Dim courseOrder As Collection
    Dim passedOrderCollection As Collection

    On Error GoTo FileError

    fileNo = FreeFile
    Open filePath For Input As #fileNo

    If EOF(fileNo) Then GoTo FileError

    Line Input #fileNo, lineText
    fields = ParseCsvLine(lineText)
    Set headers = BuildHeaderMap(fields)

    If Not headers.Exists(LCase$("Course title")) Then GoTo FileError

    Do While Not EOF(fileNo)

        Line Input #fileNo, lineText
        fields = ParseCsvLine(lineText)

        courseTitle = CsvField(fields, headers, "Course title")

        If Len(courseTitle) = 0 Then GoTo NextLine

        RecordCourseTitleDebug courseTitle, courseMap.Exists(LCase$(courseTitle))

        If Not courseMap.Exists(LCase$(courseTitle)) Then
            gUnmatchedCourseRows = gUnmatchedCourseRows + 1
            GoTo NextLine
        End If

        gMatchedExamRows = gMatchedExamRows + 1
        trackName = CStr(courseMap(LCase$(courseTitle)))

        If StrComp(trackName, expectedTrack, vbTextCompare) <> 0 Then
            GoTo NextLine
        End If

        username = CsvField(fields, headers, "Username")
        email = CsvField(fields, headers, "Email")
        fullName = CsvField(fields, headers, "Full Name")

        enrollmentDate = CsvField(fields, headers, "Enrollment Date")
        enrollmentEndDate = CsvField(fields, headers, "Enrollment End Date")
        courseStatus = CsvField(fields, headers, "Course Enrollment Status")
        finalScore = CsvField(fields, headers, "Final Score")

        userKey = MakeUserKey(username, email, fullName)
        If Len(userKey) = 0 Then GoTo NextLine

        Set users = userDataByTrack(trackName)
        Set userOrder = userOrderByTrack(trackName)
        Set userCourses = userCoursesByTrack(trackName)

        If Not users.Exists(userKey) Then

            userInfo = Array( _
                username, _
                email, _
                fullName, _
                enrollmentDate, _
                enrollmentEndDate)

            users.Add userKey, userInfo
            userOrder.Add userKey

            Set courseOrder = New Collection
            userCourses.Add userKey, courseOrder

        Else

            userInfo = users(userKey)

            If Len(username) > 0 Then userInfo(0) = username
            If Len(email) > 0 Then userInfo(1) = email
            If Len(fullName) > 0 Then userInfo(2) = fullName

            users(userKey) = userInfo

        End If

        Set courseOrder = userCourses(userKey)

        examKey = userKey & EXAM_SEPARATOR & LCase$(courseTitle)

        If Not userCourses.Exists(examKey) Then
            userCourses.Add examKey, Array(courseTitle)
            courseOrder.Add courseTitle
        Else
            userCourses(examKey) = Array(courseTitle)
        End If

        If StrComp(courseStatus, "Completed", vbTextCompare) = 0 Then

            Set passedUsers = passedDataByTrack(trackName)
            Set passedOrder = passedOrderByTrack(trackName)
            Set passedExamData = passedExamDataByTrack(trackName)
            Set passedExamOrder = passedExamOrderByTrack(trackName)

            If Not passedUsers.Exists(userKey) Then

                passedInfo = Array( _
                    username, _
                    email, _
                    fullName, _
                    enrollmentDate, _
                    enrollmentEndDate)

                passedUsers.Add userKey, passedInfo
                passedOrder.Add userKey

                Set passedOrderCollection = New Collection
                passedExamOrder.Add userKey, passedOrderCollection

            Else

                passedInfo = passedUsers(userKey)

                If Len(username) > 0 Then passedInfo(0) = username
                If Len(email) > 0 Then passedInfo(1) = email
                If Len(fullName) > 0 Then passedInfo(2) = fullName

                passedInfo(3) = enrollmentDate
                passedInfo(4) = enrollmentEndDate

                passedUsers(userKey) = passedInfo

            End If

            examKey = userKey & EXAM_SEPARATOR & LCase$(courseTitle)

            If Not passedExamData.Exists(examKey) Then

                passedExamData.Add examKey, _
                    Array(courseTitle, "Completed", finalScore)

                Set passedOrderCollection = passedExamOrder(userKey)
                passedOrderCollection.Add courseTitle

            Else

                passedExamData(examKey) = _
                    Array(courseTitle, "Completed", finalScore)

            End If

        End If

NextLine:
    Loop

    Close #fileNo
    ProcessCsvFile = True
    Exit Function

FileError:
    On Error Resume Next
    Close #fileNo
    On Error GoTo 0
    ProcessCsvFile = False
End Function

Private Function BuildHeaderMap(ByVal fields As Variant) As Object

    Dim d As Object
    Dim i As Long
    Dim headerText As String

    Set d = NewDictionary()

    For i = LBound(fields) To UBound(fields)

        headerText = CleanText(fields(i))
        headerText = Replace(headerText, ChrW(&HFEFF), "")

        If Len(headerText) > 0 Then
            d(LCase$(headerText)) = i
        End If

    Next i

    Set BuildHeaderMap = d
End Function

Private Function CsvField( _
    ByVal fields As Variant, _
    ByVal headers As Object, _
    ByVal headerName As String) As String

    Dim idx As Long

    If Not headers.Exists(LCase$(headerName)) Then
        CsvField = ""
        Exit Function
    End If

    idx = CLng(headers(LCase$(headerName)))

    If idx < LBound(fields) Or idx > UBound(fields) Then
        CsvField = ""
    Else
        CsvField = CleanText(fields(idx))
    End If
End Function

Private Function ParseCsvLine(ByVal lineText As String) As Variant

    Dim values As Collection
    Dim arr() As String

    Dim i As Long
    Dim ch As String
    Dim currentValue As String
    Dim inQuotes As Boolean

    Set values = New Collection

    i = 1

    Do While i <= Len(lineText)

        ch = Mid$(lineText, i, 1)

        If ch = """" Then

            If inQuotes Then

                If i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                    currentValue = currentValue & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If

            Else
                inQuotes = True
            End If

        ElseIf ch = "," And Not inQuotes Then

            values.Add currentValue
            currentValue = ""

        Else
            currentValue = currentValue & ch
        End If

        i = i + 1
    Loop

    values.Add currentValue

    ReDim arr(0 To values.Count - 1)

    For i = 1 To values.Count
        arr(i - 1) = CStr(values(i))
    Next i

    ParseCsvLine = arr
End Function


' ============================================================
' CSV aggregation diagnostics
' ============================================================

Private Sub RecordCourseTitleDebug( _
    ByVal courseTitle As String, _
    ByVal isMatched As Boolean)

    Dim key As String
    Dim info As Variant

    key = LCase$(Trim$(courseTitle))

    If Not gCourseTitleDebug.Exists(key) Then

        info = Array( _
            courseTitle, _
            IIf(isMatched, "Matched", "Unmatched"), _
            1)

        gCourseTitleDebug.Add key, info

    Else

        info = gCourseTitleDebug(key)
        info(2) = CLng(info(2)) + 1

        If isMatched Then
            info(1) = "Matched"
        End If

        gCourseTitleDebug(key) = info

    End If

End Sub


Private Sub WriteCsvAggregationDebug(ByVal courseMap As Object)

    Dim ws As Worksheet
    Dim key As Variant
    Dim info As Variant
    Dim rowNo As Long

    Set ws = GetOrCreateSheet("CsvAggregation_Debug")
    ws.Cells.Clear

    ws.Range("A1:C1").Value = Array( _
        "Course title in CSV", _
        "Exam Match", _
        "Row Count")

    ws.Rows(1).Font.Bold = True

    rowNo = 2

    For Each key In gCourseTitleDebug.Keys

        info = gCourseTitleDebug(CStr(key))

        ws.Cells(rowNo, 1).Value = info(0)
        ws.Cells(rowNo, 2).Value = info(1)
        ws.Cells(rowNo, 3).Value = info(2)

        If StrComp(CStr(info(1)), "Unmatched", vbTextCompare) = 0 Then
            ws.Cells(rowNo, 2).Interior.Color = RGB(255, 192, 0)
        End If

        rowNo = rowNo + 1

    Next key

    ws.Columns.AutoFit

    If ws.Columns(1).ColumnWidth > 80 Then
        ws.Columns(1).ColumnWidth = 80
    End If

End Sub


' ============================================================
' Output all tracks
' ============================================================

Private Sub OutputAllTracks( _
    ByVal trackNames As Variant, _
    ByVal userDataByTrack As Object, _
    ByVal userOrderByTrack As Object, _
    ByVal userCoursesByTrack As Object, _
    ByVal passedDataByTrack As Object, _
    ByVal passedOrderByTrack As Object, _
    ByVal passedExamDataByTrack As Object, _
    ByVal passedExamOrderByTrack As Object)

    Dim i As Long
    Dim trackName As String

    For i = LBound(trackNames) To UBound(trackNames)

        trackName = CStr(trackNames(i))

        OutputUserTrack _
            trackName, _
            userDataByTrack(trackName), _
            userOrderByTrack(trackName), _
            userCoursesByTrack(trackName)

        OutputPassedTrack _
            trackName, _
            passedDataByTrack(trackName), _
            passedOrderByTrack(trackName), _
            passedExamDataByTrack(trackName), _
            passedExamOrderByTrack(trackName), _
            GetMaxExamCountForTrack(trackName)

    Next i
End Sub

Private Sub OutputUserTrack( _
    ByVal trackName As String, _
    ByVal users As Object, _
    ByVal userOrder As Collection, _
    ByVal userCourses As Object)

    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim c As Long

    Dim key As String
    Dim examKey As String
    Dim courseTitle As String

    Dim info As Variant
    Dim courseInfo As Variant
    Dim courseOrder As Collection

    Dim maxCourseCount As Long
    Dim totalColumns As Long
    Dim outputData() As Variant
    Dim nonStudentCount As Long

    Set ws = GetOrCreateSheet(trackName & SHEET_SUFFIX_USER)
    ClearOutputArea ws

    For i = 1 To userOrder.Count
        key = CStr(userOrder(i))
        Set courseOrder = userCourses(key)

        If courseOrder.Count > maxCourseCount Then
            maxCourseCount = courseOrder.Count
        End If
    Next i

    nonStudentCount = CountNonStudents(users, userOrder)

    WriteUserSummary ws, users.Count, nonStudentCount
    WriteUserHeaders ws, maxCourseCount

    If users.Count = 0 Then
        FormatSheetAtRow ws, HEADER_ROW
        Exit Sub
    End If

    totalColumns = 6 + maxCourseCount
    ReDim outputData(1 To users.Count, 1 To totalColumns)

    For i = 1 To userOrder.Count

        key = CStr(userOrder(i))
        info = users(key)

        outputData(i, 1) = i
        outputData(i, 2) = info(0)
        outputData(i, 3) = info(1)
        outputData(i, 4) = info(2)
        outputData(i, 5) = info(3)
        outputData(i, 6) = info(4)

        Set courseOrder = userCourses(key)
        c = 7

        For j = 1 To courseOrder.Count

            courseTitle = CStr(courseOrder(j))
            examKey = key & EXAM_SEPARATOR & LCase$(courseTitle)
            courseInfo = userCourses(examKey)

            outputData(i, c) = courseInfo(0)
            c = c + 1

        Next j

    Next i

    ws.Columns(2).NumberFormat = "@"

    ws.Cells(HEADER_ROW + 1, 1) _
        .Resize(users.Count, totalColumns).Value = outputData

    HighlightNonStudentEmails ws, users.Count
    FormatSheetAtRow ws, HEADER_ROW
End Sub

Private Sub OutputPassedTrack( _
    ByVal trackName As String, _
    ByVal passedUsers As Object, _
    ByVal passedOrder As Collection, _
    ByVal passedExamData As Object, _
    ByVal passedExamOrder As Object, _
    ByVal maxExamCount As Long)

    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim c As Long

    Dim key As String
    Dim examKey As String
    Dim courseTitle As String

    Dim info As Variant
    Dim examInfo As Variant
    Dim examOrder As Collection

    Dim outputData() As Variant
    Dim totalColumns As Long
    Dim passedCountBuckets() As Long
    Dim examCountForUser As Long
    Dim passedHeaderRow As Long

    Set ws = GetOrCreateSheet(trackName & SHEET_SUFFIX_PASSED)
    ClearOutputArea ws

    ReDim passedCountBuckets(1 To maxExamCount)

    For i = 1 To passedOrder.Count

        key = CStr(passedOrder(i))
        Set examOrder = passedExamOrder(key)

        examCountForUser = examOrder.Count

        If examCountForUser >= 1 And examCountForUser <= maxExamCount Then
            passedCountBuckets(examCountForUser) = _
                passedCountBuckets(examCountForUser) + 1
        End If

    Next i

    passedHeaderRow = GetPassedHeaderRow(maxExamCount)

    WritePassedSummary ws, passedCountBuckets, passedUsers.Count
    WritePassedHeaders ws, maxExamCount, passedHeaderRow

    If passedUsers.Count = 0 Then
        FormatSheetAtRow ws, passedHeaderRow
        Exit Sub
    End If

    totalColumns = 6 + (maxExamCount * 3)
    ReDim outputData(1 To passedUsers.Count, 1 To totalColumns)

    For i = 1 To passedOrder.Count

        key = CStr(passedOrder(i))
        info = passedUsers(key)

        outputData(i, 1) = i
        outputData(i, 2) = info(0)
        outputData(i, 3) = info(1)
        outputData(i, 4) = info(2)
        outputData(i, 5) = info(3)
        outputData(i, 6) = info(4)

        Set examOrder = passedExamOrder(key)
        c = 7

        For j = 1 To examOrder.Count

            courseTitle = CStr(examOrder(j))
            examKey = key & EXAM_SEPARATOR & LCase$(courseTitle)
            examInfo = passedExamData(examKey)

            outputData(i, c) = examInfo(0)
            outputData(i, c + 1) = examInfo(1)
            outputData(i, c + 2) = examInfo(2)

            c = c + 3

        Next j

    Next i

    ws.Columns(2).NumberFormat = "@"

    ws.Cells(passedHeaderRow + 1, 1) _
        .Resize(passedUsers.Count, totalColumns).Value = outputData

    FormatSheetAtRow ws, passedHeaderRow
End Sub

' ============================================================
' Summaries and headers
' ============================================================

Private Sub WriteUserSummary( _
    ByVal ws As Worksheet, _
    ByVal totalUsers As Long, _
    ByVal nonStudentCount As Long)

    ws.Range("A1").Value = "Total Users"
    ws.Range("B1").Value = totalUsers

    ws.Range("A2").Value = "Non-student Users"
    ws.Range("B2").Value = nonStudentCount
End Sub

Private Sub WritePassedSummary( _
    ByVal ws As Worksheet, _
    ByRef buckets() As Long, _
    ByVal totalPassedUsers As Long)

    Dim i As Long
    Dim rowNo As Long

    rowNo = 1

    For i = LBound(buckets) To UBound(buckets)
        ws.Cells(rowNo, 1).Value = CStr(i) & " exam passed"
        ws.Cells(rowNo, 2).Value = buckets(i)
        rowNo = rowNo + 1
    Next i

    ws.Cells(rowNo, 1).Value = "Total passed users"
    ws.Cells(rowNo, 2).Value = totalPassedUsers
End Sub

Private Sub WritePassedSummaryEmpty( _
    ByVal ws As Worksheet, _
    ByVal maxExamCount As Long)

    Dim buckets() As Long
    ReDim buckets(1 To maxExamCount)
    WritePassedSummary ws, buckets, 0
End Sub

Private Sub WriteUserHeaders( _
    ByVal ws As Worksheet, _
    ByVal courseCount As Long)

    Dim i As Long
    Dim c As Long

    ws.Cells(HEADER_ROW, 1).Value = "No."
    ws.Cells(HEADER_ROW, 2).Value = "Username"
    ws.Cells(HEADER_ROW, 3).Value = "Email"
    ws.Cells(HEADER_ROW, 4).Value = "Full Name"
    ws.Cells(HEADER_ROW, 5).Value = "Enrollment Date"
    ws.Cells(HEADER_ROW, 6).Value = "Enrollment End Date"

    c = 7

    For i = 1 To courseCount
        ws.Cells(HEADER_ROW, c).Value = "Course title" & i
        c = c + 1
    Next i
End Sub

Private Sub WritePassedHeaders( _
    ByVal ws As Worksheet, _
    ByVal examCount As Long, _
    Optional ByVal headerRow As Long = HEADER_ROW)

    Dim i As Long
    Dim c As Long

    ws.Cells(headerRow, 1).Value = "No."
    ws.Cells(headerRow, 2).Value = "Username"
    ws.Cells(headerRow, 3).Value = "Email"
    ws.Cells(headerRow, 4).Value = "Full Name"
    ws.Cells(headerRow, 5).Value = "Enrollment Date"
    ws.Cells(headerRow, 6).Value = "Enrollment End Date"

    c = 7

    For i = 1 To examCount

        ws.Cells(headerRow, c).Value = "Course title" & i
        ws.Cells(headerRow, c + 1).Value = _
            "Course Enrollment Status" & i
        ws.Cells(headerRow, c + 2).Value = "Final Score" & i

        c = c + 3

    Next i
End Sub

Private Function GetPassedHeaderRow(ByVal maxExamCount As Long) As Long
    GetPassedHeaderRow = maxExamCount + 3
End Function

' ============================================================
' Student highlighting
' ============================================================

Private Function IsStudentEmail(ByVal email As String) As Boolean
    IsStudentEmail = _
        (LCase$(Trim$(email)) Like "########@lstc.adip.jp")
End Function

Private Function CountNonStudents( _
    ByVal users As Object, _
    ByVal userOrder As Collection) As Long

    Dim i As Long
    Dim key As String
    Dim info As Variant
    Dim count As Long

    For i = 1 To userOrder.Count

        key = CStr(userOrder(i))
        info = users(key)

        If Not IsStudentEmail(CStr(info(1))) Then
            count = count + 1
        End If

    Next i

    CountNonStudents = count
End Function

Private Sub HighlightNonStudentEmails( _
    ByVal ws As Worksheet, _
    ByVal userCount As Long)

    Dim r As Long
    Dim emailText As String

    If userCount <= 0 Then Exit Sub

    For r = HEADER_ROW + 1 To HEADER_ROW + userCount

        emailText = CleanText(ws.Cells(r, 3).Value)

        If Not IsStudentEmail(emailText) Then
            ws.Cells(r, 3).Interior.Color = RGB(255, 192, 0)
        Else
            ws.Cells(r, 3).Interior.Pattern = xlNone
        End If

    Next r
End Sub

' ============================================================
' Track configuration
' ============================================================

Private Function GetTrackNames() As Variant
    GetTrackNames = Array( _
        "WFD-PreReq", _
        "WFD-Design_for_Test", _
        "WFD-Design_Verification", _
        "WFD-Physical_Design", _
        "WFD-RTL_Synthesis", _
        "WFD-AMS")
End Function

Private Function BuildCourseToTrackMap() As Object

    Dim d As Object
    Set d = NewDictionary()

    AddCourse d, "Purple Certification: ASIC Design Flow Exam", "WFD-PreReq"
    AddCourse d, "Purple Certification: Digital Design Fundamentals Exam", "WFD-PreReq"
    AddCourse d, "Purple Certification: CMOS Fundamentals Exam", "WFD-PreReq"
    AddCourse d, "Purple Certification: Very Deep Submicron (VDSM) Fundamentals Exam", "WFD-PreReq"
    AddCourse d, "Purple Certification: VLSI Basics Exam", "WFD-PreReq"

    AddCourse d, "TestMAX ATPG Exam", "WFD-Design_for_Test"
    AddCourse d, "Fusion Compiler: DFT Synthesis Exam", "WFD-Design_for_Test"
    AddCourse d, "TestMAX Advisor Exam", "WFD-Design_for_Test"
    AddCourse d, "TestMAX DFT Exam", "WFD-Design_for_Test"

    AddCourse d, "SystemVerilog Assertions Exam", "WFD-Design_Verification"
    AddCourse d, "SystemVerilog Verification using UVM Exam", "WFD-Design_Verification"
    AddCourse d, "SystemVerilog Testbench Exam", "WFD-Design_Verification"

    AddCourse d, "Fusion Compiler: Design Implementation Exam", "WFD-Physical_Design"
    AddCourse d, "Fusion Compiler: Design Creation and Synthesis Exam", "WFD-Physical_Design"

    AddCourse d, "SystemVerilog for RTL Design Exam", "WFD-RTL_Synthesis"
    AddCourse d, "Design Compiler NXT: RTL Synthesis Exam", "WFD-RTL_Synthesis"
    AddCourse d, "Design Compiler NXT: Low Power Exam", "WFD-RTL_Synthesis"
    AddCourse d, "Fusion Compiler: UPF Fundamentals Exam", "WFD-RTL_Synthesis"

    AddCourse d, "Custom Compiler: Basic Layout Design Exam", "WFD-AMS"
    AddCourse d, "Custom Compiler: Introduction to Platform Exam", "WFD-AMS"
    AddCourse d, "Custom Compiler: Schematic Entry Exam", "WFD-AMS"
    AddCourse d, "PrimeWave Design Environment Exam", "WFD-AMS"

    Set BuildCourseToTrackMap = d
End Function

Private Sub AddCourse( _
    ByVal d As Object, _
    ByVal courseTitle As String, _
    ByVal trackName As String)

    d(LCase$(courseTitle)) = trackName
End Sub

Private Function GetMaxExamCountForTrack( _
    ByVal trackName As String) As Long

    Select Case trackName
        Case "WFD-PreReq": GetMaxExamCountForTrack = 5
        Case "WFD-Design_for_Test": GetMaxExamCountForTrack = 4
        Case "WFD-Design_Verification": GetMaxExamCountForTrack = 3
        Case "WFD-Physical_Design": GetMaxExamCountForTrack = 2
        Case "WFD-RTL_Synthesis": GetMaxExamCountForTrack = 4
        Case "WFD-AMS": GetMaxExamCountForTrack = 4
        Case Else: GetMaxExamCountForTrack = 0
    End Select
End Function

Private Sub InitializeTrackContainers( _
    ByVal trackNames As Variant, _
    ByVal userDataByTrack As Object, _
    ByVal userOrderByTrack As Object, _
    ByVal userCoursesByTrack As Object, _
    ByVal passedDataByTrack As Object, _
    ByVal passedOrderByTrack As Object, _
    ByVal passedExamDataByTrack As Object, _
    ByVal passedExamOrderByTrack As Object)

    Dim i As Long
    Dim trackName As String
    Dim d As Object
    Dim c As Collection

    For i = LBound(trackNames) To UBound(trackNames)

        trackName = CStr(trackNames(i))

        Set d = NewDictionary()
        userDataByTrack.Add trackName, d

        Set c = New Collection
        userOrderByTrack.Add trackName, c

        Set d = NewDictionary()
        userCoursesByTrack.Add trackName, d

        Set d = NewDictionary()
        passedDataByTrack.Add trackName, d

        Set c = New Collection
        passedOrderByTrack.Add trackName, c

        Set d = NewDictionary()
        passedExamDataByTrack.Add trackName, d

        Set d = NewDictionary()
        passedExamOrderByTrack.Add trackName, d

    Next i
End Sub

' ============================================================
' Export track workbooks
' ============================================================

Private Sub ExportTrackWorkbooks( _
    ByVal trackNames As Variant, _
    ByVal outputFolder As String)

    Dim i As Long

    For i = LBound(trackNames) To UBound(trackNames)
        ExportSingleTrackWorkbook CStr(trackNames(i)), outputFolder
    Next i
End Sub

Private Sub ExportSingleTrackWorkbook( _
    ByVal trackName As String, _
    ByVal outputFolder As String)

    Dim wsUser As Worksheet
    Dim wsPassed As Worksheet
    Dim wbOut As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim outputPath As String
    Dim tempPath As String

    On Error GoTo ExportError

    Set wsUser = ThisWorkbook.Worksheets(trackName & SHEET_SUFFIX_USER)
    Set wsPassed = ThisWorkbook.Worksheets(trackName & SHEET_SUFFIX_PASSED)

    wsUser.Copy
    Set wbOut = ActiveWorkbook

    wsPassed.Copy After:=wbOut.Worksheets(wbOut.Worksheets.Count)

    wbOut.Worksheets(1).Name = "user"
    wbOut.Worksheets(2).Name = "passed"

    For Each ws In wbOut.Worksheets

        On Error Resume Next

        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then
                shp.Delete
            End If
        Next shp

        On Error GoTo ExportError

    Next ws

    outputPath = CombinePath(outputFolder, trackName & ".xlsx")
    tempPath = CombinePath( _
        Environ$("TEMP"), _
        MakeSafeFileName(trackName) & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx")

    If FileExists(tempPath) Then Kill tempPath

    wbOut.SaveAs _
        Filename:=tempPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False

    wbOut.Close SaveChanges:=False

    CopyFileOverwrite tempPath, outputPath

    If FileExists(tempPath) Then Kill tempPath

    Exit Sub

ExportError:

    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    On Error GoTo 0

    Err.Raise errNumber, "ExportSingleTrackWorkbook", errDescription
End Sub

' ============================================================
' Period and date folders
' ============================================================

Private Sub EnsureControlSheet()

    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(CONTROL_SHEET)

    If CleanText(ws.Range("C1").Value) = "" Then
        ws.Range("C1").Value = "Start Folder"
    End If

    If CleanText(ws.Range("E1").Value) = "" Then
        ws.Range("E1").Value = "End Folder"
    End If

    ws.Range("D1").NumberFormat = "@"
    ws.Range("F1").NumberFormat = "@"

    CreateRunButton ws
End Sub

Private Function ParsePeriodCell(ByVal value As Variant) As String

    Dim s As String

    If IsEmpty(value) Or IsNull(value) Then
        ParsePeriodCell = ""
        Exit Function
    End If

    If IsDate(value) Then
        ParsePeriodCell = Format$(CDate(value), "yyyymmdd")
        Exit Function
    End If

    s = CleanText(value)

    If Len(s) = 0 Then
        ParsePeriodCell = ""
        Exit Function
    End If

    ParsePeriodCell = NormalizeDateFolderName(s)

    If Len(ParsePeriodCell) <> 8 Then
        Err.Raise vbObjectError + 906, , _
            "Invalid period value: " & s
    End If
End Function

Private Function IsFolderInPeriod( _
    ByVal folderKey As String, _
    ByVal startKey As String, _
    ByVal endKey As String) As Boolean

    If Len(folderKey) <> 8 Then Exit Function

    If Len(startKey) > 0 Then
        If CLng(folderKey) < CLng(startKey) Then Exit Function
    End If

    If Len(endKey) > 0 Then
        If CLng(folderKey) > CLng(endKey) Then Exit Function
    End If

    IsFolderInPeriod = True
End Function

Private Function GetSortedDateFolders(ByVal rootPath As String) As Variant

    Dim fso As Object
    Dim rootFolder As Object
    Dim subFolder As Object

    Dim names() As String
    Dim keys() As String
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim tempName As String
    Dim tempKey As String
    Dim key As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(rootPath)

    For Each subFolder In rootFolder.SubFolders

        key = NormalizeDateFolderName(CStr(subFolder.Name))

        If IsNormalizedDateKey(key) Then
            count = count + 1
            ReDim Preserve names(1 To count)
            ReDim Preserve keys(1 To count)
            names(count) = CStr(subFolder.Name)
            keys(count) = key
        End If

    Next subFolder

    If count = 0 Then
        GetSortedDateFolders = Empty
        Exit Function
    End If

    For i = 1 To count - 1
        For j = i + 1 To count

            If CLng(keys(i)) > CLng(keys(j)) Then

                tempName = names(i)
                names(i) = names(j)
                names(j) = tempName

                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey

            End If

        Next j
    Next i

    GetSortedDateFolders = names
End Function

Private Function NormalizeDateFolderName( _
    ByVal folderName As String) As String

    Dim s As String
    Dim parts() As String
    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long

    On Error GoTo Failed

    s = Trim$(folderName)
    s = Replace(s, "-", "_")
    s = Replace(s, ".", "_")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(&H3000), "")

    If Len(s) = 8 And IsNumeric(s) Then

        yyyy = CLng(Left$(s, 4))
        mm = CLng(Mid$(s, 5, 2))
        dd = CLng(Right$(s, 2))

        NormalizeDateFolderName = _
            Format$(yyyy, "0000") & _
            Format$(mm, "00") & _
            Format$(dd, "00")

        Exit Function
    End If

    If InStr(1, s, "_", vbBinaryCompare) > 0 Then

        parts = Split(s, "_")

        If UBound(parts) = 2 Then

            If IsNumeric(parts(0)) _
                And IsNumeric(parts(1)) _
                And IsNumeric(parts(2)) Then

                yyyy = CLng(parts(0))
                mm = CLng(parts(1))
                dd = CLng(parts(2))

                NormalizeDateFolderName = _
                    Format$(yyyy, "0000") & _
                    Format$(mm, "00") & _
                    Format$(dd, "00")

                Exit Function
            End If

        End If

    End If

Failed:
    NormalizeDateFolderName = ""
End Function

Private Function IsNormalizedDateKey(ByVal key As String) As Boolean

    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long
    Dim dt As Date

    On Error GoTo Failed

    If Len(key) <> 8 Then Exit Function
    If Not IsNumeric(key) Then Exit Function

    yyyy = CLng(Left$(key, 4))
    mm = CLng(Mid$(key, 5, 2))
    dd = CLng(Right$(key, 2))

    dt = DateSerial(yyyy, mm, dd)

    If Year(dt) <> yyyy Then Exit Function
    If Month(dt) <> mm Then Exit Function
    If Day(dt) <> dd Then Exit Function

    IsNormalizedDateKey = True
    Exit Function

Failed:
    IsNormalizedDateKey = False
End Function

Private Function ResolveTrackFolder( _
    ByVal dateFolderPath As String, _
    ByVal trackName As String) As String

    Dim candidate As String

    candidate = CombinePath(dateFolderPath, trackName)
    If FolderExists(candidate) Then
        ResolveTrackFolder = candidate
        Exit Function
    End If

    ' Backward-compatible aliases.
    If StrComp(trackName, "WFD-PreReq", vbTextCompare) = 0 Then

        candidate = CombinePath(dateFolderPath, "Pre-Requisite")
        If FolderExists(candidate) Then
            ResolveTrackFolder = candidate
            Exit Function
        End If

        candidate = CombinePath(dateFolderPath, "WFD-PreReq")
        If FolderExists(candidate) Then
            ResolveTrackFolder = candidate
            Exit Function
        End If

    End If

    ResolveTrackFolder = ""

End Function

Private Function GetSortedCsvFiles(ByVal folderPath As String) As Variant

    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object

    Dim arr() As String
    Dim count As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(folderPath)

    For Each fileObj In folderObj.Files

        If LCase$(fso.GetExtensionName(CStr(fileObj.Name))) = "csv" Then
            count = count + 1
            ReDim Preserve arr(1 To count)
            arr(count) = CStr(fileObj.Name)
        End If

    Next fileObj

    If count = 0 Then
        GetSortedCsvFiles = Empty
        Exit Function
    End If

    SortStringArray arr
    GetSortedCsvFiles = arr
End Function

Private Sub SortStringArray(ByRef arr As Variant)

    Dim i As Long
    Dim j As Long
    Dim tmp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)

            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                tmp = CStr(arr(i))
                arr(i) = arr(j)
                arr(j) = tmp
            End If

        Next j
    Next i
End Sub

' ============================================================
' CSV root and paths
' ============================================================

Private Function ResolveCsvRoot(ByVal basePath As String) As String

    Dim candidate As String

    candidate = CombinePath(basePath, CSV_ROOT_FOLDER)

    If FolderExists(candidate) Then
        ResolveCsvRoot = candidate
    Else
        ResolveCsvRoot = ""
    End If

End Function

Private Function ResolveWorkbookLocalPath( _
    ByVal workbookPath As String) As String

    Dim relativePath As String
    Dim oneDriveRoot As String
    Dim testPath As String

    If Len(workbookPath) = 0 Then Exit Function

    If InStr(1, workbookPath, "://", vbTextCompare) = 0 Then
        ResolveWorkbookLocalPath = workbookPath
        Exit Function
    End If

    relativePath = GetSharePointRelativePath(workbookPath)

    If Len(relativePath) = 0 Then Exit Function

    oneDriveRoot = GetBestOneDriveRoot()

    If Len(oneDriveRoot) = 0 Then Exit Function

    testPath = CombinePath(oneDriveRoot, relativePath)

    If FolderExists(testPath) Then
        ResolveWorkbookLocalPath = testPath
        Exit Function
    End If

    If LCase$(Left$(relativePath, Len("Desktop\"))) = LCase$("Desktop\") Then

        testPath = CombinePath( _
            Environ$("USERPROFILE") & "\Desktop", _
            Mid$(relativePath, Len("Desktop\") + 1))

        If FolderExists(testPath) Then
            ResolveWorkbookLocalPath = testPath
            Exit Function
        End If

    End If

    testPath = CombinePath(oneDriveRoot, "Documents\" & relativePath)

    If FolderExists(testPath) Then
        ResolveWorkbookLocalPath = testPath
    End If
End Function

Private Function GetSharePointRelativePath( _
    ByVal urlPath As String) As String

    Dim p As Long
    Dim rel As String

    p = InStr(1, urlPath, "/Documents/", vbTextCompare)

    If p = 0 Then Exit Function

    rel = Mid$(urlPath, p + Len("/Documents/"))
    rel = Replace(rel, "/", "\")
    rel = UrlDecodeBasic(rel)

    GetSharePointRelativePath = rel
End Function

Private Function GetBestOneDriveRoot() As String

    Dim candidates As Collection
    Dim c As Variant

    Set candidates = New Collection

    AddCandidate candidates, Environ$("OneDriveCommercial")
    AddCandidate candidates, Environ$("OneDrive")
    AddCandidate candidates, Environ$("OneDriveConsumer")
    AddCandidate candidates, Environ$("USERPROFILE") & "\OneDrive - Synopsys"
    AddCandidate candidates, Environ$("USERPROFILE") & "\OneDrive - Synopsys, Inc."
    AddCandidate candidates, Environ$("USERPROFILE") & "\OneDrive - Synopsys Inc."

    For Each c In candidates

        If FolderExists(CStr(c)) Then
            GetBestOneDriveRoot = CStr(c)
            Exit Function
        End If

    Next c
End Function

Private Sub AddCandidate( _
    ByRef candidates As Collection, _
    ByVal candidatePath As String)

    If Len(Trim$(candidatePath)) > 0 Then
        candidates.Add Trim$(candidatePath)
    End If
End Sub

Private Function UrlDecodeBasic(ByVal text As String) As String

    Dim s As String

    s = text

    s = Replace(s, "%20", " ")
    s = Replace(s, "%28", "(")
    s = Replace(s, "%29", ")")
    s = Replace(s, "%23", "#")
    s = Replace(s, "%25", "%")
    s = Replace(s, "%26", "&")
    s = Replace(s, "%2B", "+", , , vbTextCompare)
    s = Replace(s, "%2D", "-", , , vbTextCompare)
    s = Replace(s, "%2E", ".", , , vbTextCompare)
    s = Replace(s, "%5F", "_", , , vbTextCompare)

    UrlDecodeBasic = s
End Function

' ============================================================
' Worksheet and export helpers
' ============================================================

Private Sub CreateRunButton(ByVal ws As Worksheet)

    Dim btn As Button
    Dim shp As Shape
    Dim i As Long

    On Error Resume Next

    ' Remove any old macro button / shape in the Control button area.
    For i = ws.Shapes.Count To 1 Step -1

        Set shp = ws.Shapes(i)

        If StrComp(shp.Name, "btnRunCsvAggregation", vbTextCompare) = 0 Then
            shp.Delete

        ElseIf shp.TopLeftCell.Row = 1 _
            And shp.TopLeftCell.Column <= 2 Then

            If shp.Type = msoFormControl Then
                shp.Delete
            End If

        End If

    Next i

    On Error GoTo 0

    Set btn = ws.Buttons.Add( _
                ws.Range("A1").Left, _
                ws.Range("A1").Top, _
                ws.Range("A1:B1").Width, _
                ws.Rows(1).Height + 5)

    With btn
        ' Do not assign .Name because an existing shape name can cause Error 1004.
        .Caption = "Run"
        .OnAction = "'" & ThisWorkbook.Name & "'!RunCsvAggregation"
        .Font.Size = 10
        .Font.Bold = True
    End With

    ws.Rows(1).RowHeight = 25

End Sub

Private Function GetOrCreateSheet( _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then

        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub ClearOutputArea(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Clear
End Sub

Private Sub FormatSheetAtRow( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastCol < 1 Then Exit Sub

    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    If lastRow >= headerRow Then
        ws.Range(ws.Cells(headerRow, 1), _
                 ws.Cells(lastRow, lastCol)).AutoFilter
    End If

    ws.Columns.AutoFit

    For c = 1 To lastCol
        If ws.Columns(c).ColumnWidth > 45 Then
            ws.Columns(c).ColumnWidth = 45
        End If
    Next c
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = ws.Cells.Find( _
                        What:="*", _
                        After:=ws.Cells(1, 1), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious)

    On Error GoTo 0

    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = ws.Cells.Find( _
                        What:="*", _
                        After:=ws.Cells(1, 1), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious)

    On Error GoTo 0

    If foundCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = foundCell.Column
    End If
End Function

Private Function NewDictionary() As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set NewDictionary = d
End Function

Private Function MakeUserKey( _
    ByVal username As String, _
    ByVal email As String, _
    ByVal fullName As String) As String

    If Len(username) > 0 Or Len(email) > 0 Then

        MakeUserKey = _
            LCase$(Trim$(username)) & _
            KEY_SEPARATOR & _
            LCase$(Trim$(email))

    ElseIf Len(fullName) > 0 Then

        MakeUserKey = _
            "FULLNAME" & KEY_SEPARATOR & _
            LCase$(Trim$(fullName))

    Else
        MakeUserKey = ""
    End If
End Function

Private Function CleanText(ByVal value As Variant) As String

    If IsError(value) Then
        CleanText = ""
    ElseIf IsNull(value) Or IsEmpty(value) Then
        CleanText = ""
    Else
        CleanText = Trim$(CStr(value))
    End If
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean

    Dim fso As Object

    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    Exit Function

Failed:
    FolderExists = False
End Function

Private Function FileExists(ByVal filePath As String) As Boolean

    Dim fso As Object

    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Exit Function

Failed:
    FileExists = False
End Function

Private Function CombinePath( _
    ByVal parentPath As String, _
    ByVal childName As String) As String

    If Right$(parentPath, 1) = "\" Then
        CombinePath = parentPath & childName
    Else
        CombinePath = parentPath & "\" & childName
    End If
End Function

Private Function MakeSafeFileName(ByVal fileName As String) As String

    Dim s As String

    s = fileName

    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, Chr$(34), "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")

    MakeSafeFileName = s
End Function

Private Sub CopyFileOverwrite( _
    ByVal sourcePath As String, _
    ByVal destinationPath As String)

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(destinationPath) Then
        fso.DeleteFile destinationPath, True
    End If

    fso.CopyFile sourcePath, destinationPath, True
End Sub
