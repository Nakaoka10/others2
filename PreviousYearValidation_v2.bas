Attribute VB_Name = "PreviousYearValidation"
Option Explicit

Private Const KEY_SEP As String = "|||"
Private Const EXAM_SEP As String = "###"

Private Const SHEET_SUMMARY As String = "Validation_Summary"
Private Const SHEET_MISSING_USER As String = "Validation_MissingUser"
Private Const SHEET_EXTRA_USER As String = "Validation_ExtraUser"
Private Const SHEET_MISSING_COURSE As String = "Validation_MissingCourse"
Private Const SHEET_PASSED_MISMATCH As String = "Validation_PassedMismatch"
Private Const SHEET_NO_EXAM_ACCESS As String = "Validation_NoExamAccess"

Private gStep As String

' ============================================================
' Main validation macro
' ============================================================

Public Sub ValidatePreviousYearAggregation()

    Dim basePath As String
    Dim tracks As Variant
    Dim examMap As Object

    Dim srcUsersByTrack As Object
    Dim srcCoursesByTrack As Object
    Dim srcPassedByTrack As Object

    Dim srcAllUsers As Object
    Dim srcAllCourses As Object
    Dim srcExamUsers As Object

    Dim aggUsersByTrack As Object
    Dim aggCoursesByTrack As Object
    Dim aggPassedByTrack As Object

    Dim missingUsers As Collection
    Dim extraUsers As Collection
    Dim missingCourses As Collection
    Dim passedMismatch As Collection
    Dim noExamAccess As Collection

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gStep = "Resolving workbook folder"

    basePath = ResolveWorkbookLocalPath(ThisWorkbook.Path)

    If Len(basePath) = 0 Then
        MsgBox _
            "Could not resolve the local workbook folder." & vbCrLf & _
            ThisWorkbook.Path, _
            vbExclamation
        GoTo SafeExit
    End If

    tracks = GetTrackNames()
    Set examMap = BuildExamMap()

    gStep = "Creating validation data containers"

    Set srcUsersByTrack = CreateTrackDictionaryContainer(tracks)
    Set srcCoursesByTrack = CreateTrackDictionaryContainer(tracks)
    Set srcPassedByTrack = CreateTrackDictionaryContainer(tracks)

    Set srcAllUsers = NewDictionary()
    Set srcAllCourses = NewDictionary()
    Set srcExamUsers = NewDictionary()

    Set aggUsersByTrack = CreateTrackDictionaryContainer(tracks)
    Set aggCoursesByTrack = CreateTrackDictionaryContainer(tracks)
    Set aggPassedByTrack = CreateTrackDictionaryContainer(tracks)

    Set missingUsers = New Collection
    Set extraUsers = New Collection
    Set missingCourses = New Collection
    Set passedMismatch = New Collection
    Set noExamAccess = New Collection

    gStep = "Reading source xlsx files"

    ReadSourceData _
        basePath, _
        examMap, _
        srcUsersByTrack, _
        srcCoursesByTrack, _
        srcPassedByTrack, _
        srcAllUsers, _
        srcAllCourses, _
        srcExamUsers

    gStep = "Reading aggregated sheets"

    ReadAggregatedData _
        tracks, _
        aggUsersByTrack, _
        aggCoursesByTrack, _
        aggPassedByTrack

    gStep = "Checking missing and extra users"

    CheckUsers _
        tracks, _
        srcUsersByTrack, _
        aggUsersByTrack, _
        missingUsers, _
        extraUsers

    gStep = "Checking missing courses"

    CheckMissingCourses _
        tracks, _
        srcUsersByTrack, _
        srcCoursesByTrack, _
        aggCoursesByTrack, _
        missingCourses

    gStep = "Checking passed mismatches"

    CheckPassedMismatch _
        tracks, _
        srcUsersByTrack, _
        srcPassedByTrack, _
        aggPassedByTrack, _
        passedMismatch

    gStep = "Checking users with no exam access"

    CheckNoExamAccess _
        srcAllUsers, _
        srcAllCourses, _
        srcExamUsers, _
        noExamAccess

    gStep = "Writing validation sheets"

    WriteMissingUserSheet missingUsers
    WriteExtraUserSheet extraUsers
    WriteMissingCourseSheet missingCourses
    WritePassedMismatchSheet passedMismatch
    WriteNoExamAccessSheet noExamAccess

    WriteSummarySheet _
        tracks, _
        srcUsersByTrack, _
        aggUsersByTrack, _
        missingUsers, _
        extraUsers, _
        missingCourses, _
        passedMismatch, _
        noExamAccess

    MsgBox _
        "Validation completed." & vbCrLf & vbCrLf & _
        "Missing users: " & missingUsers.Count & vbCrLf & _
        "Extra users: " & extraUsers.Count & vbCrLf & _
        "Missing courses: " & missingCourses.Count & vbCrLf & _
        "Passed mismatches: " & passedMismatch.Count & vbCrLf & _
        "No exam access users: " & noExamAccess.Count, _
        vbInformation

SafeExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub

ErrorHandler:

    MsgBox _
        "Validation error." & vbCrLf & vbCrLf & _
        "Step: " & gStep & vbCrLf & _
        "Error number: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description, _
        vbCritical

    Resume SafeExit

End Sub

' ============================================================
' Read source data independently
' ============================================================

Private Sub ReadSourceData( _
    ByVal basePath As String, _
    ByVal examMap As Object, _
    ByVal srcUsersByTrack As Object, _
    ByVal srcCoursesByTrack As Object, _
    ByVal srcPassedByTrack As Object, _
    ByVal srcAllUsers As Object, _
    ByVal srcAllCourses As Object, _
    ByVal srcExamUsers As Object)

    Dim dateFolders As Variant
    Dim files As Variant

    Dim i As Long
    Dim j As Long

    Dim folderPath As String
    Dim filePath As String

    dateFolders = GetSortedDateFolders(basePath)

    If IsEmpty(dateFolders) Then
        Err.Raise vbObjectError + 701, , _
            "No YYYY_MM_DD folders were found."
    End If

    For i = LBound(dateFolders) To UBound(dateFolders)

        folderPath = CombinePath(basePath, CStr(dateFolders(i)))
        files = GetSortedXlsxFiles(folderPath)

        If Not IsEmpty(files) Then

            For j = LBound(files) To UBound(files)

                filePath = CombinePath(folderPath, CStr(files(j)))

                gStep = "Reading source file: " & filePath

                ReadOneSourceWorkbook _
                    filePath, _
                    CStr(dateFolders(i)), _
                    examMap, _
                    srcUsersByTrack, _
                    srcCoursesByTrack, _
                    srcPassedByTrack, _
                    srcAllUsers, _
                    srcAllCourses, _
                    srcExamUsers

            Next j

        End If

    Next i

End Sub


Private Sub ReadOneSourceWorkbook( _
    ByVal filePath As String, _
    ByVal sourceDate As String, _
    ByVal examMap As Object, _
    ByVal srcUsersByTrack As Object, _
    ByVal srcCoursesByTrack As Object, _
    ByVal srcPassedByTrack As Object, _
    ByVal srcAllUsers As Object, _
    ByVal srcAllCourses As Object, _
    ByVal srcExamUsers As Object)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim data As Variant

    Dim lastRow As Long
    Dim lastCol As Long

    Dim cUsername As Long
    Dim cEmail As Long
    Dim cFullName As Long
    Dim cCourse As Long
    Dim cStatus As Long

    Dim r As Long

    Dim username As String
    Dim email As String
    Dim fullName As String
    Dim courseTitle As String
    Dim statusText As String

    Dim userKey As String
    Dim trackName As String
    Dim pairKey As String

    Dim users As Object
    Dim courses As Object
    Dim passed As Object

    Dim info As Variant
    Dim sourceFileName As String

    On Error GoTo CleanFail

    Set wb = Workbooks.Open( _
                Filename:=filePath, _
                ReadOnly:=True, _
                UpdateLinks:=0, _
                AddToMru:=False)

    Set ws = wb.Worksheets(1)

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastRow < 2 Or lastCol < 1 Then GoTo CleanExit

    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2

    cUsername = FindHeaderColumn(data, "Username")
    cEmail = FindHeaderColumn(data, "Email")
    cFullName = FindHeaderColumn(data, "Full Name")
    cCourse = FindHeaderColumn(data, "Course title")
    cStatus = FindHeaderColumn(data, "Course Enrollment Status")

    If cCourse = 0 Then GoTo CleanExit

    sourceFileName = GetFileNameOnly(filePath)

    For r = 2 To UBound(data, 1)

        username = GetCellText(data, r, cUsername)
        email = GetCellText(data, r, cEmail)
        fullName = GetCellText(data, r, cFullName)
        courseTitle = GetCellText(data, r, cCourse)
        statusText = GetCellText(data, r, cStatus)

        userKey = MakeUserKey(username, email, fullName)

        If Len(userKey) = 0 Then GoTo NextRow

        ' All source users, regardless of course.
        If Not srcAllUsers.Exists(userKey) Then

            info = Array( _
                username, _
                email, _
                fullName, _
                sourceDate, _
                sourceFileName)

            srcAllUsers.Add userKey, info

        End If

        ' All source courses for No Exam Access review.
        If Len(courseTitle) > 0 Then

            pairKey = userKey & EXAM_SEP & LCase$(courseTitle)

            If Not srcAllCourses.Exists(pairKey) Then
                srcAllCourses.Add pairKey, courseTitle
            End If

        End If

        ' Target exam path.
        If examMap.Exists(LCase$(courseTitle)) Then

            trackName = CStr(examMap(LCase$(courseTitle)))

            Set users = srcUsersByTrack(trackName)
            Set courses = srcCoursesByTrack(trackName)
            Set passed = srcPassedByTrack(trackName)

            If Not users.Exists(userKey) Then

                info = Array( _
                    username, _
                    email, _
                    fullName, _
                    sourceDate, _
                    sourceFileName)

                users.Add userKey, info

            End If

            pairKey = userKey & EXAM_SEP & LCase$(courseTitle)

            If Not courses.Exists(pairKey) Then
                courses.Add pairKey, courseTitle
            End If

            If Not srcExamUsers.Exists(userKey) Then
                srcExamUsers.Add userKey, True
            End If

            If StrComp(statusText, "Completed", vbTextCompare) = 0 Then

                If Not passed.Exists(pairKey) Then
                    passed.Add pairKey, courseTitle
                End If

            End If

        End If

NextRow:
    Next r

CleanExit:

    wb.Close SaveChanges:=False
    Exit Sub

CleanFail:

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0

    Err.Raise Err.Number, , Err.Description

End Sub

' ============================================================
' Read aggregate result sheets
' ============================================================

Private Sub ReadAggregatedData( _
    ByVal tracks As Variant, _
    ByVal aggUsersByTrack As Object, _
    ByVal aggCoursesByTrack As Object, _
    ByVal aggPassedByTrack As Object)

    Dim i As Long
    Dim trackName As String

    For i = LBound(tracks) To UBound(tracks)

        trackName = CStr(tracks(i))

        gStep = "Reading aggregate user sheet: " & trackName
        ReadAggregateUserSheet _
            trackName, _
            aggUsersByTrack(trackName), _
            aggCoursesByTrack(trackName)

        gStep = "Reading aggregate passed sheet: " & trackName
        ReadAggregatePassedSheet _
            trackName, _
            aggPassedByTrack(trackName)

    Next i

End Sub


Private Sub ReadAggregateUserSheet( _
    ByVal trackName As String, _
    ByVal users As Object, _
    ByVal courses As Object)

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long

    Dim cUsername As Long
    Dim cEmail As Long
    Dim cFullName As Long

    Dim r As Long
    Dim c As Long

    Dim username As String
    Dim email As String
    Dim fullName As String
    Dim userKey As String
    Dim courseTitle As String
    Dim pairKey As String

    Set ws = GetSheetIfExists(trackName & "_user")
    If ws Is Nothing Then Exit Sub

    headerRow = FindHeaderRow(ws, "No.")
    If headerRow = 0 Then Exit Sub

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    cUsername = FindHeaderColumnInSheet(ws, headerRow, "Username")
    cEmail = FindHeaderColumnInSheet(ws, headerRow, "Email")
    cFullName = FindHeaderColumnInSheet(ws, headerRow, "Full Name")

    For r = headerRow + 1 To lastRow

        username = CleanText(ws.Cells(r, cUsername).Value)
        email = CleanText(ws.Cells(r, cEmail).Value)
        fullName = CleanText(ws.Cells(r, cFullName).Value)

        userKey = MakeUserKey(username, email, fullName)

        If Len(userKey) = 0 Then GoTo NextRow

        If Not users.Exists(userKey) Then
            users.Add userKey, Array(username, email, fullName)
        End If

        For c = 1 To lastCol

            If LCase$(Left$(CleanText(ws.Cells(headerRow, c).Value), 12)) = _
                LCase$("Course title") Then

                courseTitle = CleanText(ws.Cells(r, c).Value)

                If Len(courseTitle) > 0 Then

                    pairKey = userKey & EXAM_SEP & LCase$(courseTitle)

                    If Not courses.Exists(pairKey) Then
                        courses.Add pairKey, courseTitle
                    End If

                End If

            End If

        Next c

NextRow:
    Next r

End Sub


Private Sub ReadAggregatePassedSheet( _
    ByVal trackName As String, _
    ByVal passed As Object)

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long

    Dim cUsername As Long
    Dim cEmail As Long
    Dim cFullName As Long

    Dim r As Long
    Dim c As Long

    Dim username As String
    Dim email As String
    Dim fullName As String
    Dim userKey As String
    Dim courseTitle As String
    Dim statusText As String
    Dim pairKey As String

    Set ws = GetSheetIfExists(trackName & "_passed")
    If ws Is Nothing Then Exit Sub

    headerRow = FindHeaderRow(ws, "No.")
    If headerRow = 0 Then Exit Sub

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    cUsername = FindHeaderColumnInSheet(ws, headerRow, "Username")
    cEmail = FindHeaderColumnInSheet(ws, headerRow, "Email")
    cFullName = FindHeaderColumnInSheet(ws, headerRow, "Full Name")

    For r = headerRow + 1 To lastRow

        username = CleanText(ws.Cells(r, cUsername).Value)
        email = CleanText(ws.Cells(r, cEmail).Value)
        fullName = CleanText(ws.Cells(r, cFullName).Value)

        userKey = MakeUserKey(username, email, fullName)

        If Len(userKey) = 0 Then GoTo NextRow

        For c = 1 To lastCol

            If LCase$(Left$(CleanText(ws.Cells(headerRow, c).Value), 12)) = _
                LCase$("Course title") Then

                courseTitle = CleanText(ws.Cells(r, c).Value)

                If Len(courseTitle) > 0 Then

                    statusText = ""

                    If c + 1 <= lastCol Then
                        statusText = CleanText(ws.Cells(r, c + 1).Value)
                    End If

                    If StrComp(statusText, "Completed", vbTextCompare) = 0 Then

                        pairKey = userKey & EXAM_SEP & LCase$(courseTitle)

                        If Not passed.Exists(pairKey) Then
                            passed.Add pairKey, courseTitle
                        End If

                    End If

                End If

            End If

        Next c

NextRow:
    Next r

End Sub

' ============================================================
' Checks
' ============================================================

Private Sub CheckUsers( _
    ByVal tracks As Variant, _
    ByVal srcUsersByTrack As Object, _
    ByVal aggUsersByTrack As Object, _
    ByVal missingUsers As Collection, _
    ByVal extraUsers As Collection)

    Dim i As Long
    Dim trackName As String

    Dim src As Object
    Dim agg As Object

    Dim key As Variant
    Dim info As Variant

    For i = LBound(tracks) To UBound(tracks)

        trackName = CStr(tracks(i))

        Set src = srcUsersByTrack(trackName)
        Set agg = aggUsersByTrack(trackName)

        For Each key In src.Keys

            If Not agg.Exists(CStr(key)) Then

                info = src(CStr(key))

                missingUsers.Add Array( _
                    trackName, _
                    CStr(info(0)), _
                    CStr(info(1)), _
                    CStr(info(2)), _
                    CStr(info(3)), _
                    CStr(info(4)))

            End If

        Next key

        For Each key In agg.Keys

            If Not src.Exists(CStr(key)) Then

                info = agg(CStr(key))

                extraUsers.Add Array( _
                    trackName, _
                    CStr(info(0)), _
                    CStr(info(1)), _
                    CStr(info(2)))

            End If

        Next key

    Next i

End Sub


Private Sub CheckMissingCourses( _
    ByVal tracks As Variant, _
    ByVal srcUsersByTrack As Object, _
    ByVal srcCoursesByTrack As Object, _
    ByVal aggCoursesByTrack As Object, _
    ByVal missingCourses As Collection)

    Dim i As Long
    Dim trackName As String

    Dim users As Object
    Dim srcCourses As Object
    Dim aggCourses As Object

    Dim pairKey As Variant
    Dim userKey As String
    Dim courseTitle As String
    Dim info As Variant

    For i = LBound(tracks) To UBound(tracks)

        trackName = CStr(tracks(i))

        Set users = srcUsersByTrack(trackName)
        Set srcCourses = srcCoursesByTrack(trackName)
        Set aggCourses = aggCoursesByTrack(trackName)

        For Each pairKey In srcCourses.Keys

            If Not aggCourses.Exists(CStr(pairKey)) Then

                userKey = Left$(CStr(pairKey), _
                    InStr(1, CStr(pairKey), EXAM_SEP, vbBinaryCompare) - 1)

                courseTitle = CStr(srcCourses(CStr(pairKey)))

                If users.Exists(userKey) Then

                    info = users(userKey)

                    missingCourses.Add Array( _
                        trackName, _
                        CStr(info(0)), _
                        CStr(info(1)), _
                        CStr(info(2)), _
                        courseTitle, _
                        CStr(info(3)), _
                        CStr(info(4)))

                End If

            End If

        Next pairKey

    Next i

End Sub


Private Sub CheckPassedMismatch( _
    ByVal tracks As Variant, _
    ByVal srcUsersByTrack As Object, _
    ByVal srcPassedByTrack As Object, _
    ByVal aggPassedByTrack As Object, _
    ByVal passedMismatch As Collection)

    Dim i As Long
    Dim trackName As String

    Dim users As Object
    Dim srcPassed As Object
    Dim aggPassed As Object

    Dim pairKey As Variant
    Dim userKey As String
    Dim info As Variant
    Dim courseTitle As String

    For i = LBound(tracks) To UBound(tracks)

        trackName = CStr(tracks(i))

        Set users = srcUsersByTrack(trackName)
        Set srcPassed = srcPassedByTrack(trackName)
        Set aggPassed = aggPassedByTrack(trackName)

        For Each pairKey In srcPassed.Keys

            If Not aggPassed.Exists(CStr(pairKey)) Then

                userKey = Left$(CStr(pairKey), _
                    InStr(1, CStr(pairKey), EXAM_SEP, vbBinaryCompare) - 1)

                courseTitle = CStr(srcPassed(CStr(pairKey)))

                If users.Exists(userKey) Then

                    info = users(userKey)

                    passedMismatch.Add Array( _
                        trackName, _
                        CStr(info(0)), _
                        CStr(info(1)), _
                        CStr(info(2)), _
                        courseTitle, _
                        CStr(info(3)), _
                        CStr(info(4)))

                End If

            End If

        Next pairKey

    Next i

End Sub


Private Sub CheckNoExamAccess( _
    ByVal srcAllUsers As Object, _
    ByVal srcAllCourses As Object, _
    ByVal srcExamUsers As Object, _
    ByVal noExamAccess As Collection)

    Dim key As Variant
    Dim info As Variant
    Dim coursesText As String

    For Each key In srcAllUsers.Keys

        If Not srcExamUsers.Exists(CStr(key)) Then

            info = srcAllUsers(CStr(key))
            coursesText = GetAllCoursesForUser(CStr(key), srcAllCourses)

            noExamAccess.Add Array( _
                CStr(info(0)), _
                CStr(info(1)), _
                CStr(info(2)), _
                IIf(IsStudentEmail(CStr(info(1))), "Yes", "No"), _
                coursesText, _
                CStr(info(3)), _
                CStr(info(4)))

        End If

    Next key

End Sub

' ============================================================
' Output validation sheets
' ============================================================

Private Sub WriteMissingUserSheet(ByVal rows As Collection)

    Dim ws As Worksheet

    Set ws = PrepareValidationSheet(SHEET_MISSING_USER)

    ws.Range("A1:F1").Value = Array( _
        "Track", _
        "Username", _
        "Email", _
        "Full Name", _
        "Source Date", _
        "Source File")

    WriteCollectionRows ws, rows, 6
    FormatValidationSheet ws

End Sub


Private Sub WriteExtraUserSheet(ByVal rows As Collection)

    Dim ws As Worksheet

    Set ws = PrepareValidationSheet(SHEET_EXTRA_USER)

    ws.Range("A1:D1").Value = Array( _
        "Track", _
        "Username", _
        "Email", _
        "Full Name")

    WriteCollectionRows ws, rows, 4
    FormatValidationSheet ws

End Sub


Private Sub WriteMissingCourseSheet(ByVal rows As Collection)

    Dim ws As Worksheet

    Set ws = PrepareValidationSheet(SHEET_MISSING_COURSE)

    ws.Range("A1:G1").Value = Array( _
        "Track", _
        "Username", _
        "Email", _
        "Full Name", _
        "Course title", _
        "Source Date", _
        "Source File")

    WriteCollectionRows ws, rows, 7
    FormatValidationSheet ws

End Sub


Private Sub WritePassedMismatchSheet(ByVal rows As Collection)

    Dim ws As Worksheet

    Set ws = PrepareValidationSheet(SHEET_PASSED_MISMATCH)

    ws.Range("A1:G1").Value = Array( _
        "Track", _
        "Username", _
        "Email", _
        "Full Name", _
        "Completed Course title", _
        "Source Date", _
        "Source File")

    WriteCollectionRows ws, rows, 7
    FormatValidationSheet ws

End Sub


Private Sub WriteNoExamAccessSheet(ByVal rows As Collection)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = PrepareValidationSheet(SHEET_NO_EXAM_ACCESS)

    ws.Range("A1:G1").Value = Array( _
        "Username", _
        "Email", _
        "Full Name", _
        "LSTC Student", _
        "Non-Exam Course titles", _
        "First Source Date", _
        "First Source File")

    WriteCollectionRows ws, rows, 7

    ' Highlight likely students with no exam access.
    For r = 2 To rows.Count + 1

        If StrComp(CleanText(ws.Cells(r, 4).Value), "Yes", vbTextCompare) = 0 Then

            ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = _
                RGB(255, 230, 153)

        End If

    Next r

    FormatValidationSheet ws

End Sub


Private Sub WriteSummarySheet( _
    ByVal tracks As Variant, _
    ByVal srcUsersByTrack As Object, _
    ByVal aggUsersByTrack As Object, _
    ByVal missingUsers As Collection, _
    ByVal extraUsers As Collection, _
    ByVal missingCourses As Collection, _
    ByVal passedMismatch As Collection, _
    ByVal noExamAccess As Collection)

    Dim ws As Worksheet
    Dim i As Long
    Dim rowNo As Long
    Dim trackName As String

    Set ws = PrepareValidationSheet(SHEET_SUMMARY)

    ws.Range("A1:G1").Value = Array( _
        "Track", _
        "Source Users", _
        "Aggregated Users", _
        "Missing Users", _
        "Extra Users", _
        "Missing Courses", _
        "Passed Mismatch")

    rowNo = 2

    For i = LBound(tracks) To UBound(tracks)

        trackName = CStr(tracks(i))

        ws.Cells(rowNo, 1).Value = trackName
        ws.Cells(rowNo, 2).Value = srcUsersByTrack(trackName).Count
        ws.Cells(rowNo, 3).Value = aggUsersByTrack(trackName).Count
        ws.Cells(rowNo, 4).Value = CountRowsByTrack(missingUsers, trackName)
        ws.Cells(rowNo, 5).Value = CountRowsByTrack(extraUsers, trackName)
        ws.Cells(rowNo, 6).Value = CountRowsByTrack(missingCourses, trackName)
        ws.Cells(rowNo, 7).Value = CountRowsByTrack(passedMismatch, trackName)

        rowNo = rowNo + 1

    Next i

    rowNo = rowNo + 2

    ws.Cells(rowNo, 1).Value = "No Exam Access - All Users"
    ws.Cells(rowNo, 2).Value = noExamAccess.Count

    rowNo = rowNo + 1

    ws.Cells(rowNo, 1).Value = "No Exam Access - LSTC Students"
    ws.Cells(rowNo, 2).Value = CountStudentNoExamRows(noExamAccess)

    rowNo = rowNo + 3

    WriteValidationExplanation ws, rowNo

    FormatValidationSheet ws

End Sub


Private Sub WriteValidationExplanation( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long)

    Dim r As Long

    r = startRow

    ws.Cells(r, 1).Value = "Validation Item"
    ws.Cells(r, 2).Value = "Description"

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
        .Font.Bold = True
    End With

    r = r + 1
    ws.Cells(r, 1).Value = "Source Users"
    ws.Cells(r, 2).Value = _
        "Number of unique users found in the original source xlsx files for the target exams in each track."

    r = r + 1
    ws.Cells(r, 1).Value = "Aggregated Users"
    ws.Cells(r, 2).Value = _
        "Number of unique users listed in the aggregated *_user sheet for each track."

    r = r + 1
    ws.Cells(r, 1).Value = "Missing Users"
    ws.Cells(r, 2).Value = _
        "Users found in the original source xlsx files for a target exam but not found in the corresponding aggregated *_user sheet. A value of 0 means no user omission was detected."

    r = r + 1
    ws.Cells(r, 1).Value = "Extra Users"
    ws.Cells(r, 2).Value = _
        "Users found in the aggregated *_user sheet but not found in the original source data for the target exams in that track. A value of 0 means no unexpected user was detected."

    r = r + 1
    ws.Cells(r, 1).Value = "Missing Courses"
    ws.Cells(r, 2).Value = _
        "Target exam records that exist in the original source data but are missing from the corresponding aggregated user record. A value of 0 means no target exam record omission was detected."

    r = r + 1
    ws.Cells(r, 1).Value = "Passed Mismatch"
    ws.Cells(r, 2).Value = _
        "Target exams marked as Completed in the original source data but not found in the corresponding aggregated *_passed sheet. A value of 0 means no completed-exam omission was detected."

    r = r + 1
    ws.Cells(r, 1).Value = "No Exam Access - All Users"
    ws.Cells(r, 2).Value = _
        "Users found in the original source xlsx files who do not appear in any of the defined target exam Course titles. This check is intended to identify users who may not have accessed any target exam."

    r = r + 1
    ws.Cells(r, 1).Value = "No Exam Access - LSTC Students"
    ws.Cells(r, 2).Value = _
        "Among users with no target exam access, the number whose Email matches the LSTC student format: 8 digits followed by @lstc.adip.jp."

    r = r + 2
    ws.Cells(r, 1).Value = "Overall interpretation"
    ws.Cells(r, 2).Value = _
        "If Missing Users, Extra Users, Missing Courses, Passed Mismatch, and No Exam Access are all 0, no discrepancy was detected by the automated validation checks."

    ws.Columns(2).WrapText = True

End Sub


Private Function PrepareValidationSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    Set ws = GetSheetIfExists(sheetName)

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = sheetName

    Else

        ws.Cells.Clear
        DeleteAllShapes ws

    End If

    Set PrepareValidationSheet = ws

End Function


Private Sub WriteCollectionRows( _
    ByVal ws As Worksheet, _
    ByVal rows As Collection, _
    ByVal colCount As Long)

    Dim output() As Variant
    Dim i As Long
    Dim j As Long
    Dim item As Variant

    If rows.Count = 0 Then Exit Sub

    ReDim output(1 To rows.Count, 1 To colCount)

    For i = 1 To rows.Count

        item = rows(i)

        For j = 0 To colCount - 1
            output(i, j + 1) = item(j)
        Next j

    Next i

    ws.Cells(2, 1).Resize(rows.Count, colCount).Value = output

End Sub


Private Sub FormatValidationSheet(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastCol < 1 Then Exit Sub

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    If lastRow >= 1 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    End If

    ws.Columns.AutoFit

    For c = 1 To lastCol
        If ws.Columns(c).ColumnWidth > 55 Then
            ws.Columns(c).ColumnWidth = 55
        End If
    Next c

End Sub

' ============================================================
' Helpers for No Exam Access
' ============================================================

Private Function GetAllCoursesForUser( _
    ByVal userKey As String, _
    ByVal srcAllCourses As Object) As String

    Dim key As Variant
    Dim prefix As String
    Dim result As String
    Dim title As String

    prefix = userKey & EXAM_SEP

    For Each key In srcAllCourses.Keys

        If Left$(CStr(key), Len(prefix)) = prefix Then

            title = CStr(srcAllCourses(CStr(key)))

            If Len(result) > 0 Then
                result = result & " | "
            End If

            result = result & title

        End If

    Next key

    GetAllCoursesForUser = result

End Function


Private Function IsStudentEmail(ByVal email As String) As Boolean

    IsStudentEmail = _
        (LCase$(Trim$(email)) Like "########@lstc.adip.jp")

End Function


Private Function CountStudentNoExamRows(ByVal rows As Collection) As Long

    Dim i As Long
    Dim item As Variant
    Dim n As Long

    For i = 1 To rows.Count

        item = rows(i)

        If StrComp(CStr(item(3)), "Yes", vbTextCompare) = 0 Then
            n = n + 1
        End If

    Next i

    CountStudentNoExamRows = n

End Function

' ============================================================
' Summary helpers
' ============================================================

Private Function CountRowsByTrack( _
    ByVal rows As Collection, _
    ByVal trackName As String) As Long

    Dim i As Long
    Dim item As Variant
    Dim n As Long

    For i = 1 To rows.Count

        item = rows(i)

        If StrComp(CStr(item(0)), trackName, vbTextCompare) = 0 Then
            n = n + 1
        End If

    Next i

    CountRowsByTrack = n

End Function

' ============================================================
' Track and exam configuration
' ============================================================

Private Function GetTrackNames() As Variant

    GetTrackNames = Array( _
        "Pre-Requisite", _
        "WFD-Design_for_Test", _
        "WFD-Design_Verification", _
        "WFD-Physical_Design", _
        "WFD-RTL_Synthesis", _
        "WFD-AMS")

End Function


Private Function BuildExamMap() As Object

    Dim d As Object

    Set d = NewDictionary()

    AddExam d, "Purple Certification: ASIC Design Flow Exam", "Pre-Requisite"
    AddExam d, "Purple Certification: Digital Design Fundamentals Exam", "Pre-Requisite"
    AddExam d, "Purple Certification: CMOS Fundamentals Exam", "Pre-Requisite"
    AddExam d, "Purple Certification: Very Deep Submicron (VDSM) Fundamentals Exam", "Pre-Requisite"
    AddExam d, "Purple Certification: VLSI Basics Exam", "Pre-Requisite"

    AddExam d, "TestMAX ATPG Exam", "WFD-Design_for_Test"
    AddExam d, "Fusion Compiler: DFT Synthesis Exam", "WFD-Design_for_Test"
    AddExam d, "TestMAX Advisor Exam", "WFD-Design_for_Test"
    AddExam d, "TestMAX DFT Exam", "WFD-Design_for_Test"

    AddExam d, "SystemVerilog Assertions Exam", "WFD-Design_Verification"
    AddExam d, "SystemVerilog Verification using UVM Exam", "WFD-Design_Verification"
    AddExam d, "SystemVerilog Testbench Exam", "WFD-Design_Verification"

    AddExam d, "Fusion Compiler: Design Implementation Exam", "WFD-Physical_Design"
    AddExam d, "Fusion Compiler: Design Creation and Synthesis Exam", "WFD-Physical_Design"

    AddExam d, "SystemVerilog for RTL Design Exam", "WFD-RTL_Synthesis"
    AddExam d, "Design Compiler NXT: RTL Synthesis Exam", "WFD-RTL_Synthesis"
    AddExam d, "Design Compiler NXT: Low Power Exam", "WFD-RTL_Synthesis"
    AddExam d, "Fusion Compiler: UPF Fundamentals Exam", "WFD-RTL_Synthesis"

    AddExam d, "Custom Compiler: Basic Layout Design Exam", "WFD-AMS"
    AddExam d, "Custom Compiler: Introduction to Platform Exam", "WFD-AMS"
    AddExam d, "Custom Compiler: Schematic Entry Exam", "WFD-AMS"
    AddExam d, "PrimeWave Design Environment Exam", "WFD-AMS"

    Set BuildExamMap = d

End Function


Private Sub AddExam( _
    ByVal d As Object, _
    ByVal examTitle As String, _
    ByVal trackName As String)

    d(LCase$(examTitle)) = trackName

End Sub

' ============================================================
' Generic dictionary helpers
' ============================================================

Private Function NewDictionary() As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set NewDictionary = d

End Function


Private Function CreateTrackDictionaryContainer( _
    ByVal tracks As Variant) As Object

    Dim outer As Object
    Dim inner As Object
    Dim i As Long

    Set outer = NewDictionary()

    For i = LBound(tracks) To UBound(tracks)

        Set inner = NewDictionary()
        outer.Add CStr(tracks(i)), inner

    Next i

    Set CreateTrackDictionaryContainer = outer

End Function

' ============================================================
' Header, identity, and sheet helpers
' ============================================================

Private Function MakeUserKey( _
    ByVal username As String, _
    ByVal email As String, _
    ByVal fullName As String) As String

    If Len(username) > 0 Or Len(email) > 0 Then

        MakeUserKey = _
            LCase$(Trim$(username)) & _
            KEY_SEP & _
            LCase$(Trim$(email))

    ElseIf Len(fullName) > 0 Then

        MakeUserKey = _
            "FULLNAME" & KEY_SEP & _
            LCase$(Trim$(fullName))

    Else

        MakeUserKey = ""

    End If

End Function


Private Function GetCellText( _
    ByVal data As Variant, _
    ByVal rowNo As Long, _
    ByVal colNo As Long) As String

    If colNo <= 0 Then
        GetCellText = ""
    Else
        GetCellText = CleanText(data(rowNo, colNo))
    End If

End Function


Private Function FindHeaderColumn( _
    ByVal data As Variant, _
    ByVal headerName As String) As Long

    Dim c As Long
    Dim s As String

    For c = 1 To UBound(data, 2)

        s = CleanText(data(1, c))
        s = Replace(s, ChrW(&HFEFF), "")

        If StrComp(Trim$(s), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If

    Next c

End Function


Private Function FindHeaderRow( _
    ByVal ws As Worksheet, _
    ByVal headerText As String) As Long

    Dim r As Long
    Dim maxSearch As Long

    maxSearch = WorksheetFunction.Min(20, LastUsedRow(ws))

    For r = 1 To maxSearch

        If StrComp( _
            CleanText(ws.Cells(r, 1).Value), _
            headerText, _
            vbTextCompare) = 0 Then

            FindHeaderRow = r
            Exit Function

        End If

    Next r

End Function


Private Function FindHeaderColumnInSheet( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal headerText As String) As Long

    Dim c As Long
    Dim lastCol As Long

    lastCol = LastUsedColumn(ws)

    For c = 1 To lastCol

        If StrComp( _
            CleanText(ws.Cells(headerRow, c).Value), _
            headerText, _
            vbTextCompare) = 0 Then

            FindHeaderColumnInSheet = c
            Exit Function

        End If

    Next c

End Function


Private Function GetSheetIfExists(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetSheetIfExists = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

End Function


Private Sub DeleteAllShapes(ByVal ws As Worksheet)

    Dim i As Long

    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    On Error GoTo 0

End Sub

' ============================================================
' Source folder enumeration
' ============================================================

Private Function GetSortedDateFolders( _
    ByVal rootPath As String) As Variant

    Dim fso As Object
    Dim rootFolder As Object
    Dim subFolder As Object

    Dim arr() As String
    Dim count As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(rootPath)

    For Each subFolder In rootFolder.SubFolders

        If IsDateFolderName(CStr(subFolder.Name)) Then

            count = count + 1
            ReDim Preserve arr(1 To count)
            arr(count) = CStr(subFolder.Name)

        End If

    Next subFolder

    If count = 0 Then

        GetSortedDateFolders = Empty
        Exit Function

    End If

    SortStringArray arr
    GetSortedDateFolders = arr

End Function


Private Function GetSortedXlsxFiles( _
    ByVal folderPath As String) As Variant

    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object

    Dim arr() As String
    Dim count As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(folderPath)

    For Each fileObj In folderObj.Files

        If LCase$(fso.GetExtensionName(CStr(fileObj.Name))) = "xlsx" Then

            count = count + 1
            ReDim Preserve arr(1 To count)
            arr(count) = CStr(fileObj.Name)

        End If

    Next fileObj

    If count = 0 Then

        GetSortedXlsxFiles = Empty
        Exit Function

    End If

    SortStringArray arr
    GetSortedXlsxFiles = arr

End Function


Private Function IsDateFolderName(ByVal folderName As String) As Boolean

    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long
    Dim dt As Date

    On Error GoTo InvalidDate

    If Len(folderName) <> 10 Then Exit Function
    If Mid$(folderName, 5, 1) <> "_" Then Exit Function
    If Mid$(folderName, 8, 1) <> "_" Then Exit Function

    If Not IsNumeric(Left$(folderName, 4)) Then Exit Function
    If Not IsNumeric(Mid$(folderName, 6, 2)) Then Exit Function
    If Not IsNumeric(Right$(folderName, 2)) Then Exit Function

    yyyy = CLng(Left$(folderName, 4))
    mm = CLng(Mid$(folderName, 6, 2))
    dd = CLng(Right$(folderName, 2))

    dt = DateSerial(yyyy, mm, dd)

    If Format$(dt, "yyyy_mm_dd") <> folderName Then Exit Function

    IsDateFolderName = True
    Exit Function

InvalidDate:

    IsDateFolderName = False

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
' OneDrive / SharePoint path resolution
' ============================================================

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

    testPath = CombinePath( _
        oneDriveRoot, _
        "Documents\" & relativePath)

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
' Basic filesystem and worksheet helpers
' ============================================================

Private Function FolderExists(ByVal folderPath As String) As Boolean

    Dim fso As Object

    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")

    FolderExists = fso.FolderExists(folderPath)
    Exit Function

Failed:

    FolderExists = False

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


Private Function GetFileNameOnly(ByVal fullPath As String) As String

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameOnly = fso.GetFileName(fullPath)

End Function


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


Private Function CleanText(ByVal value As Variant) As String

    If IsError(value) Then

        CleanText = ""

    ElseIf IsNull(value) Or IsEmpty(value) Then

        CleanText = ""

    Else

        CleanText = Trim$(CStr(value))

    End If

End Function
