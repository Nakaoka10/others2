Attribute VB_Name = "PreviousYearTrackAggregation"
Option Explicit

Private Const HEADER_ROW As Long = 3
Private Const KEY_SEPARATOR As String = "|||"
Private Const EXAM_SEPARATOR As String = "###"
Private Const SHEET_SUFFIX_USER As String = "_user"
Private Const SHEET_SUFFIX_PASSED As String = "_passed"
Private gCurrentStep As String

Public Sub SetupPreviousYearAggregation()
    Dim trackNames As Variant, i As Long
    Dim wsUser As Worksheet, wsPassed As Worksheet
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    trackNames = GetTrackNames()
    For i = LBound(trackNames) To UBound(trackNames)
        Set wsUser = GetOrCreateSheet(CStr(trackNames(i)) & SHEET_SUFFIX_USER)
        Set wsPassed = GetOrCreateSheet(CStr(trackNames(i)) & SHEET_SUFFIX_PASSED)
        ClearOutputArea wsUser
        ClearOutputArea wsPassed
        WriteUserHeaders wsUser, 0
        WritePassedHeaders _
            wsPassed, _
            GetMaxExamCountForTrack(CStr(trackNames(i))), _
            GetPassedHeaderRow(GetMaxExamCountForTrack(CStr(trackNames(i))))
        CreateRunButton wsUser
        FormatSheet wsUser
        FormatSheet wsPassed
    Next i
    Application.ScreenUpdating = True
    MsgBox "Setup completed." & vbCrLf & vbCrLf & _
           "Use the Run Previous Year Aggregation button on any *_user sheet.", vbInformation
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Setup error." & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
End Sub

Public Sub RunPreviousYearAggregation()
    Dim basePath As String, dateFolders As Variant, xlsxFiles As Variant
    Dim trackNames As Variant, courseMap As Object
    Dim userDataByTrack As Object, userOrderByTrack As Object, userCoursesByTrack As Object
    Dim passedDataByTrack As Object, passedOrderByTrack As Object
    Dim passedExamDataByTrack As Object, passedExamOrderByTrack As Object
    Dim i As Long, j As Long
    Dim dateFolderName As String, dateFolderPath As String, xlsxFilePath As String
    Dim processedFileCount As Long, skippedFileCount As Long
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gCurrentStep = "Resolving workbook folder"
    basePath = ResolveWorkbookLocalPath(ThisWorkbook.Path)
    If Len(basePath) = 0 Then
        MsgBox "Could not resolve the local folder for this workbook." & vbCrLf & vbCrLf & _
               "Workbook path:" & vbCrLf & ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "Make sure the workbook folder is synchronized to this PC.", vbExclamation
        GoTo SafeExit
    End If

    gCurrentStep = "Preparing configuration"
    trackNames = GetTrackNames()
    Set courseMap = BuildCourseToTrackMap()
    Set userDataByTrack = CreateObject("Scripting.Dictionary")
    Set userOrderByTrack = CreateObject("Scripting.Dictionary")
    Set userCoursesByTrack = CreateObject("Scripting.Dictionary")
    Set passedDataByTrack = CreateObject("Scripting.Dictionary")
    Set passedOrderByTrack = CreateObject("Scripting.Dictionary")
    Set passedExamDataByTrack = CreateObject("Scripting.Dictionary")
    Set passedExamOrderByTrack = CreateObject("Scripting.Dictionary")

    InitializeTrackContainers trackNames, userDataByTrack, userOrderByTrack, userCoursesByTrack, _
                              passedDataByTrack, passedOrderByTrack, passedExamDataByTrack, passedExamOrderByTrack

    gCurrentStep = "Reading date folders"
    dateFolders = GetSortedDateFolders(basePath)
    If IsEmpty(dateFolders) Then
        MsgBox "No YYYY_MM_DD date folders were found under:" & vbCrLf & basePath, vbExclamation
        GoTo SafeExit
    End If

    For i = LBound(dateFolders) To UBound(dateFolders)
        dateFolderName = CStr(dateFolders(i))
        dateFolderPath = CombinePath(basePath, dateFolderName)
        gCurrentStep = "Reading xlsx files in: " & dateFolderPath
        xlsxFiles = GetSortedXlsxFiles(dateFolderPath)
        If Not IsEmpty(xlsxFiles) Then
            For j = LBound(xlsxFiles) To UBound(xlsxFiles)
                xlsxFilePath = CombinePath(dateFolderPath, CStr(xlsxFiles(j)))
                gCurrentStep = "Processing workbook: " & xlsxFilePath
                If ProcessSourceWorkbook(xlsxFilePath, courseMap, userDataByTrack, userOrderByTrack, userCoursesByTrack, _
                                         passedDataByTrack, passedOrderByTrack, passedExamDataByTrack, passedExamOrderByTrack) Then
                    processedFileCount = processedFileCount + 1
                Else
                    skippedFileCount = skippedFileCount + 1
                End If
            Next j
        End If
    Next i

    gCurrentStep = "Writing output sheets"
    OutputAllTracks trackNames, userDataByTrack, userOrderByTrack, userCoursesByTrack, _
                    passedDataByTrack, passedOrderByTrack, passedExamDataByTrack, passedExamOrderByTrack

    MsgBox "Previous year aggregation completed." & vbCrLf & vbCrLf & _
           "Processed xlsx files: " & processedFileCount & vbCrLf & _
           "Skipped xlsx files: " & skippedFileCount, vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred." & vbCrLf & vbCrLf & _
           "Step: " & gCurrentStep & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
    Resume SafeExit
End Sub

Private Function ProcessSourceWorkbook(ByVal filePath As String, ByVal courseMap As Object, _
    ByVal userDataByTrack As Object, ByVal userOrderByTrack As Object, ByVal userCoursesByTrack As Object, _
    ByVal passedDataByTrack As Object, ByVal passedOrderByTrack As Object, _
    ByVal passedExamDataByTrack As Object, ByVal passedExamOrderByTrack As Object) As Boolean

    Dim wb As Workbook, ws As Worksheet, data As Variant
    Dim lastRow As Long, lastCol As Long
    Dim colUsername As Long, colEmail As Long, colFullName As Long, colCourseTitle As Long
    Dim colEnrollmentDate As Long, colEnrollmentEndDate As Long, colStatus As Long, colFinalScore As Long
    Dim r As Long
    Dim username As String, email As String, fullName As String, courseTitle As String, courseStatus As String
    Dim enrollmentDate As Variant, enrollmentEndDate As Variant, finalScore As Variant
    Dim trackName As String, userKey As String, examKey As String
    Dim users As Object, userOrder As Collection, userCourses As Object
    Dim passedUsers As Object, passedOrder As Collection, passedExamData As Object, passedExamOrder As Object
    Dim userInfo As Variant, passedInfo As Variant, examInfo As Variant
    Dim courseOrder As Collection, passedOrderCollection As Collection

    On Error GoTo FileError
    Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Set ws = wb.Worksheets(1)
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)
    If lastRow < 2 Or lastCol < 1 Then GoTo FileError
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2

    colUsername = FindHeaderColumn(data, "Username")
    colEmail = FindHeaderColumn(data, "Email")
    colFullName = FindHeaderColumn(data, "Full Name")
    colCourseTitle = FindHeaderColumn(data, "Course title")
    colEnrollmentDate = FindHeaderColumn(data, "Enrollment Date")
    colEnrollmentEndDate = FindHeaderColumn(data, "Enrollment End Date")
    colStatus = FindHeaderColumn(data, "Course Enrollment Status")
    colFinalScore = FindHeaderColumn(data, "Final Score")

    If colCourseTitle = 0 Then GoTo FileError

    For r = 2 To UBound(data, 1)
        courseTitle = CleanText(data(r, colCourseTitle))
        If Len(courseTitle) = 0 Then GoTo NextRow
        If Not courseMap.Exists(LCase$(courseTitle)) Then GoTo NextRow

        trackName = CStr(courseMap(LCase$(courseTitle)))
        username = GetCellText(data, r, colUsername)
        email = GetCellText(data, r, colEmail)
        fullName = GetCellText(data, r, colFullName)
        enrollmentDate = GetCellValue(data, r, colEnrollmentDate)
        enrollmentEndDate = GetCellValue(data, r, colEnrollmentEndDate)
        userKey = MakeUserKey(username, email, fullName)
        If Len(userKey) = 0 Then GoTo NextRow

        Set users = userDataByTrack(trackName)
        Set userOrder = userOrderByTrack(trackName)
        Set userCourses = userCoursesByTrack(trackName)

        If Not users.Exists(userKey) Then
            ' Keep the first Enrollment Date / End Date found for this user.
            ' Date folders are processed from oldest to newest.
            userInfo = Array(username, email, fullName, enrollmentDate, enrollmentEndDate)
            users.Add userKey, userInfo
            userOrder.Add userKey
            Set courseOrder = New Collection
            userCourses.Add userKey, courseOrder
        Else
            userInfo = users(userKey)
            If Len(username) > 0 Then userInfo(0) = username
            If Len(email) > 0 Then userInfo(1) = email
            If Len(fullName) > 0 Then userInfo(2) = fullName
            ' userInfo(3) and userInfo(4) intentionally remain unchanged.
            users(userKey) = userInfo
        End If

        Set courseOrder = userCourses(userKey)
        examKey = userKey & EXAM_SEPARATOR & LCase$(courseTitle)
        If Not userCourses.Exists(examKey) Then
            userCourses.Add examKey, Array(courseTitle, enrollmentDate, enrollmentEndDate)
            courseOrder.Add courseTitle
        Else
            userCourses(examKey) = Array(courseTitle, enrollmentDate, enrollmentEndDate)
        End If

        courseStatus = GetCellText(data, r, colStatus)
        If StrComp(courseStatus, "Completed", vbTextCompare) = 0 Then
            finalScore = GetCellValue(data, r, colFinalScore)
            Set passedUsers = passedDataByTrack(trackName)
            Set passedOrder = passedOrderByTrack(trackName)
            Set passedExamData = passedExamDataByTrack(trackName)
            Set passedExamOrder = passedExamOrderByTrack(trackName)

            If Not passedUsers.Exists(userKey) Then
                passedInfo = Array(username, email, fullName, enrollmentDate, enrollmentEndDate)
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
                examInfo = Array(courseTitle, "Completed", finalScore)
                passedExamData.Add examKey, examInfo
                Set passedOrderCollection = passedExamOrder(userKey)
                passedOrderCollection.Add courseTitle
            Else
                passedExamData(examKey) = Array(courseTitle, "Completed", finalScore)
            End If
        End If
NextRow:
    Next r

    wb.Close SaveChanges:=False
    ProcessSourceWorkbook = True
    Exit Function

FileError:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    ProcessSourceWorkbook = False
End Function

Private Sub OutputAllTracks(ByVal trackNames As Variant, ByVal userDataByTrack As Object, _
    ByVal userOrderByTrack As Object, ByVal userCoursesByTrack As Object, _
    ByVal passedDataByTrack As Object, ByVal passedOrderByTrack As Object, _
    ByVal passedExamDataByTrack As Object, ByVal passedExamOrderByTrack As Object)

    Dim i As Long, trackName As String
    For i = LBound(trackNames) To UBound(trackNames)
        trackName = CStr(trackNames(i))
        OutputUserTrack trackName, userDataByTrack(trackName), userOrderByTrack(trackName), userCoursesByTrack(trackName)
        OutputPassedTrack trackName, passedDataByTrack(trackName), passedOrderByTrack(trackName), _
                          passedExamDataByTrack(trackName), passedExamOrderByTrack(trackName), _
                          GetMaxExamCountForTrack(trackName)
    Next i
End Sub

Private Sub OutputUserTrack(ByVal trackName As String, ByVal users As Object, _
    ByVal userOrder As Collection, ByVal userCourses As Object)

    Dim ws As Worksheet, i As Long, j As Long, c As Long
    Dim key As String, examKey As String, courseTitle As String
    Dim info As Variant, courseInfo As Variant, courseOrder As Collection
    Dim maxCourseCount As Long, totalColumns As Long, outputData() As Variant
    Dim nonStudentCount As Long

    Set ws = GetOrCreateSheet(trackName & SHEET_SUFFIX_USER)
    ClearOutputArea ws

    For i = 1 To userOrder.Count
        key = CStr(userOrder(i))
        Set courseOrder = userCourses(key)
        If courseOrder.Count > maxCourseCount Then maxCourseCount = courseOrder.Count
    Next i

    nonStudentCount = CountNonStudents(users, userOrder)

    WriteUserSummary ws, users.Count, nonStudentCount
    WriteUserHeaders ws, maxCourseCount
    CreateRunButton ws

    If users.Count = 0 Then
        FormatSheetAtRow ws, HEADER_ROW
        Exit Sub
    End If

    ' No., Username, Email, Full Name, Enrollment Date,
    ' Enrollment End Date, then Course title1, Course title2, ...
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

    ws.Cells(HEADER_ROW + 1, 1).Resize(users.Count, totalColumns).Value = outputData

    ws.Columns(5).NumberFormat = "yyyy/mm/dd"
    ws.Columns(6).NumberFormat = "yyyy/mm/dd"

    HighlightNonStudentEmails ws, users.Count

    FormatSheetAtRow ws, HEADER_ROW

End Sub

Private Sub OutputPassedTrack(ByVal trackName As String, ByVal passedUsers As Object, _
    ByVal passedOrder As Collection, ByVal passedExamData As Object, ByVal passedExamOrder As Object, _
    ByVal maxExamCount As Long)

    Dim ws As Worksheet, i As Long, j As Long, c As Long
    Dim key As String, examKey As String, courseTitle As String
    Dim info As Variant, examInfo As Variant, examOrder As Collection
    Dim outputData() As Variant, totalColumns As Long
    Dim passedCountBuckets() As Long, examCountForUser As Long
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

    ws.Cells(passedHeaderRow + 1, 1) _
        .Resize(passedUsers.Count, totalColumns).Value = outputData

    ws.Columns(5).NumberFormat = "yyyy/mm/dd"
    ws.Columns(6).NumberFormat = "yyyy/mm/dd"

    FormatSheetAtRow ws, passedHeaderRow

End Sub

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


Private Function GetPassedHeaderRow( _
    ByVal maxExamCount As Long) As Long

    ' One row for each exam-count bucket, one row for total,
    ' then one blank row before the header.
    GetPassedHeaderRow = maxExamCount + 3

End Function


Private Function IsStudentEmail(ByVal email As String) As Boolean

    Dim s As String

    s = LCase$(Trim$(email))

    IsStudentEmail = (s Like "########@lstc.adip.jp")

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

Private Function GetTrackNames() As Variant
    GetTrackNames = Array("Pre-Requisite", "WFD-Design_for_Test", "WFD-Design_Verification", _
                          "WFD-Physical_Design", "WFD-RTL_Synthesis", "WFD-AMS")
End Function

Private Function BuildCourseToTrackMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    AddCourse d, "Purple Certification: ASIC Design Flow Exam", "Pre-Requisite"
    AddCourse d, "Purple Certification: Digital Design Fundamentals Exam", "Pre-Requisite"
    AddCourse d, "Purple Certification: CMOS Fundamentals Exam", "Pre-Requisite"
    AddCourse d, "Purple Certification: Very Deep Submicron (VDSM) Fundamentals Exam", "Pre-Requisite"
    AddCourse d, "Purple Certification: VLSI Basics Exam", "Pre-Requisite"

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

Private Sub AddCourse(ByVal d As Object, ByVal courseTitle As String, ByVal trackName As String)
    d(LCase$(courseTitle)) = trackName
End Sub

Private Function GetMaxExamCountForTrack(ByVal trackName As String) As Long
    Select Case trackName
        Case "Pre-Requisite": GetMaxExamCountForTrack = 5
        Case "WFD-Design_for_Test": GetMaxExamCountForTrack = 4
        Case "WFD-Design_Verification": GetMaxExamCountForTrack = 3
        Case "WFD-Physical_Design": GetMaxExamCountForTrack = 2
        Case "WFD-RTL_Synthesis": GetMaxExamCountForTrack = 4
        Case "WFD-AMS": GetMaxExamCountForTrack = 4
        Case Else: GetMaxExamCountForTrack = 0
    End Select
End Function

Private Sub InitializeTrackContainers(ByVal trackNames As Variant, _
    ByVal userDataByTrack As Object, ByVal userOrderByTrack As Object, ByVal userCoursesByTrack As Object, _
    ByVal passedDataByTrack As Object, ByVal passedOrderByTrack As Object, _
    ByVal passedExamDataByTrack As Object, ByVal passedExamOrderByTrack As Object)

    Dim i As Long, trackName As String
    Dim d As Object, c As Collection

    For i = LBound(trackNames) To UBound(trackNames)
        trackName = CStr(trackNames(i))
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare: userDataByTrack.Add trackName, d
        Set c = New Collection: userOrderByTrack.Add trackName, c
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare: userCoursesByTrack.Add trackName, d
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare: passedDataByTrack.Add trackName, d
        Set c = New Collection: passedOrderByTrack.Add trackName, c
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare: passedExamDataByTrack.Add trackName, d
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare: passedExamOrderByTrack.Add trackName, d
    Next i
End Sub

Private Function GetSortedDateFolders(ByVal rootPath As String) As Variant
    Dim fso As Object, rootFolder As Object, subFolder As Object
    Dim arr() As String, count As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(rootPath)
    For Each subFolder In rootFolder.SubFolders
        If IsDateFolderName(CStr(subFolder.Name)) Then
            count = count + 1
            ReDim Preserve arr(1 To count)
            arr(count) = CStr(subFolder.Name)
        End If
    Next subFolder
    If count = 0 Then GetSortedDateFolders = Empty: Exit Function
    SortStringArray arr
    GetSortedDateFolders = arr
End Function

Private Function GetSortedXlsxFiles(ByVal folderPath As String) As Variant
    Dim fso As Object, folderObj As Object, fileObj As Object
    Dim arr() As String, count As Long, ext As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(folderPath)
    For Each fileObj In folderObj.Files
        ext = LCase$(fso.GetExtensionName(CStr(fileObj.Name)))
        If ext = "xlsx" Then
            count = count + 1
            ReDim Preserve arr(1 To count)
            arr(count) = CStr(fileObj.Name)
        End If
    Next fileObj
    If count = 0 Then GetSortedXlsxFiles = Empty: Exit Function
    SortStringArray arr
    GetSortedXlsxFiles = arr
End Function

Private Function IsDateFolderName(ByVal folderName As String) As Boolean
    Dim yyyy As Long, mm As Long, dd As Long, dt As Date
    On Error GoTo InvalidDate
    If Len(folderName) <> 10 Then Exit Function
    If Mid$(folderName, 5, 1) <> "_" Or Mid$(folderName, 8, 1) <> "_" Then Exit Function
    If Not IsNumeric(Left$(folderName, 4)) Or Not IsNumeric(Mid$(folderName, 6, 2)) Or Not IsNumeric(Right$(folderName, 2)) Then Exit Function
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
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                tmp = CStr(arr(i)): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function ResolveWorkbookLocalPath(ByVal workbookPath As String) As String
    Dim localPath As String, relativePath As String, oneDriveRoot As String
    If Len(workbookPath) = 0 Then Exit Function
    If InStr(1, workbookPath, "://", vbTextCompare) = 0 Then
        ResolveWorkbookLocalPath = workbookPath
        Exit Function
    End If
    relativePath = GetSharePointRelativePath(workbookPath)
    If Len(relativePath) = 0 Then Exit Function
    oneDriveRoot = GetBestOneDriveRoot()
    If Len(oneDriveRoot) = 0 Then Exit Function
    localPath = CombinePath(oneDriveRoot, relativePath)
    If FolderExists(localPath) Then ResolveWorkbookLocalPath = localPath: Exit Function
    ResolveWorkbookLocalPath = TryAlternativeOneDriveLayouts(oneDriveRoot, relativePath)
End Function

Private Function GetSharePointRelativePath(ByVal urlPath As String) As String
    Dim p As Long, rel As String
    p = InStr(1, urlPath, "/Documents/", vbTextCompare)
    If p = 0 Then Exit Function
    rel = Mid$(urlPath, p + Len("/Documents/"))
    rel = Replace(rel, "/", "\")
    rel = UrlDecodeBasic(rel)
    GetSharePointRelativePath = rel
End Function

Private Function GetBestOneDriveRoot() As String
    Dim candidates As Collection, c As Variant
    Set candidates = New Collection
    On Error Resume Next
    AddCandidate candidates, Environ$("OneDriveCommercial")
    AddCandidate candidates, Environ$("OneDrive")
    AddCandidate candidates, Environ$("OneDriveConsumer")
    AddCandidate candidates, Environ$("USERPROFILE") & "\OneDrive - Synopsys"
    AddCandidate candidates, Environ$("USERPROFILE") & "\OneDrive - Synopsys, Inc."
    AddCandidate candidates, Environ$("USERPROFILE") & "\OneDrive - Synopsys Inc."
    On Error GoTo 0
    For Each c In candidates
        If FolderExists(CStr(c)) Then GetBestOneDriveRoot = CStr(c): Exit Function
    Next c
End Function

Private Sub AddCandidate(ByRef candidates As Collection, ByVal candidatePath As String)
    If Len(Trim$(candidatePath)) > 0 Then candidates.Add Trim$(candidatePath)
End Sub

Private Function TryAlternativeOneDriveLayouts(ByVal oneDriveRoot As String, ByVal relativePath As String) As String
    Dim testPath As String, rel As String
    rel = relativePath
    testPath = CombinePath(oneDriveRoot, rel)
    If FolderExists(testPath) Then TryAlternativeOneDriveLayouts = testPath: Exit Function
    If LCase$(Left$(rel, Len("Desktop\"))) = LCase$("Desktop\") Then
        testPath = CombinePath(Environ$("USERPROFILE") & "\Desktop", Mid$(rel, Len("Desktop\") + 1))
        If FolderExists(testPath) Then TryAlternativeOneDriveLayouts = testPath: Exit Function
    End If
    testPath = CombinePath(oneDriveRoot, "Documents\" & rel)
    If FolderExists(testPath) Then TryAlternativeOneDriveLayouts = testPath
End Function

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

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    Exit Function
Failed:
    FolderExists = False
End Function

Private Function CombinePath(ByVal parentPath As String, ByVal childName As String) As String
    If Right$(parentPath, 1) = "\" Then CombinePath = parentPath & childName Else CombinePath = parentPath & "\" & childName
End Function

Private Function FindHeaderColumn(ByVal data As Variant, ByVal headerName As String) As Long
    Dim c As Long, text As String
    For c = 1 To UBound(data, 2)
        text = CleanText(data(1, c))
        text = Replace(text, ChrW(&HFEFF), "")
        If StrComp(Trim$(text), headerName, vbTextCompare) = 0 Then FindHeaderColumn = c: Exit Function
    Next c
End Function

Private Function MakeUserKey(ByVal username As String, ByVal email As String, ByVal fullName As String) As String
    If Len(username) > 0 Or Len(email) > 0 Then
        MakeUserKey = LCase$(Trim$(username)) & KEY_SEPARATOR & LCase$(Trim$(email))
    ElseIf Len(fullName) > 0 Then
        MakeUserKey = "FULLNAME" & KEY_SEPARATOR & LCase$(Trim$(fullName))
    End If
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub CreateRunButton(ByVal ws As Worksheet)
    Dim btn As Button
    On Error Resume Next
    ws.Buttons("btnRunPreviousYearAggregation").Delete
    On Error GoTo 0
    Set btn = ws.Buttons.Add(ws.Range("D1").Left, ws.Range("D1").Top, ws.Range("D1:F1").Width, ws.Rows(1).Height + 5)
    With btn
        .Name = "btnRunPreviousYearAggregation"
        .Caption = "Run Previous Year Aggregation"
        .OnAction = "'" & ThisWorkbook.Name & "'!RunPreviousYearAggregation"
        .Font.Size = 10
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 25
End Sub

Private Sub ClearOutputArea(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(ws): lastCol = LastUsedColumn(ws)
    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Clear
End Sub

Private Sub FormatSheet(ByVal ws As Worksheet)

    FormatSheetAtRow ws, HEADER_ROW

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

    With ws.Range( _
        ws.Cells(headerRow, 1), _
        ws.Cells(headerRow, lastCol))

        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    If lastRow >= headerRow Then

        ws.Range( _
            ws.Cells(headerRow, 1), _
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
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If foundCell Is Nothing Then LastUsedRow = 1 Else LastUsedRow = foundCell.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If foundCell Is Nothing Then LastUsedColumn = 1 Else LastUsedColumn = foundCell.Column
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

Private Function GetCellText(ByVal data As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    If colIndex <= 0 Then
        GetCellText = ""
    Else
        GetCellText = CleanText(data(rowIndex, colIndex))
    End If
End Function

Private Function GetCellValue(ByVal data As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If colIndex <= 0 Then
        GetCellValue = Empty
    Else
        GetCellValue = data(rowIndex, colIndex)
    End If
End Function
