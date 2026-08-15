Attribute VB_Name = "CsvAggregation_AMS"
Option Explicit

Private Const CSV_ROOT_FOLDER As String = "csv"
Private Const TARGET_TRACK As String = "WFD-AMS"

Private Const SHEET_USERS As String = "AMS_UserList"
Private Const SHEET_PASSED As String = "AMS_PassedList"

Private Const HEADER_ROW As Long = 3
Private Const KEY_SEPARATOR As String = "|||"

Private gCurrentStep As String

Public Sub SetupAMSAggregation()

    Dim wsUsers As Worksheet
    Dim wsPassed As Worksheet

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Set wsUsers = GetOrCreateSheet(SHEET_USERS)
    Set wsPassed = GetOrCreateSheet(SHEET_PASSED)

    WriteUserHeaders wsUsers
    WritePassedBaseHeaders wsPassed
    CreateRunButton wsUsers

    FormatSheet wsUsers
    FormatSheet wsPassed

    Application.ScreenUpdating = True

    MsgBox _
        "Setup completed." & vbCrLf & vbCrLf & _
        "Use the Run AMS Aggregation button on the first row of the " & _
        SHEET_USERS & " sheet.", _
        vbInformation

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Setup error." & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

End Sub

Public Sub RunAMSAggregation()

    Dim basePath As String
    Dim csvRootPath As String

    Dim wsUsers As Worksheet
    Dim wsPassed As Worksheet

    Dim users As Object
    Dim userOrder As Collection

    Dim passedUsers As Object
    Dim passedOrder As Collection

    Dim examData As Object
    Dim examOrders As Object

    Dim dateFolders As Variant
    Dim csvFiles As Variant

    Dim i As Long
    Dim j As Long

    Dim dateFolderName As String
    Dim trackFolderPath As String
    Dim csvFilePath As String

    Dim processedFileCount As Long
    Dim skippedFileCount As Long

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gCurrentStep = "Resolving workbook folder"

    basePath = ResolveWorkbookLocalPath(ThisWorkbook.Path)

    If Len(basePath) = 0 Then
        MsgBox _
            "Could not resolve the local OneDrive/SharePoint sync folder." & vbCrLf & vbCrLf & _
            "Workbook path:" & vbCrLf & _
            ThisWorkbook.Path & vbCrLf & vbCrLf & _
            "Make sure this folder is synchronized to this PC.", _
            vbExclamation
        GoTo SafeExit
    End If

    csvRootPath = CombinePath(basePath, CSV_ROOT_FOLDER)

    gCurrentStep = "Checking csv folder: " & csvRootPath

    If Not FolderExists(csvRootPath) Then
        MsgBox _
            "CSV folder was not found." & vbCrLf & vbCrLf & _
            "Resolved local workbook folder:" & vbCrLf & _
            basePath & vbCrLf & vbCrLf & _
            "Expected CSV folder:" & vbCrLf & _
            csvRootPath, _
            vbExclamation
        GoTo SafeExit
    End If

    gCurrentStep = "Preparing worksheets"

    Set wsUsers = GetOrCreateSheet(SHEET_USERS)
    Set wsPassed = GetOrCreateSheet(SHEET_PASSED)

    Set users = CreateObject("Scripting.Dictionary")
    users.CompareMode = vbTextCompare
    Set userOrder = New Collection

    Set passedUsers = CreateObject("Scripting.Dictionary")
    passedUsers.CompareMode = vbTextCompare
    Set passedOrder = New Collection

    Set examData = CreateObject("Scripting.Dictionary")
    examData.CompareMode = vbTextCompare

    Set examOrders = CreateObject("Scripting.Dictionary")
    examOrders.CompareMode = vbTextCompare

    gCurrentStep = "Reading date folders"

    dateFolders = GetSortedDateFolders(csvRootPath)

    If IsEmpty(dateFolders) Then
        MsgBox _
            "No valid date folders were found." & vbCrLf & _
            "Example: csv\20260806\WFD-AMS\", _
            vbExclamation
        GoTo SafeExit
    End If

    For i = LBound(dateFolders) To UBound(dateFolders)

        dateFolderName = CStr(dateFolders(i))

        trackFolderPath = CombinePath( _
            CombinePath(csvRootPath, dateFolderName), _
            TARGET_TRACK)

        gCurrentStep = "Checking track folder: " & trackFolderPath

        If FolderExists(trackFolderPath) Then

            gCurrentStep = "Reading CSV files in: " & trackFolderPath
            csvFiles = GetSortedCsvFiles(trackFolderPath)

            If Not IsEmpty(csvFiles) Then

                For j = LBound(csvFiles) To UBound(csvFiles)

                    csvFilePath = CombinePath( _
                        trackFolderPath, _
                        CStr(csvFiles(j)))

                    gCurrentStep = "Processing CSV: " & csvFilePath

                    If ProcessCsvFile( _
                        csvFilePath, _
                        users, _
                        userOrder, _
                        passedUsers, _
                        passedOrder, _
                        examData, _
                        examOrders) Then

                        processedFileCount = processedFileCount + 1
                    Else
                        skippedFileCount = skippedFileCount + 1
                    End If

                Next j

            End If

        End If

    Next i

    gCurrentStep = "Writing UserList"
    OutputUsers wsUsers, users, userOrder

    gCurrentStep = "Writing PassedList"
    OutputPassedUsers _
        wsPassed, _
        passedUsers, _
        passedOrder, _
        examData, _
        examOrders

    gCurrentStep = "Completed"

    MsgBox _
        "CSV aggregation completed." & vbCrLf & vbCrLf & _
        "Resolved local folder:" & vbCrLf & _
        basePath & vbCrLf & vbCrLf & _
        "Processed CSV files: " & processedFileCount & vbCrLf & _
        "Skipped CSV files: " & skippedFileCount & vbCrLf & _
        "Total users: " & users.Count & vbCrLf & _
        "Passed users: " & passedUsers.Count, _
        vbInformation

SafeExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub

ErrorHandler:

    MsgBox _
        "An error occurred." & vbCrLf & vbCrLf & _
        "Step: " & gCurrentStep & vbCrLf & _
        "Error number: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description, _
        vbCritical

    Resume SafeExit

End Sub

Private Function ResolveWorkbookLocalPath(ByVal workbookPath As String) As String

    Dim localPath As String
    Dim relativePath As String
    Dim oneDriveRoot As String

    If Len(workbookPath) = 0 Then
        ResolveWorkbookLocalPath = ""
        Exit Function
    End If

    If InStr(1, workbookPath, "://", vbTextCompare) = 0 Then
        ResolveWorkbookLocalPath = workbookPath
        Exit Function
    End If

    relativePath = GetSharePointRelativePath(workbookPath)

    If Len(relativePath) = 0 Then
        ResolveWorkbookLocalPath = ""
        Exit Function
    End If

    oneDriveRoot = GetBestOneDriveRoot()

    If Len(oneDriveRoot) = 0 Then
        ResolveWorkbookLocalPath = ""
        Exit Function
    End If

    localPath = CombinePath(oneDriveRoot, relativePath)

    If FolderExists(localPath) Then
        ResolveWorkbookLocalPath = localPath
        Exit Function
    End If

    localPath = TryAlternativeOneDriveLayouts(oneDriveRoot, relativePath)

    If Len(localPath) > 0 Then
        ResolveWorkbookLocalPath = localPath
    Else
        ResolveWorkbookLocalPath = ""
    End If

End Function

Private Function GetSharePointRelativePath(ByVal urlPath As String) As String

    Dim p As Long
    Dim rel As String

    p = InStr(1, urlPath, "/Documents/", vbTextCompare)

    If p = 0 Then
        GetSharePointRelativePath = ""
        Exit Function
    End If

    rel = Mid$(urlPath, p + Len("/Documents/"))

    rel = Replace(rel, "/", "\")
    rel = UrlDecodeBasic(rel)

    GetSharePointRelativePath = rel

End Function

Private Function GetBestOneDriveRoot() As String

    Dim candidates As Collection
    Dim c As Variant

    Set candidates = New Collection

    On Error Resume Next

    AddCandidate candidates, Environ$("OneDriveCommercial")
    AddCandidate candidates, Environ$("OneDrive")
    AddCandidate candidates, Environ$("OneDriveConsumer")

    AddCandidate candidates, _
        Environ$("USERPROFILE") & "\OneDrive - Synopsys"

    AddCandidate candidates, _
        Environ$("USERPROFILE") & "\OneDrive - Synopsys, Inc."

    AddCandidate candidates, _
        Environ$("USERPROFILE") & "\OneDrive - Synopsys Inc."

    On Error GoTo 0

    For Each c In candidates
        If FolderExists(CStr(c)) Then
            GetBestOneDriveRoot = CStr(c)
            Exit Function
        End If
    Next c

    GetBestOneDriveRoot = ""

End Function

Private Sub AddCandidate( _
    ByRef candidates As Collection, _
    ByVal candidatePath As String)

    If Len(Trim$(candidatePath)) > 0 Then
        candidates.Add Trim$(candidatePath)
    End If

End Sub

Private Function TryAlternativeOneDriveLayouts( _
    ByVal oneDriveRoot As String, _
    ByVal relativePath As String) As String

    Dim testPath As String
    Dim rel As String

    rel = relativePath

    testPath = CombinePath(oneDriveRoot, rel)
    If FolderExists(testPath) Then
        TryAlternativeOneDriveLayouts = testPath
        Exit Function
    End If

    If LCase$(Left$(rel, Len("Desktop\"))) = LCase$("Desktop\") Then

        testPath = CombinePath( _
            Environ$("USERPROFILE") & "\Desktop", _
            Mid$(rel, Len("Desktop\") + 1))

        If FolderExists(testPath) Then
            TryAlternativeOneDriveLayouts = testPath
            Exit Function
        End If

    End If

    testPath = CombinePath( _
        oneDriveRoot, _
        "Documents\" & rel)

    If FolderExists(testPath) Then
        TryAlternativeOneDriveLayouts = testPath
        Exit Function
    End If

    TryAlternativeOneDriveLayouts = ""

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

Private Function ProcessCsvFile( _
    ByVal csvFilePath As String, _
    ByRef users As Object, _
    ByRef userOrder As Collection, _
    ByRef passedUsers As Object, _
    ByRef passedOrder As Collection, _
    ByRef examData As Object, _
    ByRef examOrders As Object) As Boolean

    Dim wbCsv As Workbook
    Dim wsCsv As Worksheet
    Dim data As Variant

    Dim lastRow As Long
    Dim lastCol As Long

    Dim colUsername As Long
    Dim colEmail As Long
    Dim colFullName As Long
    Dim colCourseTitle As Long
    Dim colEnrollmentDate As Long
    Dim colEnrollmentEndDate As Long
    Dim colStatus As Long
    Dim colFinalScore As Long

    Dim r As Long

    Dim username As String
    Dim email As String
    Dim fullName As String
    Dim courseTitle As String
    Dim courseStatus As String

    Dim enrollmentDate As Variant
    Dim enrollmentEndDate As Variant
    Dim finalScore As Variant

    Dim userKey As String
    Dim examName As String
    Dim examKey As String

    Dim userInfo As Variant
    Dim passInfo As Variant
    Dim examInfo As Variant

    Dim orderCollection As Collection

    On Error GoTo FileError

    Set wbCsv = Workbooks.Open( _
                    Filename:=csvFilePath, _
                    ReadOnly:=True, _
                    Local:=True)

    Set wsCsv = wbCsv.Worksheets(1)

    lastRow = LastUsedRow(wsCsv)
    lastCol = LastUsedColumn(wsCsv)

    If lastRow < 2 Or lastCol < 1 Then GoTo FileError

    data = wsCsv.Range( _
                wsCsv.Cells(1, 1), _
                wsCsv.Cells(lastRow, lastCol)).Value2

    colUsername = FindHeaderColumn(data, "Username")
    colEmail = FindHeaderColumn(data, "Email")
    colFullName = FindHeaderColumn(data, "Full Name")
    colCourseTitle = FindHeaderColumn(data, "Course title")
    colEnrollmentDate = FindHeaderColumn(data, "Enrollment Date")
    colEnrollmentEndDate = FindHeaderColumnAny( _
                                data, _
                                Array( _
                                    "Enrollment End Date", _
                                    "End of validity"))
    colStatus = FindHeaderColumn(data, "Course Enrollment Status")
    colFinalScore = FindHeaderColumn(data, "Final Score")

    If colUsername = 0 _
        Or colEmail = 0 _
        Or colFullName = 0 Then
        GoTo FileError
    End If

    For r = 2 To UBound(data, 1)

        username = CleanText(data(r, colUsername))
        email = CleanText(data(r, colEmail))
        fullName = CleanText(data(r, colFullName))

        enrollmentDate = GetOptionalCellValue(data, r, colEnrollmentDate)
        enrollmentEndDate = GetOptionalCellValue(data, r, colEnrollmentEndDate)

        userKey = MakeUserKey(username, email, fullName)

        If Len(userKey) = 0 Then GoTo NextRow

        If Not users.Exists(userKey) Then

            userInfo = Array( _
                username, _
                email, _
                fullName, _
                enrollmentDate, _
                enrollmentEndDate)

            users.Add userKey, userInfo
            userOrder.Add userKey

        Else

            userInfo = users(userKey)

            userInfo(0) = username
            userInfo(1) = email
            userInfo(2) = fullName
            userInfo(3) = enrollmentDate
            userInfo(4) = enrollmentEndDate

            users(userKey) = userInfo

        End If

        If colCourseTitle > 0 _
            And colStatus > 0 _
            And colFinalScore > 0 Then

            courseTitle = CleanText(data(r, colCourseTitle))
            courseStatus = CleanText(data(r, colStatus))
            finalScore = data(r, colFinalScore)

            If StrComp(courseStatus, "Completed", vbTextCompare) = 0 Then

                examName = GetTargetExamName(courseTitle)

                If Len(examName) > 0 Then

                    If Not passedUsers.Exists(userKey) Then

                        passInfo = Array( _
                            username, _
                            email, _
                            fullName, _
                            enrollmentDate, _
                            enrollmentEndDate)

                        passedUsers.Add userKey, passInfo
                        passedOrder.Add userKey

                        Set orderCollection = New Collection
                        examOrders.Add userKey, orderCollection

                    Else

                        passInfo = passedUsers(userKey)

                        passInfo(0) = username
                        passInfo(1) = email
                        passInfo(2) = fullName
                        passInfo(3) = enrollmentDate
                        passInfo(4) = enrollmentEndDate

                        passedUsers(userKey) = passInfo

                    End If

                    examKey = _
                        userKey & KEY_SEPARATOR & _
                        LCase$(examName)

                    If Not examData.Exists(examKey) Then

                        examInfo = Array( _
                            courseTitle, _
                            "Completed", _
                            finalScore)

                        examData.Add examKey, examInfo

                        Set orderCollection = examOrders(userKey)
                        orderCollection.Add examName

                    Else

                        examInfo = examData(examKey)

                        examInfo(0) = courseTitle
                        examInfo(1) = "Completed"
                        examInfo(2) = finalScore

                        examData(examKey) = examInfo

                    End If

                End If

            End If

        End If

NextRow:
    Next r

    wbCsv.Close SaveChanges:=False
    ProcessCsvFile = True
    Exit Function

FileError:

    On Error Resume Next
    If Not wbCsv Is Nothing Then
        wbCsv.Close SaveChanges:=False
    End If
    On Error GoTo 0

    ProcessCsvFile = False

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

Private Function CombinePath( _
    ByVal parentPath As String, _
    ByVal childName As String) As String

    If Right$(parentPath, 1) = "\" Then
        CombinePath = parentPath & childName
    Else
        CombinePath = parentPath & "\" & childName
    End If

End Function

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

Private Function GetSortedCsvFiles( _
    ByVal folderPath As String) As Variant

    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object

    Dim arr() As String
    Dim count As Long
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(folderPath)

    For Each fileObj In folderObj.Files

        ext = LCase$(fso.GetExtensionName(CStr(fileObj.Name)))

        If ext = "csv" Then
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

Private Sub OutputUsers( _
    ByVal ws As Worksheet, _
    ByVal users As Object, _
    ByVal userOrder As Collection)

    Dim outputData() As Variant
    Dim i As Long
    Dim key As String
    Dim info As Variant

    ClearOutputArea ws
    WriteUserHeaders ws
    CreateRunButton ws

    If users.Count = 0 Then
        FormatSheet ws
        Exit Sub
    End If

    ReDim outputData(1 To users.Count, 1 To 6)

    For i = 1 To userOrder.Count

        key = CStr(userOrder(i))
        info = users(key)

        outputData(i, 1) = i
        outputData(i, 2) = info(0)
        outputData(i, 3) = info(1)
        outputData(i, 4) = info(2)
        outputData(i, 5) = info(3)
        outputData(i, 6) = info(4)

    Next i

    ws.Cells(HEADER_ROW + 1, 1) _
        .Resize(users.Count, 6).Value = outputData

    ws.Columns(5).NumberFormat = "yyyy/mm/dd"
    ws.Columns(6).NumberFormat = "yyyy/mm/dd"

    FormatSheet ws

End Sub

Private Sub OutputPassedUsers( _
    ByVal ws As Worksheet, _
    ByVal passedUsers As Object, _
    ByVal passedOrder As Collection, _
    ByVal examData As Object, _
    ByVal examOrders As Object)

    Dim i As Long
    Dim j As Long

    Dim key As String
    Dim examKey As String
    Dim examName As String

    Dim info As Variant
    Dim examInfo As Variant

    Dim orderCollection As Collection

    Dim maxExamCount As Long
    Dim totalColumns As Long
    Dim outputData() As Variant
    Dim c As Long

    ClearOutputArea ws

    maxExamCount = 0

    For i = 1 To passedOrder.Count

        key = CStr(passedOrder(i))
        Set orderCollection = examOrders(key)

        If orderCollection.Count > maxExamCount Then
            maxExamCount = orderCollection.Count
        End If

    Next i

    WritePassedHeaders ws, maxExamCount

    If passedUsers.Count = 0 Then
        FormatSheet ws
        Exit Sub
    End If

    totalColumns = 6 + (maxExamCount * 3)

    ReDim outputData( _
        1 To passedUsers.Count, _
        1 To totalColumns)

    For i = 1 To passedOrder.Count

        key = CStr(passedOrder(i))
        info = passedUsers(key)

        outputData(i, 1) = i
        outputData(i, 2) = info(0)
        outputData(i, 3) = info(1)
        outputData(i, 4) = info(2)
        outputData(i, 5) = info(3)
        outputData(i, 6) = info(4)

        Set orderCollection = examOrders(key)

        c = 7

        For j = 1 To orderCollection.Count

            examName = CStr(orderCollection(j))

            examKey = _
                key & KEY_SEPARATOR & _
                LCase$(examName)

            examInfo = examData(examKey)

            outputData(i, c) = examInfo(0)
            outputData(i, c + 1) = examInfo(1)
            outputData(i, c + 2) = examInfo(2)

            c = c + 3

        Next j

    Next i

    ws.Cells(HEADER_ROW + 1, 1) _
        .Resize(passedUsers.Count, totalColumns).Value = outputData

    ws.Columns(5).NumberFormat = "yyyy/mm/dd"
    ws.Columns(6).NumberFormat = "yyyy/mm/dd"

    FormatSheet ws

End Sub

Private Function GetTargetExamName( _
    ByVal courseTitle As String) As String

    Dim title As String

    title = Trim$(courseTitle)

    If InStr(1, title, _
        "Purple Certification:", _
        vbTextCompare) = 1 Then

        title = Trim$(Mid$( _
            title, _
            Len("Purple Certification:") + 1))

    End If

    Select Case LCase$(title)

        Case LCase$("Custom Compiler: Basic Layout Design Exam")
            GetTargetExamName = _
                "Custom Compiler: Basic Layout Design Exam"

        Case LCase$("Custom Compiler: Introduction to Platform Exam")
            GetTargetExamName = _
                "Custom Compiler: Introduction to Platform Exam"

        Case LCase$("Custom Compiler: Schematic Entry Exam")
            GetTargetExamName = _
                "Custom Compiler: Schematic Entry Exam"

        Case LCase$("PrimeWave Design Environment Exam")
            GetTargetExamName = "PrimeWave Design Environment Exam"

        Case Else
            GetTargetExamName = ""

    End Select

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
            "FULLNAME" & _
            KEY_SEPARATOR & _
            LCase$(Trim$(fullName))

    Else

        MakeUserKey = ""

    End If

End Function

Private Function FindHeaderColumn( _
    ByVal data As Variant, _
    ByVal headerName As String) As Long

    Dim c As Long
    Dim text As String

    For c = 1 To UBound(data, 2)

        text = CleanText(data(1, c))
        text = Replace(text, ChrW(&HFEFF), "")

        If StrComp( _
            Trim$(text), _
            headerName, _
            vbTextCompare) = 0 Then

            FindHeaderColumn = c
            Exit Function

        End If

    Next c

    FindHeaderColumn = 0

End Function

Private Function FindHeaderColumnAny( _
    ByVal data As Variant, _
    ByVal headerNames As Variant) As Long

    Dim i As Long
    Dim col As Long

    For i = LBound(headerNames) To UBound(headerNames)

        col = FindHeaderColumn(data, CStr(headerNames(i)))

        If col > 0 Then
            FindHeaderColumnAny = col
            Exit Function
        End If

    Next i

    FindHeaderColumnAny = 0

End Function

Private Sub SortStringArray(ByRef arr As Variant)

    Dim i As Long
    Dim j As Long
    Dim temp As String

    For i = LBound(arr) To UBound(arr) - 1

        For j = i + 1 To UBound(arr)

            If StrComp( _
                CStr(arr(i)), _
                CStr(arr(j)), _
                vbTextCompare) > 0 Then

                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp

            End If

        Next j

    Next i

End Sub

Private Function IsDateFolderName( _
    ByVal folderName As String) As Boolean

    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long

    On Error GoTo InvalidDate

    If Len(folderName) <> 8 Then Exit Function
    If Not IsNumeric(folderName) Then Exit Function

    yyyy = CLng(Left$(folderName, 4))
    mm = CLng(Mid$(folderName, 5, 2))
    dd = CLng(Right$(folderName, 2))

    If Format$(DateSerial(yyyy, mm, dd), "yyyymmdd") <> folderName Then
        Exit Function
    End If

    IsDateFolderName = True
    Exit Function

InvalidDate:
    IsDateFolderName = False

End Function

Private Sub WriteUserHeaders(ByVal ws As Worksheet)

    ws.Cells(HEADER_ROW, 1).Value = "No."
    ws.Cells(HEADER_ROW, 2).Value = "Username"
    ws.Cells(HEADER_ROW, 3).Value = "Email"
    ws.Cells(HEADER_ROW, 4).Value = "Full Name"
    ws.Cells(HEADER_ROW, 5).Value = "Enrollment Date"
    ws.Cells(HEADER_ROW, 6).Value = "Enrollment End Date"

End Sub

Private Sub WritePassedBaseHeaders(ByVal ws As Worksheet)
    WritePassedHeaders ws, 0
End Sub

Private Sub WritePassedHeaders( _
    ByVal ws As Worksheet, _
    ByVal examCount As Long)

    Dim i As Long
    Dim c As Long

    ws.Cells(HEADER_ROW, 1).Value = "No."
    ws.Cells(HEADER_ROW, 2).Value = "Username"
    ws.Cells(HEADER_ROW, 3).Value = "Email"
    ws.Cells(HEADER_ROW, 4).Value = "Full Name"
    ws.Cells(HEADER_ROW, 5).Value = "Enrollment Date"
    ws.Cells(HEADER_ROW, 6).Value = "Enrollment End Date"

    c = 7

    For i = 1 To examCount

        ws.Cells(HEADER_ROW, c).Value = "Course title" & i
        ws.Cells(HEADER_ROW, c + 1).Value = _
            "Course Enrollment Status" & i
        ws.Cells(HEADER_ROW, c + 2).Value = "Final Score" & i

        c = c + 3

    Next i

End Sub

Private Sub ClearOutputArea(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    If lastCol < 1 Then lastCol = 1

    ws.Range( _
        ws.Cells(HEADER_ROW, 1), _
        ws.Cells(lastRow, lastCol)).Clear

End Sub

Private Function GetOrCreateSheet( _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = sheetName

    End If

    Set GetOrCreateSheet = ws

End Function

Private Sub CreateRunButton(ByVal ws As Worksheet)

    Dim btn As Button
    Dim leftPos As Double
    Dim topPos As Double
    Dim btnWidth As Double
    Dim btnHeight As Double

    On Error Resume Next
    ws.Buttons("btnRunAMSAggregation").Delete
    On Error GoTo 0

    leftPos = ws.Range("A1").Left
    topPos = ws.Range("A1").Top
    btnWidth = ws.Range("A1:C1").Width
    btnHeight = ws.Rows(1).Height + 5

    Set btn = ws.Buttons.Add( _
                leftPos, _
                topPos, _
                btnWidth, _
                btnHeight)

    With btn
        .Name = "btnRunAMSAggregation"
        .Caption = "Run AMS Aggregation"

        .OnAction = _
            "'" & ThisWorkbook.Name & _
            "'!RunAMSAggregation"

        .Font.Size = 11
        .Font.Bold = True
    End With

    ws.Rows(1).RowHeight = 25

End Sub

Private Sub FormatSheet(ByVal ws As Worksheet)

    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long

    lastCol = LastUsedColumn(ws)
    lastRow = LastUsedRow(ws)

    If lastCol < 1 Then Exit Sub

    With ws.Range( _
        ws.Cells(HEADER_ROW, 1), _
        ws.Cells(HEADER_ROW, lastCol))

        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    If lastRow >= HEADER_ROW Then
        ws.Range( _
            ws.Cells(HEADER_ROW, 1), _
            ws.Cells(lastRow, lastCol)).AutoFilter
    End If

    ws.Columns.AutoFit

    For c = 1 To lastCol
        If ws.Columns(c).ColumnWidth > 45 Then
            ws.Columns(c).ColumnWidth = 45
        End If
    Next c

End Sub

Private Function LastUsedRow( _
    ByVal ws As Worksheet) As Long

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

Private Function LastUsedColumn( _
    ByVal ws As Worksheet) As Long

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

Private Function CleanText( _
    ByVal value As Variant) As String

    If IsError(value) Then
        CleanText = ""
    ElseIf IsNull(value) Or IsEmpty(value) Then
        CleanText = ""
    Else
        CleanText = Trim$(CStr(value))
    End If

End Function

Private Function GetOptionalCellValue( _
    ByVal data As Variant, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long) As Variant

    If colIndex <= 0 Then
        GetOptionalCellValue = ""
    Else
        GetOptionalCellValue = data(rowIndex, colIndex)
    End If

End Function

