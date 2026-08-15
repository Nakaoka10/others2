Attribute VB_Name = "CsvAggregation_DFT"
Option Explicit

Private Const SHEET_USER_LIST As String = "DFT_UserList"
Private Const SHEET_PASSED_LIST As String = "DFT_PassedList"
Private Const CSV_FOLDER_NAME As String = "csv"
Private Const TARGET_SUBFOLDER_NAME As String = "WFD-Design_for_Test"
Private Const STATUS_COMPLETED As String = "Completed"

Private Const HEADER_USERNAME As String = "Username"
Private Const HEADER_EMAIL As String = "Email"
Private Const HEADER_FULL_NAME As String = "Full Name"
Private Const HEADER_COURSE_TITLE As String = "Course title"
Private Const HEADER_STATUS As String = "Course Enrollment Status"
Private Const HEADER_FINAL_SCORE As String = "Final Score"
Private Const HEADER_SCAN_LIMIT As Long = 50

Public Sub SetupDFTAggregation()
    Dim wsUsers As Worksheet
    Dim wsPassed As Worksheet

    Set wsUsers = EnsureSheet(SHEET_USER_LIST)
    Set wsPassed = EnsureSheet(SHEET_PASSED_LIST)

    wsUsers.Cells.Clear
    wsPassed.Cells.Clear

    WriteBaseHeaders wsUsers
    WriteBaseHeaders wsPassed
    WritePassedGroupHeaders wsPassed, 1

    AddRunButton wsUsers

    wsUsers.Columns.AutoFit
    wsPassed.Columns.AutoFit

    MsgBox "DFT aggregation sheets were prepared.", vbInformation
End Sub

Public Sub RunDFTAggregation()
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim workbookFolder As String
    Dim csvRootPath As String
    Dim csvRoot As Object
    Dim dateFolders As Collection
    Dim allUsers As Object
    Dim passedUsers As Object
    Dim targetTitles As Object
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    workbookFolder = ResolveLocalWorkbookFolder(fso)
    csvRootPath = fso.BuildPath(workbookFolder, CSV_FOLDER_NAME)

    If Not fso.FolderExists(csvRootPath) Then
        Err.Raise vbObjectError + 101, , "CSV root folder was not found: " & csvRootPath
    End If

    Set csvRoot = fso.GetFolder(csvRootPath)
    Set dateFolders = GetSortedDateFolders(csvRoot)
    Set allUsers = CreateTextDictionary()
    Set passedUsers = CreateTextDictionary()
    Set targetTitles = BuildTargetTitleDictionary()

    For i = 1 To dateFolders.Count
        ProcessDateFolder fso, dateFolders(i), allUsers, passedUsers, targetTitles
    Next i

    OutputResults allUsers, passedUsers

    MsgBox "DFT aggregation completed." & vbCrLf & _
           "Resolved local folder: " & workbookFolder & vbCrLf & _
           "Date folders processed: " & CStr(dateFolders.Count), vbInformation
    Exit Sub

ErrHandler:
    MsgBox "DFT aggregation failed." & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub ProcessDateFolder(ByVal fso As Object, ByVal dateFolder As Object, ByVal allUsers As Object, ByVal passedUsers As Object, ByVal targetTitles As Object)
    Dim targetFolderPath As String
    Dim targetFolder As Object
    Dim fileObj As Object

    targetFolderPath = fso.BuildPath(dateFolder.Path, TARGET_SUBFOLDER_NAME)
    If Not fso.FolderExists(targetFolderPath) Then Exit Sub

    Set targetFolder = fso.GetFolder(targetFolderPath)

    For Each fileObj In targetFolder.Files
        If LCase$(fso.GetExtensionName(fileObj.Name)) = "csv" Then
            ProcessCsvFile fso, fileObj.Path, dateFolder.Name, allUsers, passedUsers, targetTitles
        End If
    Next fileObj
End Sub

Private Sub ProcessCsvFile(ByVal fso As Object, ByVal filePath As String, ByVal sourceDate As String, ByVal allUsers As Object, ByVal passedUsers As Object, ByVal targetTitles As Object)
    Dim ts As Object
    Dim lineText As String
    Dim rowNo As Long
    Dim headers As Variant
    Dim headerMap As Object
    Dim fields As Variant
    Dim delimiter As String
    Dim userKey As String
    Dim title As String
    Dim statusText As String
    Dim finalScore As String
    Dim rec As Object
    Dim bestHeaders As Variant
    Dim bestHeaderMap As Object
    Dim currentHeaderMap As Object
    Dim bestDelimiter As String
    Dim bestScore As Long
    Dim bestRowNo As Long
    Dim currentScore As Long

    Set ts = fso.OpenTextFile(filePath, 1, False, -2)
    rowNo = 0
    Set headerMap = Nothing
    bestScore = -1
    bestRowNo = 0

    Do While Not ts.AtEndOfStream
        lineText = CleanInputLine(ts.ReadLine)
        rowNo = rowNo + 1

        If headerMap Is Nothing Then
            If Len(Trim$(lineText)) = 0 Then GoTo ContinueLoop

            delimiter = DetectDelimiter(lineText)
            headers = ParseDelimitedLine(lineText, delimiter)
            Set currentHeaderMap = BuildHeaderMap(headers)
            currentScore = CountRequiredHeaders(currentHeaderMap)

            If currentScore > bestScore Then
                bestHeaders = headers
                Set bestHeaderMap = currentHeaderMap
                bestDelimiter = delimiter
                bestScore = currentScore
                bestRowNo = rowNo
            End If

            If currentScore = RequiredHeaderCount() Then
                Set headerMap = currentHeaderMap
                GoTo ContinueLoop
            End If

            If rowNo >= HEADER_SCAN_LIMIT Then
                ValidateHeaders bestHeaderMap, filePath, bestHeaders, bestDelimiter, bestRowNo
            End If

            GoTo ContinueLoop
        End If

        If Len(Trim$(lineText)) = 0 Then GoTo ContinueLoop

        fields = ParseDelimitedLine(lineText, delimiter)
        userKey = BuildUserKey(GetField(fields, headerMap, HEADER_USERNAME), _
                               GetField(fields, headerMap, HEADER_EMAIL), _
                               GetField(fields, headerMap, HEADER_FULL_NAME))

        If Len(userKey) = 0 Then GoTo ContinueLoop

        Set rec = GetOrCreateUser(allUsers, userKey)
        UpdateBasicInfo rec, fields, headerMap, sourceDate

        title = GetField(fields, headerMap, HEADER_COURSE_TITLE)
        statusText = GetField(fields, headerMap, HEADER_STATUS)
        finalScore = GetField(fields, headerMap, HEADER_FINAL_SCORE)

        If targetTitles.Exists(title) And StrComp(statusText, STATUS_COMPLETED, vbTextCompare) = 0 Then
            Set rec = GetOrCreateUser(passedUsers, userKey)
            UpdateBasicInfo rec, fields, headerMap, sourceDate
            UpdateExam rec, title, statusText, finalScore, sourceDate
        End If

ContinueLoop:
    Loop

    ts.Close

    If headerMap Is Nothing Then
        ValidateHeaders bestHeaderMap, filePath, bestHeaders, bestDelimiter, bestRowNo
    End If
End Sub

Private Sub OutputResults(ByVal allUsers As Object, ByVal passedUsers As Object)
    Dim wsUsers As Worksheet
    Dim wsPassed As Worksheet
    Dim maxExamCount As Long

    Set wsUsers = EnsureSheet(SHEET_USER_LIST)
    Set wsPassed = EnsureSheet(SHEET_PASSED_LIST)

    ClearOutputArea wsUsers
    ClearOutputArea wsPassed

    WriteBaseHeaders wsUsers
    WriteBaseHeaders wsPassed

    maxExamCount = GetMaxExamCount(passedUsers)
    If maxExamCount < 1 Then maxExamCount = 1
    WritePassedGroupHeaders wsPassed, maxExamCount

    WriteUserRows wsUsers, allUsers, False
    WriteUserRows wsPassed, passedUsers, True

    AddRunButton wsUsers

    wsUsers.Columns.AutoFit
    wsPassed.Columns.AutoFit
End Sub

Private Sub WriteUserRows(ByVal ws As Worksheet, ByVal users As Object, ByVal includeExams As Boolean)
    Dim keys As Variant
    Dim i As Long
    Dim rowNo As Long
    Dim rec As Object
    Dim userKey As Variant

    If users.Count = 0 Then Exit Sub

    keys = SortedDictionaryKeys(users)
    rowNo = 4

    For i = LBound(keys) To UBound(keys)
        userKey = keys(i)
        Set rec = users(userKey)

        ws.Cells(rowNo, 1).Value = rec("Username")
        ws.Cells(rowNo, 2).Value = rec("Email")
        ws.Cells(rowNo, 3).Value = rec("Full Name")
        ws.Cells(rowNo, 4).Value = rec("Source Date")

        If includeExams Then WriteExamCells ws, rowNo, rec

        rowNo = rowNo + 1
    Next i
End Sub

Private Sub WriteExamCells(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal rec As Object)
    Dim order As Collection
    Dim exams As Object
    Dim i As Long
    Dim baseCol As Long
    Dim title As String
    Dim exam As Object

    Set order = rec("ExamOrder")
    Set exams = rec("Exams")

    For i = 1 To order.Count
        title = CStr(order(i))
        Set exam = exams(title)
        baseCol = 5 + ((i - 1) * 3)
        ws.Cells(rowNo, baseCol).Value = exam("Course title")
        ws.Cells(rowNo, baseCol + 1).Value = exam("Course Enrollment Status")
        ws.Cells(rowNo, baseCol + 2).Value = exam("Final Score")
    Next i
End Sub

Private Sub ClearOutputArea(ByVal ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Sub WriteBaseHeaders(ByVal ws As Worksheet)
    ws.Cells(3, 1).Value = HEADER_USERNAME
    ws.Cells(3, 2).Value = HEADER_EMAIL
    ws.Cells(3, 3).Value = HEADER_FULL_NAME
    ws.Cells(3, 4).Value = "Source Date"
End Sub

Private Sub WritePassedGroupHeaders(ByVal ws As Worksheet, ByVal groupCount As Long)
    Dim i As Long
    Dim baseCol As Long

    For i = 1 To groupCount
        baseCol = 5 + ((i - 1) * 3)
        ws.Cells(3, baseCol).Value = HEADER_COURSE_TITLE & CStr(i)
        ws.Cells(3, baseCol + 1).Value = HEADER_STATUS & CStr(i)
        ws.Cells(3, baseCol + 2).Value = HEADER_FINAL_SCORE & CStr(i)
    Next i
End Sub

Private Sub AddRunButton(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim btn As Button

    For Each shp In ws.Shapes
        If shp.Name = "btnRunDFTAggregation" Then shp.Delete
    Next shp

    Set btn = ws.Buttons.Add(ws.Cells(1, 1).Left, ws.Cells(1, 1).Top, 170, 28)
    btn.Name = "btnRunDFTAggregation"
    btn.Caption = "Run DFT Aggregation"
    btn.OnAction = "RunDFTAggregation"
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Function GetOrCreateUser(ByVal users As Object, ByVal userKey As String) As Object
    Dim rec As Object

    If users.Exists(userKey) Then
        Set GetOrCreateUser = users(userKey)
        Exit Function
    End If

    Set rec = CreateTextDictionary()
    rec("Username") = vbNullString
    rec("Email") = vbNullString
    rec("Full Name") = vbNullString
    rec("Source Date") = vbNullString
    Set rec("Exams") = CreateTextDictionary()
    Set rec("ExamOrder") = New Collection

    users.Add userKey, rec
    Set GetOrCreateUser = rec
End Function

Private Sub UpdateBasicInfo(ByVal rec As Object, ByVal fields As Variant, ByVal headerMap As Object, ByVal sourceDate As String)
    If CStr(rec("Source Date")) <= sourceDate Then
        rec("Username") = GetField(fields, headerMap, HEADER_USERNAME)
        rec("Email") = GetField(fields, headerMap, HEADER_EMAIL)
        rec("Full Name") = GetField(fields, headerMap, HEADER_FULL_NAME)
        rec("Source Date") = sourceDate
    End If
End Sub

Private Sub UpdateExam(ByVal rec As Object, ByVal title As String, ByVal statusText As String, ByVal finalScore As String, ByVal sourceDate As String)
    Dim exams As Object
    Dim exam As Object
    Dim order As Collection

    Set exams = rec("Exams")
    Set order = rec("ExamOrder")

    If exams.Exists(title) Then
        Set exam = exams(title)
        If CStr(exam("Source Date")) <= sourceDate Then
            exam("Course title") = title
            exam("Course Enrollment Status") = statusText
            exam("Final Score") = finalScore
            exam("Source Date") = sourceDate
        End If
    Else
        Set exam = CreateTextDictionary()
        exam("Course title") = title
        exam("Course Enrollment Status") = statusText
        exam("Final Score") = finalScore
        exam("Source Date") = sourceDate
        exams.Add title, exam
        order.Add title
    End If
End Sub

Private Function BuildTargetTitleDictionary() As Object
    Set BuildTargetTitleDictionary = CreateBinaryDictionary()
    BuildTargetTitleDictionary.Add "TestMAX ATPG Exam", True
    BuildTargetTitleDictionary.Add "Fusion Compiler: DFT Synthesis Exam", True
    BuildTargetTitleDictionary.Add "TestMAX Advisor Exam", True
    BuildTargetTitleDictionary.Add "TestMAX DFT Exam", True
End Function

Private Function BuildUserKey(ByVal username As String, ByVal email As String, ByVal fullName As String) As String
    username = Trim$(username)
    email = Trim$(email)
    fullName = Trim$(fullName)

    If Len(username) > 0 Or Len(email) > 0 Then
        BuildUserKey = LCase$(username) & "|" & LCase$(email)
    Else
        BuildUserKey = LCase$(fullName)
    End If
End Function

Private Function GetMaxExamCount(ByVal users As Object) As Long
    Dim key As Variant
    Dim rec As Object
    Dim countValue As Long

    GetMaxExamCount = 0
    For Each key In users.Keys
        Set rec = users(key)
        countValue = rec("ExamOrder").Count
        If countValue > GetMaxExamCount Then GetMaxExamCount = countValue
    Next key
End Function

Private Function GetSortedDateFolders(ByVal csvRoot As Object) As Collection
    Dim rawFolders As Collection
    Dim sortedFolders As Collection
    Dim arr() As Object
    Dim subFolder As Object
    Dim i As Long
    Dim j As Long
    Dim temp As Object

    Set rawFolders = New Collection

    For Each subFolder In csvRoot.SubFolders
        If IsDateFolderName(subFolder.Name) Then rawFolders.Add subFolder
    Next subFolder

    Set sortedFolders = New Collection
    If rawFolders.Count = 0 Then
        Set GetSortedDateFolders = sortedFolders
        Exit Function
    End If

    ReDim arr(1 To rawFolders.Count)
    For i = 1 To rawFolders.Count
        Set arr(i) = rawFolders(i)
    Next i

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(i).Name) > CStr(arr(j).Name) Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i

    For i = LBound(arr) To UBound(arr)
        sortedFolders.Add arr(i)
    Next i

    Set GetSortedDateFolders = sortedFolders
End Function

Private Function IsDateFolderName(ByVal folderName As String) As Boolean
    Dim i As Long
    If Len(folderName) <> 8 Then Exit Function
    For i = 1 To 8
        If Mid$(folderName, i, 1) < "0" Or Mid$(folderName, i, 1) > "9" Then Exit Function
    Next i
    IsDateFolderName = True
End Function

Private Function SortedDictionaryKeys(ByVal dict As Object) As Variant
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim temp As Variant

    keys = dict.Keys
    If dict.Count <= 1 Then
        SortedDictionaryKeys = keys
        Exit Function
    End If

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CStr(keys(i)) > CStr(keys(j)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i

    SortedDictionaryKeys = keys
End Function

Private Function BuildHeaderMap(ByVal headers As Variant) As Object
    Dim map As Object
    Dim i As Long
    Dim headerText As String

    Set map = CreateTextDictionary()

    For i = LBound(headers) To UBound(headers)
        headerText = NormalizeHeaderText(CStr(headers(i)))
        If Len(headerText) > 0 Then
            If Not map.Exists(headerText) Then map.Add headerText, i
        End If
    Next i

    Set BuildHeaderMap = map
End Function

Private Function CleanInputLine(ByVal lineText As String) As String
    CleanInputLine = Replace(lineText, Chr$(0), vbNullString)
End Function

Private Function NormalizeHeaderText(ByVal headerText As String) As String
    headerText = Trim$(headerText)
    headerText = Replace(headerText, Chr$(0), vbNullString)
    headerText = Replace(headerText, Chr$(255) & Chr$(254), vbNullString)
    headerText = Replace(headerText, Chr$(254) & Chr$(255), vbNullString)

    If Len(headerText) > 0 Then
        If AscW(Left$(headerText, 1)) = -257 Then headerText = Mid$(headerText, 2)
    End If

    If Len(headerText) >= 2 Then
        If Left$(headerText, 1) = """" And Right$(headerText, 1) = """" Then
            headerText = Mid$(headerText, 2, Len(headerText) - 2)
        End If
    End If

    NormalizeHeaderText = Trim$(headerText)
End Function

Private Function RequiredHeaderCount() As Long
    RequiredHeaderCount = 6
End Function

Private Function CountRequiredHeaders(ByVal headerMap As Object) As Long
    If headerMap Is Nothing Then Exit Function
    If headerMap.Exists(HEADER_USERNAME) Then CountRequiredHeaders = CountRequiredHeaders + 1
    If headerMap.Exists(HEADER_EMAIL) Then CountRequiredHeaders = CountRequiredHeaders + 1
    If headerMap.Exists(HEADER_FULL_NAME) Then CountRequiredHeaders = CountRequiredHeaders + 1
    If headerMap.Exists(HEADER_COURSE_TITLE) Then CountRequiredHeaders = CountRequiredHeaders + 1
    If headerMap.Exists(HEADER_STATUS) Then CountRequiredHeaders = CountRequiredHeaders + 1
    If headerMap.Exists(HEADER_FINAL_SCORE) Then CountRequiredHeaders = CountRequiredHeaders + 1
End Function

Private Sub ValidateHeaders(ByVal headerMap As Object, ByVal filePath As String, ByVal headers As Variant, ByVal delimiter As String, ByVal rowNo As Long)
    Dim missing As String
    Dim detail As String

    AddMissingHeader missing, headerMap, HEADER_USERNAME
    AddMissingHeader missing, headerMap, HEADER_EMAIL
    AddMissingHeader missing, headerMap, HEADER_FULL_NAME
    AddMissingHeader missing, headerMap, HEADER_COURSE_TITLE
    AddMissingHeader missing, headerMap, HEADER_STATUS
    AddMissingHeader missing, headerMap, HEADER_FINAL_SCORE

    If Len(missing) > 0 Then
        detail = "Required CSV header(s) missing in " & filePath & ": " & Mid$(missing, 3) & vbCrLf & _
                 "Best header candidate row: " & CStr(rowNo) & vbCrLf & _
                 "Detected delimiter: " & DelimiterName(delimiter) & vbCrLf & _
                 "Detected header fields: " & HeaderArrayPreview(headers)
        Err.Raise vbObjectError + 102, , detail
    End If
End Sub

Private Function DelimiterName(ByVal delimiter As String) As String
    If delimiter = vbTab Then
        DelimiterName = "tab"
    ElseIf delimiter = "," Then
        DelimiterName = "comma"
    ElseIf delimiter = ";" Then
        DelimiterName = "semicolon"
    ElseIf delimiter = "|" Then
        DelimiterName = "pipe"
    Else
        DelimiterName = "unknown"
    End If
End Function

Private Function HeaderArrayPreview(ByVal headers As Variant) As String
    Dim i As Long
    Dim itemText As String

    On Error GoTo NoHeaders

    For i = LBound(headers) To UBound(headers)
        itemText = NormalizeHeaderText(CStr(headers(i)))
        If Len(itemText) > 80 Then itemText = Left$(itemText, 80) & "..."
        HeaderArrayPreview = HeaderArrayPreview & " [" & CStr(i + 1) & "] " & itemText
    Next i

    If Len(HeaderArrayPreview) = 0 Then HeaderArrayPreview = "(blank)"
    Exit Function

NoHeaders:
    HeaderArrayPreview = "(none)"
End Function

Private Sub AddMissingHeader(ByRef missing As String, ByVal headerMap As Object, ByVal headerName As String)
    If headerMap Is Nothing Then
        missing = missing & ", " & headerName
    ElseIf Not headerMap.Exists(headerName) Then
        missing = missing & ", " & headerName
    End If
End Sub

Private Function GetField(ByVal fields As Variant, ByVal headerMap As Object, ByVal headerName As String) As String
    Dim idx As Long

    If headerMap Is Nothing Then Exit Function
    If Not headerMap.Exists(headerName) Then Exit Function

    idx = CLng(headerMap(headerName))
    If idx < LBound(fields) Or idx > UBound(fields) Then Exit Function

    GetField = Trim$(CStr(fields(idx)))
End Function

Private Function ParseCsvLine(ByVal lineText As String) As Variant
    ParseCsvLine = ParseDelimitedLine(lineText, DetectDelimiter(lineText))
End Function

Private Function ParseDelimitedLine(ByVal lineText As String, ByVal delimiter As String) As Variant
    Dim values As Collection
    Dim currentValue As String
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean

    Set values = New Collection
    currentValue = vbNullString
    inQuotes = False

    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                currentValue = currentValue & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = delimiter And Not inQuotes Then
            values.Add currentValue
            currentValue = vbNullString
        Else
            currentValue = currentValue & ch
        End If
    Next i

    values.Add currentValue
    ParseDelimitedLine = CollectionToArray(values)
End Function

Private Function DetectDelimiter(ByVal lineText As String) As String
    Dim commaCount As Long
    Dim tabCount As Long
    Dim semicolonCount As Long
    Dim pipeCount As Long

    commaCount = CountDelimiterOutsideQuotes(lineText, ",")
    tabCount = CountDelimiterOutsideQuotes(lineText, vbTab)
    semicolonCount = CountDelimiterOutsideQuotes(lineText, ";")
    pipeCount = CountDelimiterOutsideQuotes(lineText, "|")

    If tabCount > commaCount And tabCount >= semicolonCount And tabCount >= pipeCount Then
        DetectDelimiter = vbTab
    ElseIf semicolonCount > commaCount And semicolonCount > tabCount And semicolonCount >= pipeCount Then
        DetectDelimiter = ";"
    ElseIf pipeCount > commaCount And pipeCount > tabCount And pipeCount > semicolonCount Then
        DetectDelimiter = "|"
    Else
        DetectDelimiter = ","
    End If
End Function

Private Function CountDelimiterOutsideQuotes(ByVal lineText As String, ByVal delimiter As String) As Long
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean

    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = delimiter And Not inQuotes Then
            CountDelimiterOutsideQuotes = CountDelimiterOutsideQuotes + 1
        End If
    Next i
End Function

Private Function CollectionToArray(ByVal values As Collection) As Variant
    Dim arr() As String
    Dim i As Long

    ReDim arr(0 To values.Count - 1)
    For i = 1 To values.Count
        arr(i - 1) = CStr(values(i))
    Next i

    CollectionToArray = arr
End Function

Private Function CreateTextDictionary() As Object
    Set CreateTextDictionary = CreateObject("Scripting.Dictionary")
    CreateTextDictionary.CompareMode = vbTextCompare
End Function

Private Function CreateBinaryDictionary() As Object
    Set CreateBinaryDictionary = CreateObject("Scripting.Dictionary")
    CreateBinaryDictionary.CompareMode = vbBinaryCompare
End Function

Private Function ResolveLocalWorkbookFolder(ByVal fso As Object) As String
    Dim workbookPath As String

    workbookPath = ThisWorkbook.Path
    If Len(workbookPath) = 0 Then
        Err.Raise vbObjectError + 103, , "Please save this workbook before running aggregation."
    End If

    If IsHttpPath(workbookPath) Then
        ResolveLocalWorkbookFolder = ResolveSharePointFolderToLocal(fso, workbookPath)
    Else
        ResolveLocalWorkbookFolder = workbookPath
    End If

    If Len(ResolveLocalWorkbookFolder) = 0 Or Not fso.FolderExists(ResolveLocalWorkbookFolder) Then
        Err.Raise vbObjectError + 104, , "Could not resolve workbook folder to a local synced folder: " & workbookPath
    End If
End Function

Private Function IsHttpPath(ByVal pathText As String) As Boolean
    IsHttpPath = (LCase$(Left$(pathText, 7)) = "http://" Or LCase$(Left$(pathText, 8)) = "https://")
End Function

Private Function ResolveSharePointFolderToLocal(ByVal fso As Object, ByVal urlPath As String) As String
    Dim suffix As String
    Dim roots As Collection
    Dim rootPath As Variant
    Dim candidate As String

    suffix = ExtractSharePointDocumentsSuffix(urlPath)
    If Len(suffix) = 0 Then Exit Function

    Set roots = GetOneDriveRootCandidates(fso)
    For Each rootPath In roots
        candidate = fso.BuildPath(CStr(rootPath), suffix)
        If fso.FolderExists(candidate) Then
            ResolveSharePointFolderToLocal = candidate
            Exit Function
        End If
    Next rootPath
End Function

Private Function ExtractSharePointDocumentsSuffix(ByVal urlPath As String) As String
    Dim decoded As String
    Dim lowerDecoded As String
    Dim pos As Long
    Dim marker As Variant
    Dim markers As Variant
    Dim queryPos As Long

    decoded = UrlDecodeAscii(urlPath)
    queryPos = InStr(1, decoded, "?", vbTextCompare)
    If queryPos > 0 Then decoded = Left$(decoded, queryPos - 1)

    lowerDecoded = LCase$(decoded)
    markers = Array("/documents/", "/shared documents/")

    For Each marker In markers
        pos = InStr(1, lowerDecoded, CStr(marker), vbTextCompare)
        If pos > 0 Then
            ExtractSharePointDocumentsSuffix = Mid$(decoded, pos + Len(CStr(marker)))
            ExtractSharePointDocumentsSuffix = Replace(ExtractSharePointDocumentsSuffix, "/", "\")
            Exit Function
        End If
    Next marker
End Function

Private Function GetOneDriveRootCandidates(ByVal fso As Object) As Collection
    Dim roots As Collection
    Dim shellObj As Object
    Dim env As Object
    Dim varName As Variant
    Dim rootPath As String
    Dim userProfile As String
    Dim userFolder As Object
    Dim subFolder As Object

    Set roots = New Collection
    Set shellObj = CreateObject("WScript.Shell")
    Set env = shellObj.Environment("PROCESS")

    For Each varName In Array("OneDriveCommercial", "OneDriveConsumer", "OneDrive")
        rootPath = CStr(env(CStr(varName)))
        AddRootCandidate roots, fso, rootPath
    Next varName

    userProfile = CStr(env("USERPROFILE"))
    If Len(userProfile) > 0 And fso.FolderExists(userProfile) Then
        Set userFolder = fso.GetFolder(userProfile)
        For Each subFolder In userFolder.SubFolders
            If LCase$(Left$(subFolder.Name, 8)) = "onedrive" Then
                AddRootCandidate roots, fso, subFolder.Path
            End If
        Next subFolder
    End If

    Set GetOneDriveRootCandidates = roots
End Function

Private Sub AddRootCandidate(ByVal roots As Collection, ByVal fso As Object, ByVal rootPath As String)
    Dim existing As Variant

    rootPath = Trim$(rootPath)
    If Len(rootPath) = 0 Then Exit Sub
    If Not fso.FolderExists(rootPath) Then Exit Sub

    For Each existing In roots
        If StrComp(CStr(existing), rootPath, vbTextCompare) = 0 Then Exit Sub
    Next existing

    roots.Add rootPath
End Sub

Private Function UrlDecodeAscii(ByVal encodedText As String) As String
    Dim i As Long
    Dim ch As String
    Dim hexValue As String
    Dim result As String
    Dim bytes() As Byte
    Dim byteCount As Long

    i = 1
    Do While i <= Len(encodedText)
        ch = Mid$(encodedText, i, 1)
        If ch = "%" And i + 2 <= Len(encodedText) Then
            byteCount = 0
            Erase bytes

            Do While i <= Len(encodedText) - 2 And Mid$(encodedText, i, 1) = "%"
                hexValue = Mid$(encodedText, i + 1, 2)
                If Not IsHexPair(hexValue) Then Exit Do

                byteCount = byteCount + 1
                ReDim Preserve bytes(1 To byteCount)
                bytes(byteCount) = CByte(CLng("&H" & hexValue))
                i = i + 3
            Loop

            If byteCount > 0 Then
                result = result & Utf8BytesToString(bytes, byteCount)
            Else
                result = result & ch
                i = i + 1
            End If
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    UrlDecodeAscii = result
End Function

Private Function Utf8BytesToString(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim i As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim b4 As Long
    Dim codePoint As Long

    i = 1
    Do While i <= byteCount
        b1 = CLng(bytes(i))

        If b1 < 128 Then
            Utf8BytesToString = Utf8BytesToString & ChrWValue(b1)
            i = i + 1
        ElseIf b1 >= 192 And b1 < 224 And i + 1 <= byteCount Then
            b2 = CLng(bytes(i + 1))
            codePoint = ((b1 And 31) * 64) + (b2 And 63)
            Utf8BytesToString = Utf8BytesToString & ChrWValue(codePoint)
            i = i + 2
        ElseIf b1 >= 224 And b1 < 240 And i + 2 <= byteCount Then
            b2 = CLng(bytes(i + 1))
            b3 = CLng(bytes(i + 2))
            codePoint = ((b1 And 15) * 4096) + ((b2 And 63) * 64) + (b3 And 63)
            Utf8BytesToString = Utf8BytesToString & ChrWValue(codePoint)
            i = i + 3
        ElseIf b1 >= 240 And b1 < 248 And i + 3 <= byteCount Then
            b2 = CLng(bytes(i + 1))
            b3 = CLng(bytes(i + 2))
            b4 = CLng(bytes(i + 3))
            codePoint = ((b1 And 7) * 262144) + ((b2 And 63) * 4096) + ((b3 And 63) * 64) + (b4 And 63)
            Utf8BytesToString = Utf8BytesToString & CodePointToString(codePoint)
            i = i + 4
        Else
            Utf8BytesToString = Utf8BytesToString & "?"
            i = i + 1
        End If
    Loop
End Function

Private Function CodePointToString(ByVal codePoint As Long) As String
    Dim value As Long
    Dim highSurrogate As Long
    Dim lowSurrogate As Long

    If codePoint <= 65535 Then
        CodePointToString = ChrWValue(codePoint)
    Else
        value = codePoint - 65536
        highSurrogate = 55296 + (value \ 1024)
        lowSurrogate = 56320 + (value Mod 1024)
        CodePointToString = ChrWValue(highSurrogate) & ChrWValue(lowSurrogate)
    End If
End Function

Private Function ChrWValue(ByVal codePoint As Long) As String
    If codePoint > 32767 Then codePoint = codePoint - 65536
    ChrWValue = ChrW$(codePoint)
End Function

Private Function IsHexPair(ByVal valueText As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(valueText) <> 2 Then Exit Function
    For i = 1 To 2
        ch = UCase$(Mid$(valueText, i, 1))
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "F")) Then Exit Function
    Next i

    IsHexPair = True
End Function
