Attribute VB_Name = "SourceDataSorter"
Option Explicit

Private Const OUTPUT_ROOT As String = "Sorted_Data"
Private Const BUTTON_NAME As String = "btnSortSourceData"
Private gStep As String

Public Sub SetupSourceDataSorter()
    Dim ws As Worksheet
    Dim btn As Button
    Set ws = ActiveSheet

    On Error Resume Next
    ws.Buttons(BUTTON_NAME).Delete
    On Error GoTo 0

    Set btn = ws.Buttons.Add(ws.Range("D1").Left, ws.Range("D1").Top, _
                             ws.Range("D1:F1").Width, ws.Rows(1).Height + 5)
    With btn
        .Name = BUTTON_NAME
        .Caption = "Sort"
        .OnAction = "'" & ThisWorkbook.Name & "'!RunSourceDataSorter"
        .Font.Size = 10
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 25
    MsgBox "Sort button created.", vbInformation
End Sub

Public Sub RunSourceDataSorter()
    Dim basePath As String, outputPath As String
    Dim dateFolders As Variant
    Dim i As Long, srcDate As String, srcDatePath As String, outDate As String
    Dim processedDateCount As Long

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gStep = "Resolving workbook folder"
    basePath = ResolveWorkbookLocalPath(ThisWorkbook.Path)
    If Len(basePath) = 0 Then Err.Raise vbObjectError + 801, , "Could not resolve the local workbook folder."

    outputPath = CombinePath(basePath, OUTPUT_ROOT)
    EnsureFolderExists outputPath

    gStep = "Preparing sort logs"
    PrepareSortLog

    gStep = "Finding date folders"
    dateFolders = GetSortedDateFolders(basePath)
    If IsEmpty(dateFolders) Then Err.Raise vbObjectError + 802, , "No date folders were found."

    For i = LBound(dateFolders) To UBound(dateFolders)
        srcDatePath = CStr(dateFolders(i))
        srcDate = GetFolderNameOnly(srcDatePath)
        outDate = NormalizeDateFolderName(srcDate)

        gStep = "Sorting date folder: " & srcDate

        SortOneDateFolder srcDatePath, CombinePath(outputPath, outDate), outDate

        processedDateCount = processedDateCount + 1
        AddSortLogRow srcDate, srcDatePath, outDate, "Processed"
    Next i

    MsgBox "Sorting completed." & vbCrLf & vbCrLf & _
           "Processed date folders: " & processedDateCount & vbCrLf & _
           "Output folder:" & vbCrLf & outputPath, vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Sorting error." & vbCrLf & vbCrLf & _
           "Step: " & gStep & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
    Resume SafeExit
End Sub

Private Sub SortOneDateFolder(ByVal sourceFolder As String, ByVal outputDateFolder As String, ByVal normalizedDate As String)
    Dim files As Variant, tracks As Variant
    Dim headers As Collection, headerIndex As Object, trackRows As Object
    Dim i As Long, trackName As String

    tracks = GetTrackNames()
    Set headers = New Collection
    Set headerIndex = NewDictionary()
    Set trackRows = CreateTrackRowContainer(tracks)

    EnsureFolderExists outputDateFolder
    For i = LBound(tracks) To UBound(tracks)
        EnsureFolderExists CombinePath(outputDateFolder, CStr(tracks(i)))
    Next i

    files = GetSortedXlsxFiles(sourceFolder)
    If IsEmpty(files) Then Exit Sub

    For i = LBound(files) To UBound(files)
        gStep = "Reading headers: " & CStr(files(i))
        CollectHeadersFromWorkbook CombinePath(sourceFolder, CStr(files(i))), headers, headerIndex
    Next i

    If headers.Count = 0 Then Exit Sub

    For i = LBound(files) To UBound(files)
        gStep = "Sorting rows: " & CStr(files(i))
        CollectRowsFromWorkbook CombinePath(sourceFolder, CStr(files(i))), headers, headerIndex, trackRows
    Next i

    For i = LBound(tracks) To UBound(tracks)
        trackName = CStr(tracks(i))
        gStep = "Writing: " & normalizedDate & " / " & trackName
        WriteTrackWorkbook CombinePath(outputDateFolder, trackName), normalizedDate, trackName, headers, trackRows(trackName)
    Next i
End Sub

Private Sub CollectHeadersFromWorkbook(ByVal filePath As String, ByVal headers As Collection, ByVal headerIndex As Object)
    Dim wb As Workbook, ws As Worksheet
    Dim lastCol As Long, c As Long, h As String

    On Error GoTo CleanFail
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    For Each ws In wb.Worksheets
        lastCol = LastUsedColumn(ws)
        For c = 1 To lastCol
            h = CleanText(ws.Cells(1, c).Value)
            If Len(h) > 0 Then
                If Not headerIndex.Exists(LCase$(h)) Then
                    headers.Add h
                    headerIndex.Add LCase$(h), headers.Count
                End If
            End If
        Next c
    Next ws

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub CollectRowsFromWorkbook(ByVal filePath As String, ByVal headers As Collection, ByVal headerIndex As Object, ByVal trackRows As Object)
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim localMap As Object
    Dim cCourse As Long, r As Long, c As Long, destCol As Long
    Dim courseTitle As String, trackName As String, headerText As String
    Dim rowData() As Variant
    Dim rows As Collection

    On Error GoTo CleanFail
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    For Each ws In wb.Worksheets
        lastRow = LastUsedRow(ws)
        lastCol = LastUsedColumn(ws)
        If lastRow < 2 Or lastCol < 1 Then GoTo NextSheet

        Set localMap = NewDictionary()
        For c = 1 To lastCol
            headerText = CleanText(ws.Cells(1, c).Value)
            If Len(headerText) > 0 Then localMap(LCase$(headerText)) = c
        Next c

        If Not localMap.Exists(LCase$("Course title")) Then GoTo NextSheet
        cCourse = CLng(localMap(LCase$("Course title")))

        For r = 2 To lastRow
            courseTitle = CleanText(ws.Cells(r, cCourse).Value)
            trackName = GetTrackForCourse(courseTitle)
            If Len(trackName) = 0 Then GoTo NextRow

            ReDim rowData(1 To headers.Count)
            For c = 1 To lastCol
                headerText = CleanText(ws.Cells(1, c).Value)
                If Len(headerText) > 0 Then
                    If headerIndex.Exists(LCase$(headerText)) Then
                        destCol = CLng(headerIndex(LCase$(headerText)))
                        If StrComp(headerText, "Username", vbTextCompare) = 0 Then
                            rowData(destCol) = CleanText(ws.Cells(r, c).Text)
                        Else
                            rowData(destCol) = ws.Cells(r, c).Value
                        End If
                    End If
                End If
            Next c

            Set rows = trackRows(trackName)
            rows.Add rowData
NextRow:
        Next r
NextSheet:
    Next ws

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub WriteTrackWorkbook( _
    ByVal trackFolder As String, _
    ByVal normalizedDate As String, _
    ByVal trackName As String, _
    ByVal headers As Collection, _
    ByVal rows As Collection)

    Dim wbOut As Workbook
    Dim wsOut As Worksheet
    Dim outputPath As String

    Dim headerData() As Variant
    Dim outputData() As Variant
    Dim item As Variant

    Dim i As Long
    Dim j As Long
    Dim usernameCol As Long
    Dim lastRow As Long

    On Error GoTo WriteFail

    gStep = "Creating output folder: " & normalizedDate & " / " & trackName
    EnsureFolderExists trackFolder

    gStep = "Creating workbook: " & normalizedDate & " / " & trackName
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)

    wsOut.Name = "Data"

    ReDim headerData(1 To 1, 1 To headers.Count)

    For i = 1 To headers.Count

        headerData(1, i) = CStr(headers(i))

        If StrComp(CStr(headers(i)), "Username", vbTextCompare) = 0 Then
            usernameCol = i
        End If

    Next i

    gStep = "Writing headers: " & normalizedDate & " / " & trackName
    wsOut.Cells(1, 1).Resize(1, headers.Count).Value = headerData

    If usernameCol > 0 Then
        wsOut.Columns(usernameCol).NumberFormat = "@"
    End If

    If rows.Count > 0 Then

        ReDim outputData(1 To rows.Count, 1 To headers.Count)

        For i = 1 To rows.Count

            item = rows(i)

            For j = 1 To headers.Count
                outputData(i, j) = item(j)
            Next j

        Next i

        gStep = "Writing data: " & normalizedDate & " / " & trackName
        wsOut.Cells(2, 1).Resize(rows.Count, headers.Count).Value = outputData

    End If

    ' Avoid AutoFilter edge cases on empty data sets.
    gStep = "Formatting workbook: " & normalizedDate & " / " & trackName

    lastRow = 1 + rows.Count

    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, headers.Count))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If wsOut.AutoFilterMode Then
        wsOut.AutoFilterMode = False
    End If

    If rows.Count > 0 Then
        wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastRow, headers.Count)).AutoFilter
    End If

    wsOut.Columns.AutoFit

    For i = 1 To headers.Count
        If wsOut.Columns(i).ColumnWidth > 45 Then
            wsOut.Columns(i).ColumnWidth = 45
        End If
    Next i

    outputPath = CombinePath( _
                    trackFolder, _
                    normalizedDate & "_" & trackName & ".xlsx")

    gStep = "Preparing output file: " & outputPath

    SaveWorkbookThroughTemp _
        wbOut, _
        outputPath, _
        normalizedDate & "_" & trackName

    Exit Sub

WriteFail:

    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    On Error GoTo 0

    Err.Raise _
        errNumber, _
        "WriteTrackWorkbook", _
        errDescription

End Sub

Private Function GetTrackNames() As Variant
    GetTrackNames = Array( _
        "WFD-PreReq", _
        "WFD-Design_for_Test", _
        "WFD-Design_Verification", _
        "WFD-Physical_Design", _
        "WFD-RTL_Synthesis", _
        "WFD-AMS")
End Function

Private Function GetTrackForCourse(ByVal courseTitle As String) As String
    Select Case LCase$(Trim$(courseTitle))
        Case LCase$("Purple Certification: ASIC Design Flow Exam"), _
             LCase$("Purple Certification: Digital Design Fundamentals Exam"), _
             LCase$("Purple Certification: CMOS Fundamentals Exam"), _
             LCase$("Purple Certification: Very Deep Submicron (VDSM) Fundamentals Exam"), _
             LCase$("Purple Certification: VLSI Basics Exam")
            GetTrackForCourse = "WFD-PreReq"

        Case LCase$("TestMAX ATPG Exam"), _
             LCase$("Fusion Compiler: DFT Synthesis Exam"), _
             LCase$("TestMAX Advisor Exam"), _
             LCase$("TestMAX DFT Exam")
            GetTrackForCourse = "WFD-Design_for_Test"

        Case LCase$("SystemVerilog Assertions Exam"), _
             LCase$("SystemVerilog Verification using UVM Exam"), _
             LCase$("SystemVerilog Testbench Exam")
            GetTrackForCourse = "WFD-Design_Verification"

        Case LCase$("Fusion Compiler: Design Implementation Exam"), _
             LCase$("Fusion Compiler: Design Creation and Synthesis Exam")
            GetTrackForCourse = "WFD-Physical_Design"

        Case LCase$("SystemVerilog for RTL Design Exam"), _
             LCase$("Design Compiler NXT: RTL Synthesis Exam"), _
             LCase$("Design Compiler NXT: Low Power Exam"), _
             LCase$("Fusion Compiler: UPF Fundamentals Exam")
            GetTrackForCourse = "WFD-RTL_Synthesis"

        Case LCase$("Custom Compiler: Basic Layout Design Exam"), _
             LCase$("Custom Compiler: Introduction to Platform Exam"), _
             LCase$("Custom Compiler: Schematic Entry Exam"), _
             LCase$("PrimeWave Design Environment Exam")
            GetTrackForCourse = "WFD-AMS"

        Case Else
            GetTrackForCourse = ""
    End Select
End Function

Private Function CreateTrackRowContainer(ByVal tracks As Variant) As Object
    Dim d As Object, rows As Collection, i As Long
    Set d = NewDictionary()
    For i = LBound(tracks) To UBound(tracks)
        Set rows = New Collection
        d.Add CStr(tracks(i)), rows
    Next i
    Set CreateTrackRowContainer = d
End Function

Private Function NewDictionary() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewDictionary = d
End Function

Private Function GetSortedDateFolders(ByVal rootPath As String) As Variant
    Dim paths As Collection
    Dim arr() As String
    Dim i As Long

    Set paths = New Collection

    CollectDateFoldersRecursive rootPath, paths

    If paths.Count = 0 Then
        GetSortedDateFolders = Empty
        Exit Function
    End If

    ReDim arr(1 To paths.Count)

    For i = 1 To paths.Count
        arr(i) = CStr(paths(i))
    Next i

    SortDatePathArray arr
    GetSortedDateFolders = arr
End Function

Private Sub CollectDateFoldersRecursive(ByVal currentPath As String, ByVal results As Collection)
    Dim fso As Object, folderObj As Object, subFolder As Object
    Dim folderName As String
    Dim isDateFolder As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(currentPath)

    For Each subFolder In folderObj.SubFolders
        folderName = CStr(subFolder.Name)

        If StrComp(folderName, OUTPUT_ROOT, vbTextCompare) <> 0 Then
            isDateFolder = IsDateFolderName(folderName)

            AddDiscoveryLogRow folderName, CStr(subFolder.Path), _
                               IIf(isDateFolder, "Date folder", "Scanned folder")

            If isDateFolder Then
                results.Add CStr(subFolder.Path)
            End If

            ' Continue scanning all nested folders regardless of name.
            CollectDateFoldersRecursive CStr(subFolder.Path), results
        End If
    Next subFolder
End Sub

Private Sub SortDatePathArray(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As String
    Dim keyI As String, keyJ As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            keyI = NormalizeDateFolderName(GetFolderNameOnly(CStr(arr(i))))
            keyJ = NormalizeDateFolderName(GetFolderNameOnly(CStr(arr(j))))

            If StrComp(keyI, keyJ, vbTextCompare) > 0 Then
                tmp = CStr(arr(i))
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function GetSortedXlsxFiles(ByVal folderPath As String) As Variant
    Dim fso As Object, folderObj As Object, fileObj As Object
    Dim arr() As String, count As Long

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

    Dim normalized As String
    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long
    Dim dt As Date

    On Error GoTo InvalidDate

    normalized = NormalizeDateFolderName(folderName)

    If Len(normalized) <> 8 Then Exit Function
    If Not IsNumeric(normalized) Then Exit Function

    yyyy = CLng(Left$(normalized, 4))
    mm = CLng(Mid$(normalized, 5, 2))
    dd = CLng(Right$(normalized, 2))

    If yyyy < 1900 Or yyyy > 2100 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    dt = DateSerial(yyyy, mm, dd)

    If Year(dt) <> yyyy Then Exit Function
    If Month(dt) <> mm Then Exit Function
    If Day(dt) <> dd Then Exit Function

    IsDateFolderName = True
    Exit Function

InvalidDate:

    IsDateFolderName = False

End Function

Private Function NormalizeDateFolderName(ByVal folderName As String) As String

    Dim s As String
    Dim parts() As String
    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long

    On Error GoTo Failed

    s = Trim$(folderName)

    ' Normalize common separators to underscore.
    s = Replace(s, "-", "_")
    s = Replace(s, ".", "_")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(&H3000), "")

    ' Case 1: already YYYYMMDD.
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

    ' Case 2: YYYY_M_D, YYYY_MM_DD, etc.
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

Private Sub SortStringArray(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As String
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

Private Sub FormatOutputSheet(ByVal ws As Worksheet, ByVal headerCount As Long)
    Dim lastRow As Long, c As Long
    lastRow = LastUsedRow(ws)

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, headerCount))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, headerCount)).AutoFilter
    ws.Columns.AutoFit

    For c = 1 To headerCount
        If ws.Columns(c).ColumnWidth > 45 Then ws.Columns(c).ColumnWidth = 45
    Next c
End Sub

Private Sub PrepareSortLog()
    Dim ws As Worksheet
    Dim wsDiscovery As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sort_Log")
    Set wsDiscovery = ThisWorkbook.Worksheets("Sort_Discovery")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Sort_Log"
    Else
        ws.Cells.Clear
    End If

    If wsDiscovery Is Nothing Then
        Set wsDiscovery = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsDiscovery.Name = "Sort_Discovery"
    Else
        wsDiscovery.Cells.Clear
    End If

    ws.Range("A1:D1").Value = Array("Source Date Folder", "Source Path", "Output Date Folder", "Status")
    ws.Rows(1).Font.Bold = True

    wsDiscovery.Range("A1:D1").Value = Array("Folder Name", "Normalized", "Folder Path", "Recognition")
    wsDiscovery.Rows(1).Font.Bold = True
    wsDiscovery.Cells(2, 1).Value = "Search Root"
    wsDiscovery.Cells(2, 3).Value = ResolveWorkbookLocalPath(ThisWorkbook.Path)
End Sub

Private Sub AddSortLogRow(ByVal sourceDateFolder As String, ByVal sourcePath As String, _
                          ByVal outputDateFolder As String, ByVal statusText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Sort_Log")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = sourceDateFolder
    ws.Cells(nextRow, 2).Value = sourcePath
    ws.Cells(nextRow, 3).Value = outputDateFolder
    ws.Cells(nextRow, 4).Value = statusText
    ws.Columns.AutoFit
End Sub

Private Sub AddDiscoveryLogRow(ByVal folderName As String, _
    ByVal folderPath As String, ByVal recognition As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Sort_Discovery")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = folderName
    ws.Cells(nextRow, 2).NumberFormat = "@"
    ws.Cells(nextRow, 2).Value = NormalizeDateFolderName(folderName)
    ws.Cells(nextRow, 3).Value = folderPath
    ws.Cells(nextRow, 4).Value = recognition
End Sub

Private Function GetFolderNameOnly(ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderNameOnly = fso.GetFolder(folderPath).Name
End Function


Private Sub SaveWorkbookThroughTemp( _
    ByVal wbOut As Workbook, _
    ByVal finalPath As String, _
    ByVal baseFileName As String)

    Dim fso As Object
    Dim tempFolder As String
    Dim tempPath As String
    Dim safeName As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    tempFolder = Environ$("TEMP")

    If Len(tempFolder) = 0 Then
        tempFolder = Environ$("TMP")
    End If

    If Len(tempFolder) = 0 Then
        Err.Raise vbObjectError + 820, , _
            "Windows temporary folder could not be found."
    End If

    safeName = MakeSafeFileName(baseFileName)

    tempPath = CombinePath( _
        tempFolder, _
        safeName & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx")

    If FileExists(tempPath) Then
        Kill tempPath
    End If

    gStep = "Saving temporary workbook: " & tempPath

    wbOut.SaveAs _
        Filename:=tempPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False

    gStep = "Copying workbook to output folder: " & finalPath

    If FileExists(finalPath) Then
        fso.DeleteFile finalPath, True
    End If

    fso.CopyFile tempPath, finalPath, True

    ' Close the workbook before deleting the temporary file.
    gStep = "Closing temporary workbook: " & tempPath
    wbOut.Close SaveChanges:=False

    gStep = "Removing temporary workbook: " & tempPath

    If FileExists(tempPath) Then
        fso.DeleteFile tempPath, True
    End If

End Sub


Private Function MakeSafeFileName(ByVal fileName As String) As String

    Dim s As String

    s = fileName

    s = Replace(s, "\\", "_")
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


Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

Private Function CombinePath(ByVal parentPath As String, ByVal childName As String) As String
    If Right$(parentPath, 1) = "\" Then
        CombinePath = parentPath & childName
    Else
        CombinePath = parentPath & "\" & childName
    End If
End Function

Private Function ResolveWorkbookLocalPath(ByVal workbookPath As String) As String
    Dim relativePath As String, oneDriveRoot As String, testPath As String

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
        testPath = CombinePath(Environ$("USERPROFILE") & "\Desktop", Mid$(relativePath, Len("Desktop\") + 1))
        If FolderExists(testPath) Then
            ResolveWorkbookLocalPath = testPath
            Exit Function
        End If
    End If

    testPath = CombinePath(oneDriveRoot, "Documents\" & relativePath)
    If FolderExists(testPath) Then ResolveWorkbookLocalPath = testPath
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

Private Sub AddCandidate(ByRef candidates As Collection, ByVal candidatePath As String)
    If Len(Trim$(candidatePath)) > 0 Then candidates.Add Trim$(candidatePath)
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

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, _
                                  LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If foundCell Is Nothing Then LastUsedRow = 1 Else LastUsedRow = foundCell.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, _
                                  LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
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
