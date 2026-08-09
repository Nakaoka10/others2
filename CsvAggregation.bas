Option Explicit

'============================================================
' 設定
'============================================================

Private Const CSV_ROOT_FOLDER As String = "csv"
Private Const TARGET_TRACK As String = "WFD-PreReq"

Private Const SHEET_USERS As String = "ユーザー一覧"
Private Const SHEET_PASSED As String = "合格者一覧"

Private Const HEADER_ROW As Long = 3

' Dictionary用区切り文字
Private Const KEY_SEPARATOR As String = "|||"


'============================================================
' 初期設定
'
' 最初に1回だけ実行してください。
' ・「ユーザー一覧」シートを作成
' ・「合格者一覧」シートを作成
' ・1行目に実行ボタンを配置
'============================================================
Public Sub SetupCsvAggregation()

    Dim wsUsers As Worksheet
    Dim wsPassed As Worksheet

    Application.ScreenUpdating = False

    Set wsUsers = GetOrCreateSheet(SHEET_USERS)
    Set wsPassed = GetOrCreateSheet(SHEET_PASSED)

    ' ヘッダー
    WriteUserHeaders wsUsers
    WritePassedBaseHeaders wsPassed

    ' 実行ボタン作成
    CreateRunButton wsUsers

    ' 見た目
    FormatSheet wsUsers
    FormatSheet wsPassed

    Application.ScreenUpdating = True

    MsgBox _
        "初期設定が完了しました。" & vbCrLf & vbCrLf & _
        "今後は「" & SHEET_USERS & "」シート1行目の" & vbCrLf & _
        "「CSV集計を実行」ボタンを押してください。", _
        vbInformation

End Sub


'============================================================
' メイン処理
'
' Sheet1のボタンから呼び出されます。
'============================================================
Public Sub RunCsvAggregation()

    Dim basePath As String
    Dim csvRootPath As String

    Dim wsUsers As Worksheet
    Dim wsPassed As Worksheet

    ' 全ユーザー
    Dim users As Object
    Dim userOrder As Collection

    ' 合格者
    Dim passedUsers As Object
    Dim passedOrder As Collection

    ' 試験情報
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

    '--------------------------------------------------------
    ' 高速化
    '--------------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    '--------------------------------------------------------
    ' パス
    '--------------------------------------------------------
    basePath = ThisWorkbook.Path
    csvRootPath = basePath & "\" & CSV_ROOT_FOLDER

    If Dir(csvRootPath, vbDirectory) = "" Then

        MsgBox _
            "csvフォルダが見つかりません。" & vbCrLf & vbCrLf & _
            "以下の位置に配置してください。" & vbCrLf & _
            csvRootPath, _
            vbExclamation

        GoTo SafeExit

    End If

    '--------------------------------------------------------
    ' シート取得
    '--------------------------------------------------------
    Set wsUsers = GetOrCreateSheet(SHEET_USERS)
    Set wsPassed = GetOrCreateSheet(SHEET_PASSED)

    '--------------------------------------------------------
    ' Dictionary作成
    '--------------------------------------------------------

    ' 全ユーザー
    Set users = CreateObject("Scripting.Dictionary")
    users.CompareMode = vbTextCompare

    Set userOrder = New Collection

    ' 合格者基本情報
    Set passedUsers = CreateObject("Scripting.Dictionary")
    passedUsers.CompareMode = vbTextCompare

    Set passedOrder = New Collection

    ' 試験ごとの情報
    Set examData = CreateObject("Scripting.Dictionary")
    examData.CompareMode = vbTextCompare

    ' 各ユーザーの試験順序
    Set examOrders = CreateObject("Scripting.Dictionary")
    examOrders.CompareMode = vbTextCompare

    '--------------------------------------------------------
    ' 日付フォルダ一覧取得
    '--------------------------------------------------------
    dateFolders = GetSortedDateFolders(csvRootPath)

    If IsEmpty(dateFolders) Then

        MsgBox _
            "処理対象となる日付フォルダがありません。" & vbCrLf & _
            "例：csv\20260806\WFD-PreReq\", _
            vbExclamation

        GoTo SafeExit

    End If

    '--------------------------------------------------------
    ' 日付フォルダを古い順に処理
    '--------------------------------------------------------
    For i = LBound(dateFolders) To UBound(dateFolders)

        dateFolderName = CStr(dateFolders(i))

        trackFolderPath = _
            csvRootPath & "\" & _
            dateFolderName & "\" & _
            TARGET_TRACK

        ' WFD-PreReqが存在するときだけ処理
        If Dir(trackFolderPath, vbDirectory) <> "" Then

            csvFiles = GetSortedCsvFiles(trackFolderPath)

            If Not IsEmpty(csvFiles) Then

                For j = LBound(csvFiles) To UBound(csvFiles)

                    csvFilePath = _
                        trackFolderPath & "\" & _
                        CStr(csvFiles(j))

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

    '--------------------------------------------------------
    ' Excelへ出力
    '--------------------------------------------------------
    OutputUsers wsUsers, users, userOrder

    OutputPassedUsers _
        wsPassed, _
        passedUsers, _
        passedOrder, _
        examData, _
        examOrders

    '--------------------------------------------------------
    ' 完了
    '--------------------------------------------------------
    MsgBox _
        "CSV集計が完了しました。" & vbCrLf & vbCrLf & _
        "処理CSVファイル数：" & processedFileCount & vbCrLf & _
        "スキップファイル数：" & skippedFileCount & vbCrLf & _
        "全ユーザー数：" & users.Count & vbCrLf & _
        "合格者数：" & passedUsers.Count, _
        vbInformation

SafeExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub


ErrorHandler:

    MsgBox _
        "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
        "エラー番号：" & Err.Number & vbCrLf & _
        "内容：" & Err.Description, _
        vbCritical

    Resume SafeExit

End Sub


'============================================================
' CSV 1ファイルを処理
'============================================================
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

    '--------------------------------------------------------
    ' CSVを開く
    '--------------------------------------------------------
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

    '--------------------------------------------------------
    ' ヘッダー位置を検索
    '--------------------------------------------------------
    colUsername = FindHeaderColumn(data, "Username")
    colEmail = FindHeaderColumn(data, "Email")
    colFullName = FindHeaderColumn(data, "Full Name")

    colCourseTitle = FindHeaderColumn(data, "Course title")

    colEnrollmentDate = _
        FindHeaderColumn(data, "Enrollment Date")

    colEnrollmentEndDate = _
        FindHeaderColumn(data, "Enrollment End Date")

    colStatus = _
        FindHeaderColumn(data, "Course Enrollment Status")

    colFinalScore = _
        FindHeaderColumn(data, "Final Score")

    ' 必須列
    If colUsername = 0 _
        Or colEmail = 0 _
        Or colFullName = 0 _
        Or colEnrollmentDate = 0 _
        Or colEnrollmentEndDate = 0 Then

        GoTo FileError

    End If

    '--------------------------------------------------------
    ' 行処理
    '--------------------------------------------------------
    For r = 2 To UBound(data, 1)

        username = CleanText(data(r, colUsername))
        email = CleanText(data(r, colEmail))
        fullName = CleanText(data(r, colFullName))

        enrollmentDate = data(r, colEnrollmentDate)
        enrollmentEndDate = data(r, colEnrollmentEndDate)

        '----------------------------------------------------
        ' 人物キー作成
        '----------------------------------------------------
        userKey = MakeUserKey(username, email, fullName)

        If Len(userKey) = 0 Then
            GoTo NextRow
        End If

        '====================================================
        ' Sheet1用：全ユーザー
        '====================================================
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

            ' 日付フォルダを古い→新しい順で処理しているので
            ' 後から見つかった情報で更新
            userInfo = users(userKey)

            userInfo(0) = username
            userInfo(1) = email
            userInfo(2) = fullName
            userInfo(3) = enrollmentDate
            userInfo(4) = enrollmentEndDate

            users(userKey) = userInfo

        End If

        '====================================================
        ' Sheet2用：合格者
        '====================================================

        ' Sheet2に必要な列が存在する場合
        If colCourseTitle > 0 _
            And colStatus > 0 _
            And colFinalScore > 0 Then

            courseTitle = CleanText(data(r, colCourseTitle))
            courseStatus = CleanText(data(r, colStatus))
            finalScore = data(r, colFinalScore)

            ' Completedのみ
            If StrComp(courseStatus, "Completed", vbTextCompare) = 0 Then

                ' 対象5試験か判定
                examName = GetTargetExamName(courseTitle)

                If Len(examName) > 0 Then

                    '----------------------------------------
                    ' 合格者基本情報
                    '----------------------------------------
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

                    '----------------------------------------
                    ' 試験情報
                    '----------------------------------------
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

                        ' 同じ試験が後の日付にも存在
                        ' → 後のデータで更新
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


'============================================================
' Sheet1へ全ユーザー出力
'============================================================
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


'============================================================
' Sheet2へ合格者出力
'============================================================
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

    '--------------------------------------------------------
    ' 1人あたり最大何試験あるか
    '--------------------------------------------------------
    maxExamCount = 0

    For i = 1 To passedOrder.Count

        key = CStr(passedOrder(i))

        Set orderCollection = examOrders(key)

        If orderCollection.Count > maxExamCount Then
            maxExamCount = orderCollection.Count
        End If

    Next i

    '--------------------------------------------------------
    ' ヘッダー
    '--------------------------------------------------------
    WritePassedHeaders ws, maxExamCount

    If passedUsers.Count = 0 Then
        FormatSheet ws
        Exit Sub
    End If

    totalColumns = 6 + (maxExamCount * 3)

    ReDim outputData( _
        1 To passedUsers.Count, _
        1 To totalColumns)

    '--------------------------------------------------------
    ' データ
    '--------------------------------------------------------
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


'============================================================
' 対象試験判定
'
' 戻り値：
' 対象試験 → 試験名
' 対象外   → ""
'============================================================
Private Function GetTargetExamName( _
    ByVal courseTitle As String) As String

    Dim title As String

    title = Trim$(courseTitle)

    ' Purple Certification: を除去
    If InStr(1, title, _
        "Purple Certification:", _
        vbTextCompare) = 1 Then

        title = Trim$(Mid$( _
            title, _
            Len("Purple Certification:") + 1))

    End If

    Select Case LCase$(title)

        Case LCase$("Digital Design Fundamentals Exam")
            GetTargetExamName = _
                "Digital Design Fundamentals Exam"

        Case LCase$("ASIC Design Flow Exam")
            GetTargetExamName = _
                "ASIC Design Flow Exam"

        Case LCase$("CMOS Fundamentals Exam")
            GetTargetExamName = _
                "CMOS Fundamentals Exam"

        Case LCase$("VLSI Basics Exam")
            GetTargetExamName = _
                "VLSI Basics Exam"

        Case LCase$( _
            "Very Deep Submicron (VDSM) Fundamentals Exam")

            GetTargetExamName = _
                "Very Deep Submicron (VDSM) Fundamentals Exam"

        Case Else

            GetTargetExamName = ""

    End Select

End Function


'============================================================
' ユーザー識別キー
'
' 基本：
' Username + Email
'
' 両方空欄の場合のみFull Nameを使用
'============================================================
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


'============================================================
' ヘッダー列検索
'============================================================
Private Function FindHeaderColumn( _
    ByVal data As Variant, _
    ByVal headerName As String) As Long

    Dim c As Long
    Dim text As String

    For c = 1 To UBound(data, 2)

        text = CleanText(data(1, c))

        ' UTF-8 BOM対策
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


'============================================================
' 日付フォルダ取得・昇順ソート
'============================================================
Private Function GetSortedDateFolders( _
    ByVal rootPath As String) As Variant

    Dim folderName As String
    Dim arr() As String

    Dim count As Long

    folderName = Dir(rootPath & "\*", vbDirectory)

    Do While Len(folderName) > 0

        If folderName <> "." _
            And folderName <> ".." Then

            If (GetAttr( _
                rootPath & "\" & folderName) _
                And vbDirectory) = vbDirectory Then

                If IsDateFolderName(folderName) Then

                    count = count + 1
                    ReDim Preserve arr(1 To count)

                    arr(count) = folderName

                End If

            End If

        End If

        folderName = Dir()

    Loop

    If count = 0 Then

        GetSortedDateFolders = Empty
        Exit Function

    End If

    SortStringArray arr

    GetSortedDateFolders = arr

End Function


'============================================================
' CSVファイル一覧取得・ファイル名順
'============================================================
Private Function GetSortedCsvFiles( _
    ByVal folderPath As String) As Variant

    Dim fileName As String
    Dim arr() As String

    Dim count As Long

    fileName = Dir(folderPath & "\*.csv")

    Do While Len(fileName) > 0

        count = count + 1

        ReDim Preserve arr(1 To count)

        arr(count) = fileName

        fileName = Dir()

    Loop

    If count = 0 Then

        GetSortedCsvFiles = Empty
        Exit Function

    End If

    SortStringArray arr

    GetSortedCsvFiles = arr

End Function


'============================================================
' 文字列配列昇順ソート
'============================================================
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


'============================================================
' YYYYMMDD形式か確認
'============================================================
Private Function IsDateFolderName( _
    ByVal folderName As String) As Boolean

    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long

    On Error GoTo InvalidDate

    If Len(folderName) <> 8 Then
        Exit Function
    End If

    If Not IsNumeric(folderName) Then
        Exit Function
    End If

    yyyy = CLng(Left$(folderName, 4))
    mm = CLng(Mid$(folderName, 5, 2))
    dd = CLng(Right$(folderName, 2))

    ' DateSerialで実在日付か確認
    If Format$(DateSerial(yyyy, mm, dd), "yyyymmdd") _
        <> folderName Then

        Exit Function

    End If

    IsDateFolderName = True

    Exit Function

InvalidDate:

    IsDateFolderName = False

End Function


'============================================================
' Sheet1ヘッダー
'============================================================
Private Sub WriteUserHeaders(ByVal ws As Worksheet)

    ws.Cells(HEADER_ROW, 1).Value = "No."
    ws.Cells(HEADER_ROW, 2).Value = "Username"
    ws.Cells(HEADER_ROW, 3).Value = "Email"
    ws.Cells(HEADER_ROW, 4).Value = "Full Name"
    ws.Cells(HEADER_ROW, 5).Value = "Enrollment Date"
    ws.Cells(HEADER_ROW, 6).Value = "Enrollment End Date"

End Sub


'============================================================
' Sheet2基本ヘッダー
'============================================================
Private Sub WritePassedBaseHeaders(ByVal ws As Worksheet)

    WritePassedHeaders ws, 0

End Sub


'============================================================
' Sheet2ヘッダー
'============================================================
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

        ws.Cells(HEADER_ROW, c).Value = _
            "Course title" & i

        ws.Cells(HEADER_ROW, c + 1).Value = _
            "Course Enrollment Status" & i

        ws.Cells(HEADER_ROW, c + 2).Value = _
            "Final Score" & i

        c = c + 3

    Next i

End Sub


'============================================================
' 出力領域クリア
'
' 1～2行目は残す
' 3行目以降だけクリア
'============================================================
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


'============================================================
' シート取得または作成
'============================================================
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


'============================================================
' 実行ボタン作成
'============================================================
Private Sub CreateRunButton(ByVal ws As Worksheet)

    Dim btn As Button
    Dim leftPos As Double
    Dim topPos As Double
    Dim btnWidth As Double
    Dim btnHeight As Double

    ' 既存ボタン削除
    On Error Resume Next
    ws.Buttons("btnRunCsvAggregation").Delete
    On Error GoTo 0

    leftPos = ws.Range("A1").Left
    topPos = ws.Range("A1").Top

    ' A1～C1程度の横幅
    btnWidth = _
        ws.Range("A1:C1").Width

    btnHeight = _
        ws.Rows(1).Height + 5

    Set btn = ws.Buttons.Add( _
                leftPos, _
                topPos, _
                btnWidth, _
                btnHeight)

    With btn

        .Name = "btnRunCsvAggregation"

        .Caption = "CSV集計を実行"

        .OnAction = _
            "'" & ThisWorkbook.Name & _
            "'!RunCsvAggregation"

        .Font.Size = 11
        .Font.Bold = True

    End With

    ws.Rows(1).RowHeight = 25

End Sub


'============================================================
' シート書式
'============================================================
Private Sub FormatSheet(ByVal ws As Worksheet)

    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = LastUsedColumn(ws)
    lastRow = LastUsedRow(ws)

    If lastCol < 1 Then Exit Sub

    ' ヘッダー
    With ws.Range( _
        ws.Cells(HEADER_ROW, 1), _
        ws.Cells(HEADER_ROW, lastCol))

        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ' オートフィルター
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    If lastRow >= HEADER_ROW Then

        ws.Range( _
            ws.Cells(HEADER_ROW, 1), _
            ws.Cells(lastRow, lastCol)) _
            .AutoFilter

    End If

    ' 列幅自動調整
    ws.Columns.AutoFit

    ' Course titleが広くなりすぎないよう制限
    Dim c As Long

    For c = 1 To lastCol

        If ws.Columns(c).ColumnWidth > 45 Then
            ws.Columns(c).ColumnWidth = 45
        End If

    Next c

End Sub


'============================================================
' 最終使用行
'============================================================
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


'============================================================
' 最終使用列
'============================================================
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


'============================================================
' Null/Error等を考慮して文字列化
'============================================================
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