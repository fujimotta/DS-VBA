VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DSWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' DSWorksheet
' ------------------------------------------------------------
' ワークシート上のデータを行単位で操作するクラス。
' Init → MoveNext/MovePrevious/MoveReset でイテレーションし、
' Value プロパティで列値の取得・更新、Add/AddCommit で行追加、
' Delete/DeleteCommit で行の遅延一括削除を行う。
' ============================================================

' --- 内部状態 ---
Private targetSheet       As Worksheet
Private targetWorkbook    As Workbook
Private headerRow         As Long
Private keyColumn         As Long
Private dataStartRow      As Long
Private dataEndRow        As Long
Private columnCount       As Long
Private currentRowIndex   As Long          ' データ部における現在行（1-based）
Private isInitialized     As Boolean
Private ownsWorkbook      As Boolean       ' このクラスがブックを開いたか

' ヘッダー名_N → 列インデックスのマッピング（Collection）
Private headerMap         As Collection

' 新規行バッファ（1-based, 2次元配列: 行・列）
Private addBuffer()       As Variant
Private isAddBufferReady  As Boolean
Private maxAddRows        As Long
Private currentAddRow     As Long

' 遅延削除用 Range
Private deleteRange       As Range

' ============================================================
'  Public Properties（読み取り専用）
' ============================================================

Public Property Get Sheet() As Worksheet
    Call EnsureInitialized
    Set Sheet = targetSheet
End Property

Public Property Get RowCount() As Long
    Call EnsureInitialized
    RowCount = dataEndRow - dataStartRow + 1
End Property

Public Property Get ColCount() As Long
    Call EnsureInitialized
    ColCount = columnCount
End Property

Public Property Get CurrentRow() As Long
    Call EnsureInitialized
    CurrentRow = currentRowIndex
End Property

Public Property Get HasNext() As Boolean
    Call EnsureInitialized
    HasNext = (currentRowIndex < RowCount)
End Property

Public Property Get HasPrevious() As Boolean
    Call EnsureInitialized
    HasPrevious = (currentRowIndex > 1)
End Property

' ============================================================
'  Cell プロパティ（現在行の Range オブジェクトを直接取得）
' ============================================================

Public Property Get Cell(ByVal columnIdentifier As Variant, Optional ByVal occurrence As Long = 1) As Range
    Call EnsureCursorValid
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    Dim absoluteRow As Long
    absoluteRow = dataStartRow + currentRowIndex - 1
    
    Set Cell = targetSheet.Cells(absoluteRow, colIndex)
End Property

' ============================================================
'  Value プロパティ（現在行の列値を取得・更新）
' ============================================================

Public Property Get Value(ByVal columnIdentifier As Variant, Optional ByVal occurrence As Long = 1) As Variant
    Call EnsureCursorValid
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    Dim absoluteRow As Long
    absoluteRow = dataStartRow + currentRowIndex - 1
    
    Value = targetSheet.Cells(absoluteRow, colIndex).Value
End Property

Public Property Let Value(ByVal columnIdentifier As Variant, Optional ByVal occurrence As Long = 1, ByVal newValue As Variant)
    Call EnsureCursorValid
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    Dim absoluteRow As Long
    absoluteRow = dataStartRow + currentRowIndex - 1
    
    targetSheet.Cells(absoluteRow, colIndex).Value = newValue
End Property

' ============================================================
'  Add プロパティ / AddCommit メソッド
' ============================================================

Public Sub AddReset(Optional ByVal bufferRows As Long = 131)
    Call EnsureInitialized
    If bufferRows < 1 Then bufferRows = 1
    maxAddRows = bufferRows
    ReDim addBuffer(1 To maxAddRows, 1 To columnCount)
    currentAddRow = 1
    isAddBufferReady = True
End Sub

Public Sub AddNextRow()
    Call EnsureInitialized
    If Not isAddBufferReady Then
        Call AddReset
    End If
    
    If currentAddRow >= maxAddRows Then
        Call AddCommit
    Else
        currentAddRow = currentAddRow + 1
    End If
End Sub

Public Property Let Add(ByVal columnIdentifier As Variant, Optional ByVal occurrence As Long = 1, ByVal newValue As Variant)
    Call EnsureInitialized
    If Not isAddBufferReady Then
        Call AddReset
    End If
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    addBuffer(currentAddRow, colIndex) = newValue
End Property

Public Sub AddCommit()
    Call EnsureInitialized
    If Not isAddBufferReady Then Exit Sub
    
    ' 書き込む行数を決定（現在の行が空かチェック）
    Dim rowsToWrite As Long
    rowsToWrite = currentAddRow
    
    Dim isEmptyRow As Boolean
    isEmptyRow = True
    
    Dim c As Long
    For c = 1 To columnCount
        If Not IsEmpty(addBuffer(currentAddRow, c)) Then
            isEmptyRow = False
            Exit For
        End If
    Next c
    
    If isEmptyRow Then
        rowsToWrite = currentAddRow - 1
    End If
    
    ' 1行もデータがなければバッファをリセットして終了
    If rowsToWrite < 1 Then
        Call AddReset(maxAddRows)
        Exit Sub
    End If
    
    ' 書き込み用配列を生成
    Dim outputArray() As Variant
    ReDim outputArray(1 To rowsToWrite, 1 To columnCount)
    
    Dim r As Long
    For r = 1 To rowsToWrite
        For c = 1 To columnCount
            outputArray(r, c) = addBuffer(r, c)
        Next c
    Next r
    
    ' 一括書き込み
    Dim newRow As Long
    newRow = dataEndRow + 1
    
    targetSheet.Cells(newRow, 1).Resize(rowsToWrite, columnCount).Value = outputArray
    
    ' データ範囲を拡張
    dataEndRow = dataEndRow + rowsToWrite
    
    ' バッファ再初期化
    Call AddReset(maxAddRows)
End Sub

' ============================================================
'  Delete / DeleteCommit メソッド
' ============================================================

Public Sub Delete()
    Call EnsureCursorValid
    
    Dim absoluteRow As Long
    absoluteRow = dataStartRow + currentRowIndex - 1
    
    Dim rowRange As Range
    Set rowRange = targetSheet.Rows(absoluteRow)
    
    If deleteRange Is Nothing Then
        Set deleteRange = rowRange
    Else
        Set deleteRange = Union(deleteRange, rowRange)
    End If
End Sub

Public Sub DeleteCommit()
    Call EnsureInitialized
    If deleteRange Is Nothing Then
        Exit Sub
    End If
    
    Call deleteRange.Delete(Shift:=xlUp)
    Set deleteRange = Nothing
    
    ' データ範囲を再計算
    Call RecalculateDataRange
    
    ' カーソルをリセット
    Call MoveReset
End Sub

' ============================================================
'  ブック保存メソッド
' ============================================================

Public Sub SaveWorkbook()
    Call EnsureInitialized
    If Not targetWorkbook Is Nothing Then
        Call targetWorkbook.Save
    End If
End Sub

' ============================================================
'  DS 操作メソッド (Sort, Format, Refresh)
' ============================================================

Public Sub Sort(ByVal columnIdentifier As Variant, Optional ByVal order As XlSortOrder = xlAscending, Optional ByVal occurrence As Long = 1)
    Call EnsureInitialized
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    Dim dataRange As Range
    Set dataRange = targetSheet.Rows(dataStartRow & ":" & dataEndRow)
    
    Call dataRange.Sort(Key1:=targetSheet.Cells(dataStartRow, colIndex), _
                        Order1:=order, Header:=xlNo)
    
    ' ソート後にカーソルをリセット
    Call MoveReset
End Sub

Public Sub SetFormat(ByVal columnIdentifier As Variant, ByVal formatString As String, Optional ByVal occurrence As Long = 1)
    Call EnsureInitialized
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    targetSheet.Columns(colIndex).NumberFormat = formatString
End Sub

Public Sub Refresh()
    Call EnsureInitialized
    
    ' シート内のすべての QueryTable を更新
    Dim qt As QueryTable
    For Each qt In targetSheet.QueryTables
        Call qt.Refresh(BackgroundQuery:=False)
    Next qt
    
    ' ListObject (テーブル) も更新
    Dim lo As ListObject
    For Each lo In targetSheet.ListObjects
        Call lo.Refresh
    Next lo
    
    ' データ範囲が変わる可能性があるため再計算
    Call RecalculateDataRange
    Call MoveReset
End Sub

' ============================================================
'  フィルタ操作
' ============================================================

Public Sub ApplyFilter(ByVal columnIdentifier As Variant, ByVal Criteria1 As Variant, _
                       Optional ByVal Operator As XlAutoFilterOperator = 0, _
                       Optional ByVal Criteria2 As Variant, _
                       Optional ByVal occurrence As Long = 1)
    Call EnsureInitialized
    
    Dim colIndex As Long
    colIndex = ResolveColumnIndex(columnIdentifier, occurrence)
    
    ' ヘッダー行を含めた全範囲にオートフィルタ適用
    Dim fullRange As Range
    Set fullRange = targetSheet.Range(targetSheet.Cells(headerRow, 1), targetSheet.Cells(dataEndRow, columnCount))
    
    If Operator = 0 Then
        Call fullRange.AutoFilter(Field:=colIndex, Criteria1:=Criteria1)
    Else
        Call fullRange.AutoFilter(Field:=colIndex, Criteria1:=Criteria1, Operator:=Operator, Criteria2:=Criteria2)
    End If
End Sub

Public Sub ClearFilter()
    Call EnsureInitialized
    If targetSheet.AutoFilterMode Then
        targetSheet.ShowAllData
    End If
End Sub

' ============================================================
'  配列一括操作
' ============================================================

Public Function ToArray() As Variant()
    Call EnsureInitialized
    If dataEndRow < dataStartRow Then
        ToArray = Array()
        Exit Function
    End If
    
    ToArray = targetSheet.Range(targetSheet.Cells(dataStartRow, 1), targetSheet.Cells(dataEndRow, columnCount)).Value
End Function

Public Sub FromArray(ByVal arr As Variant, Optional ByVal targetRow As Long = 0)
    Call EnsureInitialized
    
    Dim destRow As Long
    If targetRow <= 0 Then
        ' 末尾追記
        destRow = dataEndRow + 1
    Else
        ' 指定行（1-based なのでデータ開始行 + offset）
        destRow = dataStartRow + targetRow - 1
    End If
    
    Dim rowCountArr As Long
    Dim colCountArr As Long
    
    On Error Resume Next
    rowCountArr = UBound(arr, 1) - LBound(arr, 1) + 1
    colCountArr = UBound(arr, 2) - LBound(arr, 2) + 1
    On Error GoTo 0
    
    If rowCountArr = 0 Or colCountArr = 0 Then Exit Sub
    
    ' 列数が現在のシート定義と一致するかチェック（緩やかなチェック）
    If colCountArr > columnCount Then
        columnCount = colCountArr ' 拡張を許可
    End If
    
    targetSheet.Cells(destRow, 1).Resize(rowCountArr, colCountArr).Value = arr
    
    ' 範囲再計算
    Call RecalculateDataRange
End Sub

' ============================================================
'  イテレータ操作
' ============================================================

Public Function MoveNext() As Boolean
    Call EnsureInitialized
    currentRowIndex = currentRowIndex + 1
    MoveNext = (currentRowIndex <= RowCount)
End Function

Public Function MovePrevious() As Boolean
    Call EnsureInitialized
    currentRowIndex = currentRowIndex - 1
    MovePrevious = (currentRowIndex >= 1)
End Function

Public Sub MoveReset()
    currentRowIndex = 0
End Sub

' ============================================================
'  行操作 (Copy / Move)
' ============================================================

Public Sub CopyRowTo(ByVal destDS As DSWorksheet)
    Call EnsureCursorValid
    
    ' 現在行を配列として取得
    Dim rowData() As Variant
    ReDim rowData(1 To 1, 1 To columnCount)
    
    Dim c As Long
    For c = 1 To columnCount
        rowData(1, c) = Me.Value(c)
    Next c
    
    ' 宛先の DS に追記
    Call destDS.FromArray(rowData, 0)
End Sub

Public Sub MoveRowTo(ByVal destDS As DSWorksheet)
    Call CopyRowTo(destDS)
    Call Me.Delete
End Sub

' ============================================================
'  ForEach パターン
' ============================================================

Public Sub ForEach(ByVal callbackObj As Object, ByVal methodName As String)
    Call EnsureInitialized
    
    Dim prevAppUpdating As Boolean
    prevAppUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Call Me.MoveReset
    Do While Me.MoveNext
        CallByName callbackObj, methodName, VbMethod, Me
    Loop
    
    Application.ScreenUpdating = prevAppUpdating
End Sub

' ============================================================
'  Init メソッド
' ============================================================

Public Sub Init( _
        ByVal filePath As String, _
        Optional ByVal sheetName As String = "", _
        Optional ByVal headerRowIndex As Long = 1, _
        Optional ByVal keyColumnIndex As Long = 1 _
    )
    
    If isInitialized Then
        Call Cleanup
    End If
    
    headerRow = headerRowIndex
    keyColumn = keyColumnIndex
    dataStartRow = headerRow + 1
    
    ' --- ファイルを開く ---
    Dim fileExtension As String
    fileExtension = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))
    
    Select Case fileExtension
        Case "csv", "txt", "tsv"
            Call OpenCsvAsText(filePath)
        Case Else
            ' Excel 形式
            Call OpenExcelFile(filePath, sheetName)
    End Select
    
    Call PostInit
End Sub

Public Sub InitFromSheet( _
        ByVal ws As Worksheet, _
        Optional ByVal headerRowIndex As Long = 1, _
        Optional ByVal keyColumnIndex As Long = 1 _
    )
    
    If isInitialized Then
        Call Cleanup
    End If
    
    headerRow = headerRowIndex
    keyColumn = keyColumnIndex
    dataStartRow = headerRow + 1
    
    Set targetSheet = ws
    Set targetWorkbook = ws.Parent
    ownsWorkbook = False
    
    Call PostInit
End Sub

Private Sub PostInit()
    Call BuildHeaderMap
    Call RecalculateDataRange
    Call MoveReset
    isInitialized = True
End Sub

' ============================================================
'  Cleanup / Terminate
' ============================================================

Public Sub Cleanup()
    Set deleteRange = Nothing
    Set headerMap = Nothing
    isAddBufferReady = False
    isInitialized = False
    
    If ownsWorkbook Then
        If Not targetWorkbook Is Nothing Then
            targetWorkbook.Close SaveChanges:=False
        End If
    End If
    
    Set targetSheet = Nothing
    Set targetWorkbook = Nothing
End Sub

Private Sub Class_Terminate()
    Call Cleanup
End Sub

' ============================================================
'  ヘッダー列番号取得
' ============================================================

Public Function GetColumnIndex(ByVal columnName As String, Optional ByVal occurrence As Long = 1) As Long
    Call EnsureInitialized
    
    Dim suffixedName As String
    suffixedName = columnName & "_" & occurrence
    
    On Error GoTo NotFound
    GetColumnIndex = headerMap(suffixedName)
    Exit Function
NotFound:
    Err.Raise 514, "DSWorksheet.GetColumnIndex", _
              "列名 '" & columnName & "' (出現 " & occurrence & ") が見つかりません。"
End Function

Public Function GetColumnName(ByVal colIndex As Long) As String
    Call EnsureInitialized
    
    If colIndex < 1 Or colIndex > columnCount Then
        Err.Raise 515, "DSWorksheet.GetColumnName", _
                  "列インデックス " & colIndex & " は範囲外です (1 ~ " & columnCount & ")。"
    End If
    GetColumnName = CStr(targetSheet.Cells(headerRow, colIndex).Value)
End Function

' ============================================================
'  Private ヘルパー
' ============================================================

Private Sub EnsureInitialized()
    If Not isInitialized Then
        Err.Raise 517, "DSWorksheet", "Init が呼ばれていません。"
    End If
End Sub

Private Sub EnsureCursorValid()
    Call EnsureInitialized
    If currentRowIndex < 1 Or currentRowIndex > RowCount Then
        Err.Raise 518, "DSWorksheet", "カーソルが有効な行を指していません (currentRow=" & currentRowIndex & ")。"
    End If
End Sub

Private Function ResolveColumnIndex(ByVal columnIdentifier As Variant, Optional ByVal occurrence As Long = 1) As Long
    If IsNumeric(columnIdentifier) Then
        ResolveColumnIndex = CLng(columnIdentifier)
    Else
        ResolveColumnIndex = GetColumnIndex(CStr(columnIdentifier), occurrence)
    End If
    
    ' 範囲チェック
    If ResolveColumnIndex < 1 Or ResolveColumnIndex > columnCount Then
        Err.Raise 516, "DSWorksheet.ResolveColumnIndex", _
                  "列 '" & CStr(columnIdentifier) & "' は有効範囲外です。"
    End If
End Function

Private Sub BuildHeaderMap()
    Set headerMap = New Collection
    
    ' ヘッダー行の最終列を検出
    columnCount = targetSheet.Cells(headerRow, targetSheet.Columns.Count).End(xlToLeft).Column
    
    Dim colIndex As Long
    Dim headerValue As String
    Dim occurrenceNumber As Long
    Dim suffixedName As String
    Dim testValue As Long
    
    For colIndex = 1 To columnCount
        headerValue = CStr(targetSheet.Cells(headerRow, colIndex).Value)
        If headerValue <> "" Then
            ' このヘッダー名の出現回数を検出（既存の _1, _2, ... をチェック）
            occurrenceNumber = 1
            Do
                suffixedName = headerValue & "_" & occurrenceNumber
                testValue = 0
                On Error Resume Next
                testValue = headerMap(suffixedName)
                On Error GoTo 0
                If testValue = 0 Then Exit Do
                occurrenceNumber = occurrenceNumber + 1
            Loop
            
            ' Collection に Key=列名_N, Item=列番号 で格納
            Call headerMap.Add(colIndex, suffixedName)
        End If
    Next colIndex
End Sub

Private Sub RecalculateDataRange()
    ' keyColumn の最終行を基準にデータ末尾を特定
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, keyColumn).End(xlUp).Row
    
    If lastRow <= headerRow Then
        ' データなし
        dataEndRow = headerRow  ' RowCount = 0 になる
    Else
        dataEndRow = lastRow
    End If
End Sub

' ============================================================
'  ファイルオープン
' ============================================================

Private Sub OpenExcelFile(ByVal filePath As String, ByVal sheetName As String)
    ' 既に開いているブックがあればそれを使用
    Dim fileName As String
    fileName = Dir(filePath)
    
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath, ReadOnly:=False)
        ownsWorkbook = True
    Else
        ownsWorkbook = False
    End If
    
    Set targetWorkbook = wb
    
    If sheetName <> "" Then
        Set targetSheet = targetWorkbook.Sheets(sheetName)
    Else
        Set targetSheet = targetWorkbook.ActiveSheet
    End If
End Sub

Private Sub OpenCsvAsText(ByVal filePath As String)
    ' CSV を新規ブックのシートに全列テキストとしてインポート
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add(xlWBATWorksheet)
    ownsWorkbook = True
    
    Set targetWorkbook = newWorkbook
    Set targetSheet = newWorkbook.Sheets(1)
    
    ' まず CSV の列数を推定するために1行目を読む
    Dim fileNumber As Integer
    Dim firstLine As String
    Dim delimiter As String
    
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    Line Input #fileNumber, firstLine
    Close #fileNumber
    
    ' 区切り文字の判定（タブ vs カンマ）
    If InStr(firstLine, vbTab) > 0 Then
        delimiter = vbTab
    Else
        delimiter = ","
    End If
    
    ' 列数を推定
    Dim estimatedColumns As Long
    Dim parts() As String
    parts = Split(firstLine, delimiter)
    estimatedColumns = UBound(parts) - LBound(parts) + 1
    
    ' 全列をテキスト（xlTextFormat = 2）にする FieldInfo 配列
    Dim fieldInfo() As Long
    ReDim fieldInfo(0 To estimatedColumns - 1)
    Dim i As Long
    For i = 0 To estimatedColumns - 1
        fieldInfo(i) = xlTextFormat
    Next i
    
    ' QueryTables でインポート
    Dim connectionString As String
    connectionString = "TEXT;" & filePath
    
    With targetSheet.QueryTables.Add( _
            Connection:=connectionString, _
            Destination:=targetSheet.Range("A1"))
        
        .TextFileParseType = xlDelimited
        
        If delimiter = vbTab Then
            .TextFileTabDelimiter = True
            .TextFileCommaDelimiter = False
        Else
            .TextFileTabDelimiter = False
            .TextFileCommaDelimiter = True
        End If
        
        .TextFileSemicolonDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileConsecutiveDelimiter = False
        .TextFileColumnDataTypes = fieldInfo
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
        
        ' QueryTable の接続情報を削除してクリーンに
        .Delete
    End With
End Sub
