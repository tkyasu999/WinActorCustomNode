' ファイルのパスをフルパスに変換する
fname = !ファイル名!

SetUMSVariable "$FILE_PATH_TYPE", "1"
SetUMSVariable "$PARSE_FILE_PATH", fname
filePath = GetUMSVariable("$PARSE_FILE_PATH")
If filePath = "" Then
  Err.Raise 1, "", "指定されたファイルを開くことができません。" 
End If

' workbookオブジェクトを取得する
Set workbook = Nothing
On Error Resume Next
  ' 既存のエクセルが起動されていれば警告を抑制する
  Set existingXlsApp = Nothing
  Set existingXlsApp = GetObject(, "Excel.Application")
  existingXlsApp.DisplayAlerts = False

  Set wash = CreateObject("WinActor7.ScriptHelper")
  For Each book in wash.GetExcelWorkbooks
    SetUMSVariable "$FILE_PATH_TYPE", 0
    SetUMSVariable "$PARSE_FILE_PATH", book.FullName
    bookPath = GetUMSVariable("$PARSE_FILE_PATH")
    If StrComp(bookPath, filePath, 1) = 0 Then
      Set workbook = book
      Set xlsApp = workbook.Parent
      xlsApp.Visible = True
      Exit For
    End If
  Next
  Set wash = Nothing

  ' Workbookが存在しない場合は、新たに開く。
  If workbook Is Nothing Then
    Set xlsApp = Nothing

    ' Excelが既に開かれていたならそれを再利用する
    If Not existingXlsApp Is Nothing Then
      Set xlsApp = existingXlsApp
      xlsApp.Visible = True
    Else
      Set xlsApp = CreateObject("Excel.Application")
      xlsApp.Visible = True
    End If

    Set workbook = xlsApp.Workbooks.Open(filePath)
  End If

  ' 警告の抑制を元に戻す
  existingXlsApp.DisplayAlerts = True
  Set existingXlsApp = Nothing
  On Error Goto 0

If workbook Is Nothing Then
  Err.Raise 1, "", "指定されたファイルを開くことができません。"
End If

' ====指定されたシートを取得する==================================================

sheetName = !シート名!
Set worksheet = Nothing
On Error Resume Next
  ' シート名が指定されていない場合は、アクティブシートを対象とする
  If sheetName = "" Then
    Set worksheet = workbook.ActiveSheet
  Else
    Set worksheet = workbook.Worksheets(sheetName)
  End If
On Error Goto 0

If worksheet Is Nothing Then
  Err.Raise 1, "", "指定されたシートが見つかりません。"
End If

worksheet.Activate

' ====指定された範囲をクリップボードへコピー==================================================

scell = !開始セル!
ecell = !終了セル!
range = scell&":"&ecell

On Error Resume Next
cnt = worksheet.range(range).COUNT
On Error Goto 0

'指定された範囲が有効か確認
If cnt = "" Then
  Err.Raise 1, "", "指定された範囲が無効です。"
End If

worksheet.range(range).Select
worksheet.range(range).Merge
worksheet.range(range).Borders.LineStyle = 1
worksheet.range(range).Borders.Weight = -4138

' ====終了処理==================================================
Set xlsApp = Nothing
Set existingXlsApp = Nothing
Set ExcelWorkBook = Nothing
Set Excelworksheet = Nothing
Set sheetName = Nothing
Set filePath = Nothing