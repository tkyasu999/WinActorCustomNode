'
' \par Copyright (C), 2023, tkyasu999
' @file    ExcelSearch.vbs
' @author  tkyasu999
' @version V1.0.0
' @date    2023/06/03
' @brief   Description: Excelにおいて、指定された範囲で一致したセルの行番号と列番号を取得する. その際に, 検索一致度に完全一致と部分一致を含む.
'
' 検索タイプを選択する。
searchType = !検索タイプ|文字列,日付!

LookAt = !一致度|完全一致,部分一致!
Select Case LookAt
  Case "完全一致"
    LookAt = 1
  Case "部分一致"
    LookAt = 2
End Select  

' ====指定されたファイルを開く====================================================

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

' ====指定された範囲を検索==================================================

keyword =!検索単語!
scell = !開始セル!
ecell = !終了セル!
rangeStr = scell&":"&ecell

'A1へ移動
worksheet.Activate
worksheet.Cells(1,1).Select

On Error Resume Next
cnt = worksheet.range(rangeStr).COUNT
On Error Goto 0

SetNull()
'指定された範囲が有効か確認
If cnt = "" Then
  Err.Raise 1, "", "指定された範囲が無効です。"
'指定された範囲が1の場合、単一セル検索を行う
ElseIf cnt = 1 Then
  SingleFind(scell)
Else
  FindRange()
End If

' ====終了処理==================================================
Set xlsApp = Nothing
Set existingXlsApp = Nothing
Set ExcelWorkBook = Nothing
Set Excelworksheet = Nothing
Set sheetName = Nothing
Set filePath = Nothing

' ====値設定関数==================================================
Sub SetValue(obj)
  row = obj.Row
  column = obj.Column
    
  worksheet.cells(row,column).Select
    
  Call SetUMSVariable($結果(行)$, row)
  Call SetUMSVariable($結果(列)$, column)
End Sub

' ====空値設定関数==================================================
Sub SetNull()
  Call SetUMSVariable($結果(行)$, "")
  Call SetUMSVariable($結果(列)$, "")
End Sub

' ====検索処理==================================================
Sub FindRange()
  Dim cstrkeyWord
  Select Case searchType
    '文字列検索の場合
    Case "文字列"
      set obj = worksheet.range(rangeStr).Find(keyword,worksheet.Range(ecell),-4163, LookAt)
      cstrkeyWord = keyword
  '日付検索の場合
    Case "日付"
    set obj = worksheet.range(rangeStr).Find(keyword,worksheet.Range(ecell),-4123, LookAt)
  cstrkeyWord = CStr(DateValue(keyword))
  End Select

  If obj Is Nothing Then
    SetNull()
  Else
    SetValue(obj)
  End If
End Sub

' ====単一セルの検索処理==================================================
Sub SingleFind(target)
  Dim cstrkeyWord
  Select Case searchType
    '文字列検索の場合
    Case "文字列"
      cstrkeyWord = keyword
    '日付検索の場合
    Case "日付"
      cstrkeyWord = CStr(DateValue(keyword))
  End Select
  Set cell = Nothing
  On Error Resume Next
    '単一のセル指定の為、開始セルを指定
    Set cell = worksheet.range(target)
  On Error Goto 0
  '値が取得できなかった場合、エラーとする
  If cell Is Nothing Then
    Err.Raise 1, , "値の取得に失敗しました。"
  End If
  '取得した値と検索値が完全一致かチェックする
  If CStr(cell.Value) = cstrkeyWord Then
    SetValue(cell)
  Else
    SetNull()
  End If
End Sub