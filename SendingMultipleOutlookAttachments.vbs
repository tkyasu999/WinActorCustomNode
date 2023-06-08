'
' \par Copyright (C), 2023, superbunnbun
' @file    SendingMultipleOutlookAttachments.vbs
' @author  superbunnbun
' @version V1.0.0
' @date    2023/06/06
' @brief   Description: Outlookメールで, 複数ファイルを添付して送信する.
'
'---------------------------------------------------------------
'メイン
'---------------------------------------------------------------
Dim oApp
Dim myNameSpace
Dim myFolder
Dim mITEM 'As Outlook.MailItem
Dim sendTo
Dim sendCc
Dim attachmentFolder
Dim attachmentFile
Dim attachmentTitle
Dim fname
Dim absname
Dim folderPath

sendTo = !宛先（To）!
sendCc = !宛先（Cc）!
attachmentFolder =  !添付用のフォルダ指定!

'宛先（To）の入力がない場合
If sendTo = "" Then
    Err.Raise 1, "", "宛先（To）を指定して下さい"
    WScript.Quit()
End If

'Outlook起動
Set oApp = CreateObject("Outlook.Application")

'名前空間の指定
Set myNameSpace = oApp.GetNamespace("MAPI")

'作業フォルダーの指定と表示
Set myFolder = myNameSpace.GetDefaultFolder(6)
myFolder.Display

'通常サイズ olNormalWindow=2 で表示、xlMinimized=1
oApp.ActiveWindow.WindowState = 1

'メールアイテムの作成 olMailItem=0
Set mITEM = oApp.CreateItem(0)
    
'編集画面表示
'mITEM.Display   
    
'データのセット
mITEM.Subject = !件名!
mITEM.To = sendTo
If sendCc <> "" Then
    mITEM.Cc = sendCc
End If
mITEM.Body = !本文!

'一時保存
mITEM.Save

'添付ファイルを添付 olByValue=1
Set myAttachments = mITEM.Attachments

Dim fso, folder, file

' ファイルシステムオブジェクトの作成
Set fso = CreateObject("Scripting.FileSystemObject")

' 指定したパスの情報を取得
Set folder = fso.GetFolder(attachmentFolder)

' 指定したパスにあるファイルの数だけループ
For Each file in folder.files
    ' ファイル名の表示
    myAttachments.Add file.Path, 1, 1, file.Name
Next

'送信
mITEM.Send

'5秒スリープ
WScript.Sleep 5000

'Outlookを終了
oApp.Quit

