'
' \par Copyright (C), 2023, tkyasu999
' @file    GetNumberSpecificCharacters.vbs
' @author  tkyasu999
' @version V1.0.0
' @date    2023/07/24
' @brief   Description: 対象文字列に対して、特定文字が含まれている数を取得する.
'
strNarrow = !対象文字列!
strSpe = !特定文字!

Dim i, cnt
cnt = 0
For i = 1 To Len(strNarrow)
    If Mid(strNarrow, i, 1) = strSpe Then
        cnt = cnt + 1
    End If
Next

SetUMSVariable $数$, cnt