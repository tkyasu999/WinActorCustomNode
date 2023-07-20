'
' \par Copyright (C), 2023, superbunnbun
' @file    HalfToFull.vbs
' @author  tkyasu999
' @version V1.0.0
' @date    2023/07/20
' @brief   Description: 文字列の半角を全角へ変換する.
'
Const SHIFT_CODE = &H7DE1
strNarrow = !変換前!

Dim length, index, retStr, retChar, tempChar, tempCode
retStr = ""
StrConvWide = retStr

length = Len(strNarrow)
For index = 1 To length
    tempChar = Mid(strNarrow, index, 1)
    tempCode = Asc(tempChar)
    If (tempCode >= &H30 And tempCode <= &H39) Or _
        (tempCode >= &H41 And tempCode <= &H5A) Then
        retChar = Chr(tempCode - SHIFT_CODE)
    ElseIf tempCode >= &H61 And tempCode <= &H7A Then
        retChar = Chr(tempCode - SHIFT_CODE + 1)
    Else
        retChar = tempChar
    End If
    retStr = retStr & retChar
Next

SetUMSVariable $変換後$, retStr
