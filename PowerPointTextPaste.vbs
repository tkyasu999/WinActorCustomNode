'
' \par Copyright (C), 2023, superbunnbun
' @file    PowerPointTextPaste.vbs
' @author  superbunnbun
' @version V1.0.0
' @date    2023/06/16
' @brief   Description: PowerPointにおいて、コピーしたテキストを指定したスライド、位置へ貼り付ける.
'
num_slide = Cint( !スライド番号! )

Set objPpt = GetObject(, "PowerPoint.Application")  

If objPpt is Nothing then
  Set objPpt = CreateObject("PowerPoint.Application")  
End if

If objPpt is Nothing then
  Err.Raise 1, "", "指定されたPowerPointアプリケーションが開けません。"
End if

objPpt.Visible = True

Set targetPptObj = objPpt.ActivePresentation
Set ppSlide = targetPptObj.Slides(num_slide)
ppSlide.Shapes.Paste
ppSlide.Shapes(Cint(!Index!)).Top = Cint(!Top!)
ppSlide.Shapes(Cint(!Index!)).Left = Cint(!Left!)