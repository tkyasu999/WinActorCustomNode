'
' \par Copyright (C), 2023, superbunnbun
' @file    PowerPointSlideSelect.vbs
' @author  superbunnbun
' @version V1.0.0
' @date    2023/06/17
' @brief   Description: PowerPointにおいて、指定した番号のスライドを選択する.
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
targetPptObj.Slides(num_slide).Select

