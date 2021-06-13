Sub 批量插入图片()

' 批量插入图片 宏
'
On Error Resume Next '容错处理
Application.ScreenUpdating = False '关闭屏幕更新 提升速度
Dim MR As Range
Dim pictype(1 To 3) As String
Dim i, j, k As Long

pictype(1) = ".jpg"
pictype(2) = ".jpeg"
pictype(3) = ".png"

For Each MR In Selection
    For j = 1 To UBound(pictype)
        If Not IsEmpty(MR) And Dir("E:\资料\商品图" & "\" & MR.Value & pictype(j)) <> "" Then
        MR.Select
        Set m = Cells(MR.Row, MR.Column + 1)
        ML = m.Left + 3
        MT = m.Top +3
        MW = m.Width -5
        MH = m.Height -5
        
       ' 选定图片填充区域
    ActiveSheet.Shapes.AddShape(msoShapeRectangle,ML, MT, MW, MH).Select
        
        ' 添加图片
        Selection.ShapeRange.Fill.UserPicture _
        "E:\资料\商品图" & "\" & MR.Value & pictype(j)
        
        ' 改变边框颜色
        With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        End With
        
        End If
    Next
Next


Set MR = Nothing
Application.ScreenUpdating = True '开启屏幕更新

End Sub