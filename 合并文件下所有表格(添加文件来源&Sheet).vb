
Sub 合并当前目录下所有工作簿的全部工作表()
    keyi = 0
    Do While keyi = 0
        a = InputBox("请选择要开始取的行数,可输入0结束")
        If a = "stop" Then
            MsgBox "自动处理已被停止"
            Exit Sub
        ElseIf a = 0 Then
            MsgBox "自动处理已结束"
            Application.Quit
            ThisWorkbook.Close False
            Exit Sub
        ElseIf IsNumeric(a) = False Or a < 0 Then
            MsgBox "请输入正确的数字"
        Else: keyi = 1
        End If
    Loop
    
    Dim MyPath, MyName, AWbName
    Dim Wb As Workbook, WbN As String
    Dim G As Long
    Dim Num As Long
    Dim BOX As String
    Application.ScreenUpdating = False
    MyPath = ActiveWorkbook.Path
    MyName = Dir(MyPath & "\" & "*.*")
    AWbName = ActiveWorkbook.Name
    Set twb = ThisWorkbook.Sheets(1)
    Num = 1
    num_count = 1
    
    ' 清理内容
    MsgBox "清除内容"
    Sheets(1).Select
    Cells.Select
    Selection.ClearContents
    
    Do While MyName <> ""
        If MyName <> AWbName Then
            a0 = twb.Range("C1048576").End(xlUp).Row
            Set Wb = Workbooks.Open(MyPath & "\" & MyName)
             ' Workbooks(1).ActiveSheet，选择触发宏的本工作簿的Sheet
             With Workbooks(1).ActiveSheet
                For G = 1 To Sheets.Count
                    If a - 1 > 0 Then
                        Wb.Sheets(G).Range(Rows(1), Rows(a - 1)).Delete
                    End If
                    a1 = twb.Range("C1048576").End(xlUp).Row
                     ' 向下偏移1行; ? 放入第3列
                     If num_count = 1 Then ' 第一次文件
                         Wb.Sheets(G).UsedRange.Offset(0, 0).Copy twb.Cells(twb.Range("C1048576").End(xlUp).Row, 3)
                     Else
                     ' 原数据有title的情况
                         Wb.Sheets(G).UsedRange.Offset(1, 0).Copy twb.Cells(twb.Range("C1048576").End(xlUp).Row + 1, 3)
                        ' 原数据无title的情况
                        ' Wb.Sheets(G).UsedRange.Offset(0, 0).Copy twb.Cells(twb.Range("C1048576").End(xlUp).Row + 1, 3)
                     End If
                     num_count = num_count + 1
        
                ' 添加sheet name
                    a2 = twb.Range("C1048576").End(xlUp).Row
                    a3 = Wb.Sheets(G).Name
                   
                    For i = a1 + 1 To a2
                    ' 没有添加“.”导致sheet name全是空的！.Cells才写入当下打开的Sheet里，Cells不写入！！！shit!!!
                        twb.Cells(i, 2) = a3
                    Next
                Next
        ' 第一个回车符
                WbN = WbN & Chr(13) & Wb.Name
                Wb.Close False
        
             End With
        
            ' 添加Excel name
            a4 = twb.Range("C1048576").End(xlUp).Row
            ' Find()找到目标str所在的位置，返回数值
            a5 = Application.Find(".", MyName)
            For i = a0 + 1 To a4
                twb.Cells(i, 1) = Left(MyName, a5 - 1)
            Next
    
        End If
    MyName = Dir
    Num = Num + 1
    Loop
    
    Cells(1, 1) = "来源excel"
    Cells(1, 2) = "来源sheet"
    
    
    MsgBox "共合并了" & (Num - 2) & "个工作薄下的全部工作表。如下：" & Chr(13) & WbN, vbInformation, "提示"
    Application.ScreenUpdating = True

End Sub
