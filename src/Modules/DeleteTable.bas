Public Function PROMain()
    'Part 1: 测评对象选择结果
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="测评对象选择结果", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评结果汇总", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '删除空表，从后往前删
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "测评对象选择结果", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            ActiveDocument.Range( _
                workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
                workRange.Tables(i).Range.End _
                ).Delete
        End If
    Next i
    
    'Part 2: 单项测评结果汇总
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评结果汇总", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评小结", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '删除空表，从后往前删
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "单项测评结果汇总", currentStr)
        If workRange.Tables(i).Columns.Count = 3 And Mid(workRange.Tables(i).Cell(1, 1).Range.Text, 1, 2) = "序号" Then
            ActiveDocument.Range( _
                workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious, 3).Start, _
                workRange.Tables(i).Range.End _
                ).Delete
        End If
    Next i
    
    'Part 3: 附录A
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="项目涉及信息资产", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评结果记录", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '删除空表，从后往前删
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "项目涉及信息资产", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            If ActiveDocument.Range( _
                workRange.Tables(i).Range.End, _
                workRange.Tables(i).Range.End + 2 _
                ).Text = "注：" Then
                rangeEnd = workRange.Tables(i).Range.End + 19
            Else
                rangeEnd = workRange.Tables(i).Range.End
            End If
            ActiveDocument.Range( _
                workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
                rangeEnd _
                ).Delete
        End If
    Next i
End Function


Public Function DBMain()
    'Part 1: 测评对象选择结果
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="测评对象选择结果", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评结果分析", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '空表加上描述，从后往前加
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "测评对象选择结果", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '添加行、合并、设置底色为白色、加入文字、取消加粗
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本报告不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i

    'Part 2: 附录A
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="被测对象资产", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="上次测评问题整改情况说明", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '删除空表，从后往前删
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "被测对象资产", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本报告不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Function

Public Function ManualMain()

    Set workRange = Selection.Range
    '空表加上描述，从后往前加
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus("Manual", "Manual", "自定义表格", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '添加行、合并、设置底色为白色、加入文字、取消加粗
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本报告不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
        If workRange.Tables(i).Rows.Count = 2 And _
            Len(workRange.Tables(i).Cell(2, 1).Range.Text) = 2 Then
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本报告不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Function


Public Function PlanMain()
    'Part 1: 系统构成
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="系统构成", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="前次测评问题整改情况说明", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '空表加上描述，从后往前加
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "系统构成", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '添加行、合并、设置底色为白色、加入文字、取消加粗
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本方案不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i

    'Part 2: 测评对象选择结果
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="测评对象选择结果", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="测评重点", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '空表加上描述，从后往前加
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "测评对象选择结果", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '添加行、合并、设置底色为白色、加入文字、取消加粗
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本方案不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i

    'Part 3: 扩展安全要求
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="扩展安全要求", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="整体测评", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '删除空表，从后往前删
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "扩展安全要求", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "本方案不涉及"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Function

