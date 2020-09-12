Public Function PROMain()
    '搜索区间定位到附录B
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评结果记录", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    rangeEnd = ActiveDocument.Content.End
    endStr = "End"
    
    '循环修改附录B标题样式
    rangeIndexBefore = -1
    Do While True
        DoEvents
        rangeIndexAfter = headingRange.start
        If rangeIndexBefore >= rangeIndexAfter Then Exit Do
        rangeIndexBefore = headingRange.start
        currentStr = ActiveDocument.Range(rangeIndexAfter, rangeIndexAfter + 10).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "核查表标题样式", currentStr)
        '删除原有编号，并应用样式
        If Mid(headingRange.ParagraphStyle, 1, 4) = "标题 2" Then
            headingRange.Style = "附录二级标题"
            Call headingRange.MoveEnd(wdWord, 3)
            headingRange.Delete
        End If
        If Mid(headingRange.ParagraphStyle, 1, 4) = "标题 3" Then
            headingRange.Style = "附录三级标题"
            Call headingRange.MoveEnd(wdWord, 5)
            headingRange.Delete
        End If
        Set headingRange = headingRange.GoTo(wdGoToHeading, wdGoToNext)
    Loop

End Function

Public Function DBMain()
    '搜索区间定位到附录D
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="漏洞扫描结果记录", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."

    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="单项测评结果记录", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."

    '循环修改附录D标题样式
    rangeIndexBefore = -1
    Do While True
        DoEvents
        rangeIndexAfter = headingRange.start
        If rangeIndexBefore >= rangeIndexAfter Then Exit Do
        rangeIndexBefore = headingRange.start
        '删除原有编号，并应用样式
        currentStr = ActiveDocument.Range(rangeIndexAfter, rangeIndexAfter + 10).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "核查表标题样式", currentStr)
        If Mid(headingRange.ParagraphStyle, 1, 4) = "标题 2" Then
            headingRange.Style = "附录二级标题"
            Call headingRange.MoveEnd(wdWord, 3)
            headingRange.Delete
        End If
        If Mid(headingRange.ParagraphStyle, 1, 4) = "标题 3" Then
            headingRange.Style = "附录三级标题"
            Call headingRange.MoveEnd(wdWord, 3)
            headingRange.Delete
        End If
        Set headingRange = headingRange.GoTo(wdGoToHeading, wdGoToNext)
    Loop

End Function
