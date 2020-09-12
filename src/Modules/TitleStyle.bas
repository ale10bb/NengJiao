Public Function PROMain()
    '�������䶨λ����¼B
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="������������¼", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    rangeEnd = ActiveDocument.Content.End
    endStr = "End"
    
    'ѭ���޸ĸ�¼B������ʽ
    rangeIndexBefore = -1
    Do While True
        DoEvents
        rangeIndexAfter = headingRange.start
        If rangeIndexBefore >= rangeIndexAfter Then Exit Do
        rangeIndexBefore = headingRange.start
        currentStr = ActiveDocument.Range(rangeIndexAfter, rangeIndexAfter + 10).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "�˲�������ʽ", currentStr)
        'ɾ��ԭ�б�ţ���Ӧ����ʽ
        If Mid(headingRange.ParagraphStyle, 1, 4) = "���� 2" Then
            headingRange.Style = "��¼��������"
            Call headingRange.MoveEnd(wdWord, 3)
            headingRange.Delete
        End If
        If Mid(headingRange.ParagraphStyle, 1, 4) = "���� 3" Then
            headingRange.Style = "��¼��������"
            Call headingRange.MoveEnd(wdWord, 5)
            headingRange.Delete
        End If
        Set headingRange = headingRange.GoTo(wdGoToHeading, wdGoToNext)
    Loop

End Function

Public Function DBMain()
    '�������䶨λ����¼D
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="©��ɨ������¼", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."

    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="������������¼", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."

    'ѭ���޸ĸ�¼D������ʽ
    rangeIndexBefore = -1
    Do While True
        DoEvents
        rangeIndexAfter = headingRange.start
        If rangeIndexBefore >= rangeIndexAfter Then Exit Do
        rangeIndexBefore = headingRange.start
        'ɾ��ԭ�б�ţ���Ӧ����ʽ
        currentStr = ActiveDocument.Range(rangeIndexAfter, rangeIndexAfter + 10).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "�˲�������ʽ", currentStr)
        If Mid(headingRange.ParagraphStyle, 1, 4) = "���� 2" Then
            headingRange.Style = "��¼��������"
            Call headingRange.MoveEnd(wdWord, 3)
            headingRange.Delete
        End If
        If Mid(headingRange.ParagraphStyle, 1, 4) = "���� 3" Then
            headingRange.Style = "��¼��������"
            Call headingRange.MoveEnd(wdWord, 3)
            headingRange.Delete
        End If
        Set headingRange = headingRange.GoTo(wdGoToHeading, wdGoToNext)
    Loop

End Function
