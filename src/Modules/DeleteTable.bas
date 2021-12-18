Public Function PROMain()
    'Part 1: ��������ѡ����
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������ѡ����", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������������", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    'ɾ���ձ��Ӻ���ǰɾ
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��������ѡ����", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            ActiveDocument.Range( _
                workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
                workRange.Tables(i).Range.End _
                ).Delete
        End If
    Next i
    
    'Part 2: ��������������
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������������", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="�������С��", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    'ɾ���ձ��Ӻ���ǰɾ
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��������������", currentStr)
        If workRange.Tables(i).Columns.Count = 3 And Mid(workRange.Tables(i).Cell(1, 1).Range.Text, 1, 2) = "���" Then
            ActiveDocument.Range( _
                workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious, 3).Start, _
                workRange.Tables(i).Range.End _
                ).Delete
        End If
    Next i
    
    'Part 3: ��¼A
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��Ŀ�漰��Ϣ�ʲ�", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="������������¼", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    'ɾ���ձ��Ӻ���ǰɾ
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��Ŀ�漰��Ϣ�ʲ�", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            If ActiveDocument.Range( _
                workRange.Tables(i).Range.End, _
                workRange.Tables(i).Range.End + 2 _
                ).Text = "ע��" Then
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
    'Part 1: ��������ѡ����
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������ѡ����", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������������", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '�ձ�����������Ӻ���ǰ��
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��������ѡ����", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '����С��ϲ������õ�ɫΪ��ɫ���������֡�ȡ���Ӵ�
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "�����治�漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i

    'Part 2: ��¼A
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������ʲ�", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="�ϴβ��������������˵��", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    'ɾ���ձ��Ӻ���ǰɾ
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��������ʲ�", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "�����治�漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Function

Public Function ManualMain()

    Set workRange = Selection.Range
    '�ձ�����������Ӻ���ǰ��
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus("Manual", "Manual", "�Զ�����", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '����С��ϲ������õ�ɫΪ��ɫ���������֡�ȡ���Ӵ�
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "�����治�漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
        If workRange.Tables(i).Rows.Count = 2 And _
            Len(workRange.Tables(i).Cell(2, 1).Range.Text) = 2 Then
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "�����治�漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Function


Public Function PlanMain()
    'Part 1: ϵͳ����
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="ϵͳ����", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="ǰ�β��������������˵��", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '�ձ�����������Ӻ���ǰ��
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "ϵͳ����", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '����С��ϲ������õ�ɫΪ��ɫ���������֡�ȡ���Ӵ�
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "���������漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i

    'Part 2: ��������ѡ����
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��������ѡ����", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="�����ص�", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    '�ձ�����������Ӻ���ǰ��
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToHeading, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��������ѡ����", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            '����С��ϲ������õ�ɫΪ��ɫ���������֡�ȡ���Ӵ�
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "���������漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i

    'Part 3: ��չ��ȫҪ��
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="��չ��ȫҪ��", Forward:=False)
    Call headingRange.Collapse
    rangeStart = headingRange.Start
    startStr = ActiveDocument.Range(rangeStart, rangeStart + 10).Text & "..."
    
    Set headingRange = ActiveDocument.Content
    Call headingRange.Find.Execute(FindText:="�������", Forward:=False)
    Call headingRange.Collapse
    rangeEnd = headingRange.Start
    endStr = ActiveDocument.Range(rangeEnd, rangeEnd + 10).Text & "..."
    
    Set workRange = ActiveDocument.Range(rangeStart, rangeEnd)
    'ɾ���ձ��Ӻ���ǰɾ
    For i = workRange.Tables.Count To 1 Step -1
        currentStr = ActiveDocument.Range( _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start, _
            workRange.Tables(i).Range.GoTo(wdGoToLine, wdGoToPrevious).Start + 10 _
            ).Text & "..."
        Call CommonWindow.WriteStatus(startStr, endStr, "��չ��ȫҪ��", currentStr)
        If workRange.Tables(i).Rows.Count = 1 Then
            Call workRange.Tables(i).Rows.Add
            Call workRange.Tables(i).Rows(2).Cells.Merge
            workRange.Tables(i).Cell(2, 1).Shading.Texture = wdTextureNone
            workRange.Tables(i).Cell(2, 1).Shading.ForegroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Shading.BackgroundPatternColor = wdColorAutomatic
            workRange.Tables(i).Cell(2, 1).Range.Text = "���������漰"
            workRange.Tables(i).Cell(2, 1).Range.Font.Bold = False
            workRange.Tables(i).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Function

