Public Function Main()
    Call CommonWindow.WriteStatus("����", "����", "ˢ��ͼ����", "����")
    ActiveDocument.Range( _
        ActiveDocument.Sections(ActiveDocument.Sections.Count - 1).Range.start, _
        ActiveDocument.Content.End _
    ).Fields.Update
End Function
