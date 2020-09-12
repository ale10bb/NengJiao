Public Function Main()
    Call CommonWindow.WriteStatus("¡ª¡ª", "¡ª¡ª", "Ë¢ÐÂÍ¼±í±àºÅ", "¡ª¡ª")
    ActiveDocument.Range( _
        ActiveDocument.Sections(ActiveDocument.Sections.Count - 1).Range.start, _
        ActiveDocument.Content.End _
    ).Fields.Update
End Function
