Public Sub PRO_处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.PROMain
    Unload CommonWindow

End Sub

Public Sub PRO_处理附件()
    Call CommonWindow.Show(False)
    Call Attachment.PROMain
    Unload CommonWindow

End Sub

Public Sub 等保_处理附件()
    Call CommonWindow.Show(False)
    Call Attachment.DBMain
    Unload CommonWindow

End Sub

'Public Sub 等保_处理空表()
'    Call CommonWindow.Show(False)
'    Call DeleteTable.DBMain
'    Unload CommonWindow
'End Sub

Public Sub 通用_自选区间处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.ManualMain
    Unload CommonWindow
End Sub

Public Sub 方案_处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.PlanMain
    Unload CommonWindow
End Sub

Sub 通用_刷编号()
    Call CommonWindow.Show(False)
    Call FieldUpdate.Main
    Unload CommonWindow
    Call MsgBox("记得刷目录和索引")
End Sub

'Public Sub PRO_改标题格式()
'    Call CommonWindow.Show(False)
'    Call TitleStyle.PROMain
'    Unload CommonWindow
'End Sub

'Public Sub 等保_改标题格式()
'    Call CommonWindow.Show(False)
'    Call TitleStyle.DBMain
'    Unload CommonWindow
'End Sub

Public Sub Delay(ByVal T As Single)
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents
    Loop While Timer - time1 < T
End Sub
