Public Sub PRO_����ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.PROMain
    Unload CommonWindow

End Sub

Public Sub PRO_������()
    Call CommonWindow.Show(False)
    Call Attachment.PROMain
    Unload CommonWindow

End Sub

Public Sub �ȱ�_������()
    Call CommonWindow.Show(False)
    Call Attachment.DBMain
    Unload CommonWindow

End Sub

'Public Sub �ȱ�_����ձ�()
'    Call CommonWindow.Show(False)
'    Call DeleteTable.DBMain
'    Unload CommonWindow
'End Sub

Public Sub ͨ��_��ѡ���䴦��ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.ManualMain
    Unload CommonWindow
End Sub

Public Sub ����_����ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.PlanMain
    Unload CommonWindow
End Sub

Sub ͨ��_ˢ���()
    Call CommonWindow.Show(False)
    Call FieldUpdate.Main
    Unload CommonWindow
    Call MsgBox("�ǵ�ˢĿ¼������")
End Sub

'Public Sub PRO_�ı����ʽ()
'    Call CommonWindow.Show(False)
'    Call TitleStyle.PROMain
'    Unload CommonWindow
'End Sub

'Public Sub �ȱ�_�ı����ʽ()
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
