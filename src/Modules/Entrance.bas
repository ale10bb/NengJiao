Public Sub PRO_����ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.PROMain
    Unload CommonWindow
    Call ������
End Sub

Public Sub PRO_������()
    Call CommonWindow.Show(False)
    Call Attachment.PROMain
    Unload CommonWindow
    Call ������
End Sub

Public Sub �ȱ�_������()
    Call CommonWindow.Show(False)
    Call Attachment.DBMain
    Unload CommonWindow
    Call ������
End Sub

Public Sub �ȱ�_����ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.DBMain
    Unload CommonWindow
    Call ������
End Sub

Public Sub ͨ��_��ѡ���䴦��ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.ManualMain
    Unload CommonWindow
    Call ������
End Sub

Public Sub ����_����ձ�()
    Call CommonWindow.Show(False)
    Call DeleteTable.PlanMain
    Unload CommonWindow
    Call ������
End Sub

Sub ͨ��_ˢ���()
    Call CommonWindow.Show(False)
    Call FieldUpdate.Main
    Unload CommonWindow
    Call MsgBox("�ǵ�ˢĿ¼������")
    Call ������
End Sub

Public Sub Delay(ByVal T As Single)
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents
    Loop While Timer - time1 < T
End Sub

Private Sub ������()
    currentVersionString = "1.0.2"
    currentVersionNumber = 102
    Dim latestVersion() As String
    
    Dim XMLHTTP As Object
    Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
    On Error Resume Next
    XMLHTTP.Open "GET", "http://static.chenqlz.top/converter/NengJiao.version", False
    XMLHTTP.send

    latestVersion() = Split(XMLHTTP.ResponseText, "/")
    If CInt(latestVersion(1)) > currentVersionNumber Then
        DoEvents
        response = MsgBox("�����°汾��v" & latestVersion(0) & "���Ƿ�����", vbYesNo)
        If response = vbYes Then
            Call ThisDocument.FollowHyperlink("http://static.chenqlz.top/converter/%E6%B5%8B%E8%AF%84%E8%83%BD%E8%84%9A2.dotm")
        End If
    End If
End Sub
