Public Sub PRO_处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.PROMain
    Unload CommonWindow
    Call 检查更新
End Sub

Public Sub PRO_处理附件()
    Call CommonWindow.Show(False)
    Call Attachment.PROMain
    Unload CommonWindow
    Call 检查更新
End Sub

Public Sub 等保_处理附件()
    Call CommonWindow.Show(False)
    Call Attachment.DBMain
    Unload CommonWindow
    Call 检查更新
End Sub

Public Sub 等保_处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.DBMain
    Unload CommonWindow
    Call 检查更新
End Sub

Public Sub 通用_自选区间处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.ManualMain
    Unload CommonWindow
    Call 检查更新
End Sub

Public Sub 方案_处理空表()
    Call CommonWindow.Show(False)
    Call DeleteTable.PlanMain
    Unload CommonWindow
    Call 检查更新
End Sub

Sub 通用_刷编号()
    Call CommonWindow.Show(False)
    Call FieldUpdate.Main
    Unload CommonWindow
    Call MsgBox("记得刷目录和索引")
    Call 检查更新
End Sub

Public Sub Delay(ByVal T As Single)
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents
    Loop While Timer - time1 < T
End Sub

Private Sub 检查更新()
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
        response = MsgBox("发现新版本：v" & latestVersion(0) & "，是否下载", vbYesNo)
        If response = vbYes Then
            Call ThisDocument.FollowHyperlink("http://static.chenqlz.top/converter/%E6%B5%8B%E8%AF%84%E8%83%BD%E8%84%9A2.dotm")
        End If
    End If
End Sub
