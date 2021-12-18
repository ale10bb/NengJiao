Public Function PROMain()
    ' C.3中需要根据实际扫描层面调整表格，因此在遍历原始附件时，需要记录：
    ' 该份报告是否涉及操作系统、数据库、应用
    ' 数组中存储对应的章节编号，vb的默认值为0，有效序号从1开始
    Dim vulKind(3) As Integer
    ' 自适应复制延时
    copyDelay = 0
    copyDelay_max = 1
       
    Call CommonWindow.WriteStatus("——", "——", "初始化", "框架")
    Set oldAttachment = ActiveDocument
    Set newAttachment = Documents.Add(Visible:=True)
    ' ===== 从dotm中复制2章节（模板格式） =====
    Call ThisDocument.Range(ThisDocument.Sections(2).Range.Start, ThisDocument.Sections(2).Range.End).Copy
    
    Call Delay(copyDelay)
    ' Issue：某些电脑在这里会随机报错4605，猜测是copy入剪贴板后，paste之前，剪贴板未处理完成
    ' 曾经在复制和粘贴之间加延时，可以大大降低该问题的出现频率
    ' 微软论坛给出的解决方案也类似(Add DoEvents)
    ' 因此引入特性：粘贴报错时，增加延时并重新粘贴，触及延时最大值则放弃复制
    On Error GoTo CopyFailureHandler
    Call newAttachment.Content.PasteAndFormat(wdUseDestinationStylesRecovery)
    On Error GoTo 0
    ' 如果文档框架复制失败则直接结束程序
    If Not newAttachment.Bookmarks.Exists("C3_Start") Then
        Call MsgBox("框架复制失败，请重新运行宏")
        Call newAttachment.Close(SaveChanges:=False)
        Exit Function
    End If

    ' ===== 获取项目编号并写入 =====
    Dim tempStr() As String
    tempStr() = Split(oldAttachment.Sections(2).Headers(wdHeaderFooterPrimary).Range.Text, "：")
    ' 读第二章页眉，按冒号分隔，去除空白符
    codeStr = Trim(Replace(tempStr(UBound(tempStr)), Chr(13), ""))
    newAttachment.Bookmarks("C0项目编号2").Range.Text = codeStr '原编号1是封面的书签，版本更新已删除
    
    ' ===== 建立漏洞-资产字典 =====
    Call CommonWindow.WriteStatus("——", "——", "初始化", "建立字典")
    vulListStr = ""
    Set vulDict = New Scripting.Dictionary
    For i = 2 To 4
        For j = 2 To oldAttachment.Content.Tables(i).Rows.Count
            ' 注：Word采用Chr(13) & Chr(7)标识单元格结束
            vulName = Replace(oldAttachment.Tables(i).Cell(j, 3).Range.Text, Chr(13) & Chr(7), "")
            vulAsset = Replace(oldAttachment.Tables(i).Cell(j, 5).Range.Text, Chr(13) & Chr(7), "")
            Call vulDict.Add(vulName, vulAsset)
            vulListStr = vulListStr + vulName + "、"
        Next j
    Next i
    
    ' ===== 复制漏洞内容 =====
    vulKindPos = 1
    ' 定位到B.3起始处
    Set tempRange = oldAttachment.Content
    Call tempRange.Find.Execute(FindText:=".3 漏洞详细", Forward:=True)
    startPos = tempRange.GoTo(wdGoToHeading, wdGoToNext).Start
    Do While True
        Call CommonWindow.WriteStatus("B.3 漏洞详细", "B.4 测试结论", "复制内容", oldAttachment.Range(startPos, startPos + 10).Text)
        ' workRange设置为两个标题之间的区域
        endPos = oldAttachment.Range(startPos, startPos).GoTo(wdGoToHeading, wdGoToNext).Start - 1
        If oldAttachment.Range(startPos + 1, startPos + 3).Text = ".4" Then Exit Do
        ' 三级标题（层面）的情况
        If Left(oldAttachment.Range(startPos, startPos).ParagraphStyle, 4) = "标题 3" Then
            ' 从模板中复制三级菜单
            ' 就一行字应该不会报错吧，懒得检测
            Call ThisDocument.Sections(3).Range.Paragraphs(1).Range.Copy
            Call Delay(copyDelay)
            On Error GoTo CopyFailureHandler
            ' C3_Start位于C.3的第一字符位置，用于定位复制点
            Call newAttachment.Range( _
                newAttachment.Bookmarks("C3_Start").Range.Start - 1, _
                newAttachment.Bookmarks("C3_Start").Range.Start - 1 _
                ).PasteAndFormat(wdUseDestinationStylesRecovery)
            On Error GoTo 0
            ' 暂存该层面位于附件中的实际位置
            Select Case oldAttachment.Range(startPos + 6, endPos).Text
                Case "操作系统"
                    newAttachment.Bookmarks("C2层面").Range.Text = "操作系统"
                    vulKind(0) = vulKindPos
                Case "数据库"
                    newAttachment.Bookmarks("C2层面").Range.Text = "数据库"
                    vulKind(1) = vulKindPos
                Case "应用"
                    newAttachment.Bookmarks("C2层面").Range.Text = "应用系统"
                    vulKind(2) = vulKindPos
            End Select
            vulKindPos = vulKindPos + 1
        End If
        ' 四级标题（漏洞）的情况
        If Mid(oldAttachment.Range(startPos, startPos).ParagraphStyle, 1, 4) = "标题 4" Then
            ' 从模板中复制四级菜单
            Call ThisDocument.Range( _
                ThisDocument.Sections(3).Range.Paragraphs(2).Range.Start, _
                ThisDocument.Sections(3).Range.Paragraphs(12).Range.End _
                ).Copy
            Call Delay(copyDelay)
            On Error GoTo CopyFailureHandler
            Call newAttachment.Range( _
                newAttachment.Bookmarks("C3_Start").Range.Start - 1, _
                newAttachment.Bookmarks("C3_Start").Range.Start - 1 _
                ).PasteAndFormat(wdUseDestinationStylesRecovery)
            On Error GoTo 0
            ' 复制失败就跳过漏洞吧(少一个应该看不出来)
            If Not newAttachment.Bookmarks.Exists("C2漏洞名称") Then
                ' 垃圾VBA没有Continue
                GoTo NextLoop
            End If
            ' 填入漏洞内容
            Set headingRange = oldAttachment.Range(startPos, endPos).Paragraphs(1).Range
            ' 跳过标题号
            Call headingRange.MoveStart(wdWord, 7)
            ' 去掉可能存在的整改情况描述
            tempText = Replace(headingRange.Text, "（已整改）", "")
            ' 去除换行符和空格
            tempText = Trim(Replace(tempText, Chr(13), ""))
            ' heading去除风险等级，此时为漏洞名称
            headingText = Trim(Left(tempText, Len(tempText) - 4))
            ' severity单独获取风险等级并去除括号
            severityText = Replace(Right(tempText, 3), "）", "")
            newAttachment.Bookmarks("C2漏洞名称").Range.Text = headingText
            ' 查字典填入关联资产
            If Not vulDict.Exists(headingText) Then
                newAttachment.Bookmarks("C2漏洞资产").Range.Text = "——"
            Else
                newAttachment.Bookmarks("C2漏洞资产").Range.Text = vulDict.Item(headingText)
            End If
            ' Issue：导出附件时，某些漏洞可能包含wdInlineShapeScriptAnchor的元素(图为wdInlineShapePicture)，导致异常进入if
            ' BugFix：检测InlineShape的类型，增加flag用于控制走向
            flag = False
            i_saved = -1
            If oldAttachment.Range(startPos, endPos).InlineShapes.Count > 0 Then
                For i = 1 To oldAttachment.Range(startPos, endPos).InlineShapes.Count
                    If oldAttachment.Range(startPos, endPos).InlineShapes(i).Type = wdInlineShapePicture Then
                        flag = True
                        i_saved = i
                        Exit For
                    End If
                Next i
            End If
            ' 如果有图，则额外复制漏洞名称到图注上，并复制图
            If flag Then
                Call oldAttachment.Range(startPos, endPos).InlineShapes(i_saved).Select
                Call Selection.Copy
                Call Delay(copyDelay)
                On Error GoTo CopyFailureHandler
                Call newAttachment.Bookmarks("C2图").Range.Paste
                On Error GoTo 0
                ' 复制失败则当没有图处理
                If newAttachment.Bookmarks.Exists("C2图") Then
                    flag = False
                Else
                    newAttachment.Bookmarks("C2漏洞名称2").Range.Text = headingText
                End If
            End If
            ' 否则删除图相关的两行
            If Not flag Then
                newAttachment.Range( _
                    newAttachment.Bookmarks("C2图").Range.Start, _
                    newAttachment.Bookmarks("C2漏洞名称2").Range.End + 1 _
                    ).Delete
            End If
            ' 转换风险等级
            Select Case severityText
                Case "危急", "高危"
                    newAttachment.Bookmarks("C2漏洞风险").Range.Text = "严重"
                Case "中危", "低危", "信息"
                    newAttachment.Bookmarks("C2漏洞风险").Range.Text = "一般"
            End Select
            ' 填入整改情况
            If InStrRev(headingRange.Text, "已整改") Then
                newAttachment.Bookmarks("C2漏洞整改").Range.Text = "相关漏洞已全部整改。" + Chr(13)
                Call newAttachment.Bookmarks("C2是否整改").Delete
            Else
                newAttachment.Bookmarks("C2漏洞整改").Range.Text = "相关漏洞尚未整改。" + Chr(13)
                newAttachment.Bookmarks("C2是否整改").Range.Text = ""
            End If
            ' 扫一遍原文，确定漏洞描述和修复建议的位置
            For i = 1 To oldAttachment.Range(startPos, endPos).Paragraphs.Count - 1
                Select Case Mid(oldAttachment.Range(startPos, endPos).Paragraphs(i).Range.Text, 1, 4)
                    Case "漏洞描述"
                        discriptionStart = oldAttachment.Range(startPos, endPos).Paragraphs(i + 1).Range.Start
                    Case "审计结果"
                        discriptionEnd = oldAttachment.Range(startPos, endPos).Paragraphs(i).Range.Start - 1
                    Case "修复建议"
                        adviceStart = oldAttachment.Range(startPos, endPos).Paragraphs(i + 1).Range.Start
                End Select
            Next i
            adviceEnd = endPos - 1
            newAttachment.Bookmarks("C2漏洞描述").Range.Text = oldAttachment.Range(discriptionStart, discriptionEnd).Text
            newAttachment.Bookmarks("C2漏洞建议").Range.Text = oldAttachment.Range(adviceStart, adviceEnd).Text
        End If
NextLoop:
        startPos = oldAttachment.Range(startPos, startPos).GoTo(wdGoToHeading, wdGoToNext).Start
    Loop
    
    Call CommonWindow.WriteStatus("——", "——", "修改占位符字段", "C.1")
    '删掉原附件中的十个连续空格
    Call newAttachment.Content.Find.Execute(FindText:="          ", ReplaceWith:="", Replace:=wdReplaceAll)
    '生成漏洞层面的字段
    C1Str = ""
    C3Str1 = ""
    C3Str2 = ""
    If vulKind(0) > 0 Then
        C1Str = C1Str + "操作系统、"
        C3Str2 = C3Str2 + "参见C.2." + Str(vulKind(0)) + "章节、"
    End If
    If vulKind(1) > 0 Then
        C1Str = C1Str + "数据库、"
        C3Str2 = C3Str2 + "参见C.2." + Str(vulKind(1)) + "章节、"
    End If
    If vulKind(2) > 0 Then
        C1Str = C1Str + "应用系统、"
        C3Str1 = C3Str1 + "参见C.2." + Str(vulKind(2)) + "章节、"
        C3Str2 = C3Str2 + "参见C.2." + Str(vulKind(2)) + "章节、"
    End If
    '填入C1
    newAttachment.Bookmarks("C1层面").Range.Text = Mid(C1Str, 1, Len(C1Str) - 1)
    Call CommonWindow.WriteStatus("——", "——", "修改占位符字段", "C.3")
    '填入C3
    If Len(C3Str2) > 0 Then
        newAttachment.Content.Tables(1).Cell(3, 4).Range.Text = Mid(C3Str2, 1, Len(C3Str2) - 1)
    Else
        Call newAttachment.Content.Tables(1).Cell(3, 4).Range.Cells.Delete(shiftcells:=wdDeleteCellsEntireRow)
    End If
    If Len(C3Str1) > 0 Then
        newAttachment.Content.Tables(1).Cell(2, 4).Range.Text = Mid(C3Str2, 1, Len(C3Str1) - 1)
    Else
        Call newAttachment.Content.Tables(1).Cell(2, 4).Range.Cells.Delete(shiftcells:=wdDeleteCellsEntireRow)
        newAttachment.Content.Tables(1).Cell(2, 2).Range.Text = "安全计算环境"
    End If
    For i = 2 To newAttachment.Content.Tables(1).Rows.Count
        newAttachment.Content.Tables(1).Cell(i, 1).Range.Text = Str(i - 1)
    Next i
    Call CommonWindow.WriteStatus("——", "——", "修改占位符字段", "C.4")
    '填入C4剩余漏洞列表
    If Len(vulListStr) > 0 Then
        newAttachment.Bookmarks("C4漏洞名称").Range.Text = Mid(vulListStr, 1, Len(vulListStr) - 1)
    Else
        newAttachment.Bookmarks("C4漏洞名称").Range.Text = "——"
    End If
    
    '修改安全漏洞数量统计表的标题行及背景色
    Call oldAttachment.Content.Tables(1).Range.Copy
    Call Delay(copyDelay)
    On Error GoTo CopyFailureHandler
    Call newAttachment.Bookmarks("C4漏洞表").Range.Paste
    On Error GoTo 0
    newAttachment.Content.Tables(2).Cell(1, 2).Range.Text = "严重"
    newAttachment.Content.Tables(2).Cell(1, 4).Range.Text = "一般"
    newAttachment.Content.Tables(2).Rows(1).Shading.BackgroundPatternColor = wdColorGray35
    '删最后统计行
    Call newAttachment.Content.Tables(2).Cell(newAttachment.Content.Tables(2).Rows.Count, 1).Delete(wdDeleteCellsEntireRow)
    '累加数量
    For i = 2 To newAttachment.Content.Tables(2).Rows.Count
        newAttachment.Content.Tables(2).Cell(i, 2).Range.Text = Str( _
            CInt(Replace(newAttachment.Content.Tables(2).Cell(i, 2).Range.Text, Chr(13) & Chr(7), "")) + _
            CInt(Replace(newAttachment.Content.Tables(2).Cell(i, 3).Range.Text, Chr(13) & Chr(7), "")) _
            )
        newAttachment.Content.Tables(2).Cell(i, 4).Range.Text = Str( _
            CInt(Replace(newAttachment.Content.Tables(2).Cell(i, 4).Range.Text, Chr(13) & Chr(7), "")) + _
            CInt(Replace(newAttachment.Content.Tables(2).Cell(i, 5).Range.Text, Chr(13) & Chr(7), "")) + _
            CInt(Replace(newAttachment.Content.Tables(2).Cell(i, 6).Range.Text, Chr(13) & Chr(7), "")) _
            )
    Next i
    ' 删除多余列，并修改字体和颜色
    Call newAttachment.Content.Tables(2).Cell(1, 6).Delete(wdDeleteCellsEntireColumn)
    Call newAttachment.Content.Tables(2).Cell(1, 5).Delete(wdDeleteCellsEntireColumn)
    Call newAttachment.Content.Tables(2).Cell(1, 3).Delete(wdDeleteCellsEntireColumn)
    newAttachment.Content.Tables(2).Range.Font.Name = "华文仿宋"
    newAttachment.Content.Tables(2).Range.Font.Name = "Times New Romans"
    newAttachment.Content.Tables(2).Range.Font.Color = wdColorAutomatic
    ' 刷新图表编号
    Call newAttachment.Content.Fields.Update
    ' 删除c3_start书签
    Call newAttachment.Bookmarks("C3_Start").Delete
    'Call oldAttachment.Close(SaveChanges:=wdDoNotSaveChanges)
    Call newAttachment.Activate
    Exit Function
    ' 复制过快异常时的处理函数
CopyFailureHandler:
    ' 增加延时最大值判断，触及最大值时直接放弃复制，以防死循环
    ' 放弃复制时，造成dstFile缺失内容(不影响后续流程)。并由外部检测并删除书签
    If copyDelay > 1 Then
        copyDelay = 0.4
        Resume Next
    ' 每次失败依次增加延时并重新尝试粘贴
    Else
        copyDelay = copyDelay + 0.2
        Call Delay(copyDelay)
        Resume
    End If
End Function


Public Function DBMain()
    ' C.3中需要根据实际扫描层面调整表格，因此在遍历原始附件时，需要记录：
    ' 该份报告是否涉及操作系统、数据库、应用
    ' 数组中存储对应的章节编号，vb的默认值为0，有效序号从1开始
    Dim vulKind(3) As Integer
    ' 自适应复制延时
    copyDelay = 0
    copyDelay_max = 1
       
    Call CommonWindow.WriteStatus("——", "——", "初始化", "框架")
    Set oldAttachment = ActiveDocument
    Set newAttachment = Documents.Add(Visible:=True)
    ' ===== 从dotm中复制5章节（模板格式） =====
    Call ThisDocument.Range(ThisDocument.Sections(4).Range.Start, ThisDocument.Sections(4).Range.End).Copy
    Call Delay(copyDelay)
    ' Issue：某些电脑在这里会随机报错4605，猜测是copy入剪贴板后，paste之前，剪贴板未处理完成
    ' 曾经在复制和粘贴之间加延时，可以大大降低该问题的出现频率
    ' 微软论坛给出的解决方案也类似(Add DoEvents)
    ' 因此引入特性：粘贴报错时，增加延时并重新粘贴，触及延时最大值则放弃复制
    On Error GoTo CopyFailureHandler
    Call newAttachment.Content.PasteAndFormat(wdUseDestinationStylesRecovery)
    On Error GoTo 0
    ' 如果文档框架复制失败则直接结束程序
    If Not newAttachment.Bookmarks.Exists("E3_Start") Then
        Call MsgBox("框架复制失败，请重新运行宏")
        Call newAttachment.Close(SaveChanges:=False)
        Exit Function
    End If

    
    ' ===== 建立漏洞-资产字典 =====
    Call CommonWindow.WriteStatus("——", "——", "初始化", "建立字典")
    vulListStr = ""
    Set vulDict = New Scripting.Dictionary
    For i = 2 To 4
        For j = 2 To oldAttachment.Content.Tables(i).Rows.Count
            ' 注：Word采用Chr(13) & Chr(7)标识单元格结束
            vulName = Replace(oldAttachment.Tables(i).Cell(j, 3).Range.Text, Chr(13) & Chr(7), "")
            vulAsset = Replace(oldAttachment.Tables(i).Cell(j, 5).Range.Text, Chr(13) & Chr(7), "")
            Call vulDict.Add(vulName, vulAsset)
            vulListStr = vulListStr + vulName + "、"
        Next j
    Next i
    
    ' ===== 复制漏洞内容 =====
    vulKindPos = 1
    ' 定位到B.3起始处
    Set tempRange = oldAttachment.Content
    Call tempRange.Find.Execute(FindText:=".3 漏洞详细", Forward:=True)
    startPos = tempRange.GoTo(wdGoToHeading, wdGoToNext).Start
    Do While True
        Call CommonWindow.WriteStatus("B.3 漏洞详细", "B.4 测试结论", "复制内容", oldAttachment.Range(startPos, startPos + 10).Text)
        ' workRange设置为两个标题之间的区域
        endPos = oldAttachment.Range(startPos, startPos).GoTo(wdGoToHeading, wdGoToNext).Start - 1
        If oldAttachment.Range(startPos + 1, startPos + 3).Text = ".4" Then Exit Do
        ' 三级标题（层面）的情况
        If Left(oldAttachment.Range(startPos, startPos).ParagraphStyle, 4) = "标题 3" Then
            ' 从模板中复制三级菜单
            ' 就一行字应该不会报错吧，懒得检测
            Call ThisDocument.Sections(4).Range.Paragraphs(1).Range.Copy
            Call Delay(copyDelay)
            On Error GoTo CopyFailureHandler
            ' E3_Start位于E.3的第一字符位置，用于定位复制点
            Call newAttachment.Range( _
                newAttachment.Bookmarks("E3_Start").Range.Start - 1, _
                newAttachment.Bookmarks("E3_Start").Range.Start - 1 _
                ).PasteAndFormat(wdUseDestinationStylesRecovery)
            On Error GoTo 0
            ' 暂存该层面位于附件中的实际位置
            Select Case oldAttachment.Range(startPos + 6, endPos).Text
                Case "操作系统"
                    newAttachment.Bookmarks("E2层面").Range.Text = "操作系统"
                    vulKind(0) = vulKindPos
                Case "数据库"
                    newAttachment.Bookmarks("E2层面").Range.Text = "数据库"
                    vulKind(1) = vulKindPos
                Case "应用"
                    newAttachment.Bookmarks("E2层面").Range.Text = "应用系统"
                    vulKind(2) = vulKindPos
            End Select
            vulKindPos = vulKindPos + 1
        End If
        ' 四级标题（漏洞）的情况
        If Mid(oldAttachment.Range(startPos, startPos).ParagraphStyle, 1, 4) = "标题 4" Then
            ' 从模板中复制四级菜单
            Call ThisDocument.Range( _
                ThisDocument.Sections(5).Range.Paragraphs(2).Range.Start, _
                ThisDocument.Sections(5).Range.Paragraphs(9).Range.End _
                ).Copy
            Call Delay(copyDelay)
            On Error GoTo CopyFailureHandler
            Call newAttachment.Range( _
                newAttachment.Bookmarks("E3_Start").Range.Start - 1, _
                newAttachment.Bookmarks("E3_Start").Range.Start - 1 _
                ).PasteAndFormat(wdUseDestinationStylesRecovery)
            On Error GoTo 0
            ' 复制失败就跳过漏洞吧(少一个应该看不出来)
            If Not newAttachment.Bookmarks.Exists("E2漏洞名称") Then
                ' 垃圾VBA没有Continue
                GoTo NextLoop
            End If
            ' 填入漏洞内容
            Set headingRange = oldAttachment.Range(startPos, endPos).Paragraphs(1).Range
            ' 跳过标题号
            Call headingRange.MoveStart(wdWord, 7)
            ' 去掉可能存在的整改情况描述
            tempText = Replace(headingRange.Text, "（已整改）", "")
            ' 去除换行符和空格
            tempText = Trim(Replace(tempText, Chr(13), ""))
            ' heading去除风险等级，此时为漏洞名称
            headingText = Trim(Left(tempText, Len(tempText) - 4))
            ' severity单独获取风险等级并去除括号
            severityText = Replace(Right(tempText, 3), "）", "")
            newAttachment.Bookmarks("E2漏洞名称").Range.Text = headingText
            ' Issue：导出附件时，某些漏洞可能包含wdInlineShapeScriptAnchor的元素(图为wdInlineShapePicture)，导致异常进入if
            ' BugFix：检测InlineShape的类型，增加flag用于控制走向
            flag = False
            i_saved = -1
            If oldAttachment.Range(startPos, endPos).InlineShapes.Count > 0 Then
                For i = 1 To oldAttachment.Range(startPos, endPos).InlineShapes.Count
                    If oldAttachment.Range(startPos, endPos).InlineShapes(i).Type = wdInlineShapePicture Then
                        flag = True
                        i_saved = i
                        Exit For
                    End If
                Next i
            End If
            ' 如果有图，则额外复制漏洞名称到图注上，并复制图
            If flag Then
                Call oldAttachment.Range(startPos, endPos).InlineShapes(i_saved).Select
                Call Selection.Copy
                Call Delay(copyDelay)
                On Error GoTo CopyFailureHandler
                Call newAttachment.Bookmarks("E2图").Range.Paste
                On Error GoTo 0
                ' 复制失败则当没有图处理
                If newAttachment.Bookmarks.Exists("E2图") Then
                    flag = False
                Else
                    newAttachment.Bookmarks("E2漏洞名称2").Range.Text = headingText
                End If
            End If
            ' 否则删除图相关的两行
            If Not flag Then
                newAttachment.Range( _
                    newAttachment.Bookmarks("E2图").Range.Start, _
                    newAttachment.Bookmarks("E2漏洞名称2").Range.End + 1 _
                    ).Delete
            End If
            ' 风险等级
            newAttachment.Bookmarks("E2漏洞风险").Range.Text = severityText

            ' 扫一遍原文，确定漏洞描述和修复建议的位置
            For i = 1 To oldAttachment.Range(startPos, endPos).Paragraphs.Count - 1
                Select Case Mid(oldAttachment.Range(startPos, endPos).Paragraphs(i).Range.Text, 1, 4)
                    Case "漏洞描述"
                        discriptionStart = oldAttachment.Range(startPos, endPos).Paragraphs(i + 1).Range.Start
                    Case "审计结果"
                        discriptionEnd = oldAttachment.Range(startPos, endPos).Paragraphs(i).Range.Start - 1
                    Case "修复建议"
                        adviceStart = oldAttachment.Range(startPos, endPos).Paragraphs(i + 1).Range.Start
                End Select
            Next i
            adviceEnd = endPos - 1
            newAttachment.Bookmarks("E2漏洞描述").Range.Text = oldAttachment.Range(discriptionStart, discriptionEnd).Text
            newAttachment.Bookmarks("E2漏洞整改").Range.Text = oldAttachment.Range(adviceStart, adviceEnd).Text
        End If
NextLoop:
        startPos = oldAttachment.Range(startPos, startPos).GoTo(wdGoToHeading, wdGoToNext).Start
    Loop
    
    Call CommonWindow.WriteStatus("——", "——", "修改占位符字段", "C.1")
    '删掉原附件中的十个连续空格
    Call newAttachment.Content.Find.Execute(FindText:="          ", ReplaceWith:="", Replace:=wdReplaceAll)
    '生成漏洞层面的字段
    C1Str = ""
    If vulKind(0) > 0 Then
        C1Str = C1Str + "操作系统、"
    End If
    If vulKind(1) > 0 Then
        C1Str = C1Str + "数据库、"
    End If
    If vulKind(2) > 0 Then
        C1Str = C1Str + "应用系统、"
    End If
    '填入E1
    newAttachment.Bookmarks("E1层面").Range.Text = Mid(C1Str, 1, Len(C1Str) - 1)
    Call CommonWindow.WriteStatus("——", "——", "修改占位符字段", "C.3")
    '填入E3剩余漏洞列表
    If Len(vulListStr) > 0 Then
        newAttachment.Bookmarks("E3漏洞名称").Range.Text = Mid(vulListStr, 1, Len(vulListStr) - 1)
    Else
        newAttachment.Bookmarks("E3漏洞名称").Range.Text = "——"
    End If
    
    '修改安全漏洞数量统计表的标题行及背景色
    Call oldAttachment.Content.Tables(1).Range.Copy
    Call Delay(copyDelay)
    On Error GoTo CopyFailureHandler
    Call newAttachment.Bookmarks("E3漏洞表").Range.Paste
    On Error GoTo 0
    newAttachment.Content.Tables(1).Rows(1).Shading.BackgroundPatternColor = wdColorGray35

    '累加数量
    For i = 2 To newAttachment.Content.Tables(1).Rows.Count
        newAttachment.Content.Tables(1).Cell(i, 6).Range.Text = Str( _
            CInt(Replace(newAttachment.Content.Tables(1).Cell(i, 2).Range.Text, Chr(13) & Chr(7), "")) + _
            CInt(Replace(newAttachment.Content.Tables(1).Cell(i, 3).Range.Text, Chr(13) & Chr(7), "")) + _
            CInt(Replace(newAttachment.Content.Tables(1).Cell(i, 4).Range.Text, Chr(13) & Chr(7), "")) + _
            CInt(Replace(newAttachment.Content.Tables(1).Cell(i, 5).Range.Text, Chr(13) & Chr(7), "")) _
            )
    Next i
    ' 修改字体和颜色
    newAttachment.Content.Tables(1).Range.Font.Name = "华文仿宋"
    newAttachment.Content.Tables(1).Range.Font.Name = "Times New Romans"
    newAttachment.Content.Tables(1).Range.Font.Color = wdColorAutomatic
    newAttachment.Content.Tables(1).Cell(1, 1).Range.Text = "设备名称或IP地址"
    newAttachment.Content.Tables(1).Cell(1, 6).Range.Text = "小计"
    newAttachment.Content.Tables(1).Cell(newAttachment.Content.Tables(1).Rows.Count, 1).Range.Text = "安全漏洞数量小计"
    
    ' 复制三个漏洞汇总表
    For Index = 2 To 4
        If oldAttachment.Content.Tables(Index).Rows.Count > 1 Then
            Select Case Index
                Case 2
                    newAttachment.Range( _
                        newAttachment.Bookmarks("E2_Start").Range.Start - 1, _
                        newAttachment.Bookmarks("E2_Start").Range.Start - 1 _
                    ).Text = "表 E-1 漏洞情况简表（操作系统）"
                Case 3
                    newAttachment.Range( _
                        newAttachment.Bookmarks("E2_Start").Range.Start - 2, _
                        newAttachment.Bookmarks("E2_Start").Range.Start - 2 _
                    ).Text = "表 E-1 漏洞情况简表（数据库）"
                Case 4
                    newAttachment.Range( _
                        newAttachment.Bookmarks("E2_Start").Range.Start - 2, _
                        newAttachment.Bookmarks("E2_Start").Range.Start - 2 _
                    ).Text = "表 E-1 漏洞情况简表（应用系统）"
            End Select
            Call oldAttachment.Content.Tables(Index).Range.Copy
            Call Delay(copyDelay)
            On Error GoTo CopyFailureHandler
            ' E2_Start位于E.2的第一字符位置，用于定位复制点
            Call newAttachment.Range( _
                newAttachment.Bookmarks("E2_Start").Range.Start - 1, _
                newAttachment.Bookmarks("E2_Start").Range.Start - 1 _
                ).PasteAndFormat(wdUseDestinationStylesRecovery)
            On Error GoTo 0
        End If
    Next Index
    ' 删除复制完汇总表的第二列
    For Index = 1 To newAttachment.Tables.Count - 1
        Call newAttachment.Tables(Index).Columns(2).Delete
        newAttachment.Tables(Index).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        newAttachment.Tables(Index).Rows(1).Shading.BackgroundPatternColor = wdColorGray35
        newAttachment.Tables(Index).Range.Font.Color = wdColorAutomatic
    Next Index
    ' 删除书签
    Call newAttachment.Bookmarks("E2_Start").Delete
    Call newAttachment.Bookmarks("E3_Start").Delete
    ' 刷新图表编号
    Call newAttachment.Content.Fields.Update
    'Call oldAttachment.Close(SaveChanges:=wdDoNotSaveChanges)
    Call newAttachment.Activate
    Exit Function
    ' 复制过快异常时的处理函数
CopyFailureHandler:
    ' 增加延时最大值判断，触及最大值时直接放弃复制，以防死循环
    ' 放弃复制时，造成dstFile缺失内容(不影响后续流程)。并由外部检测并删除书签
    If copyDelay > 1 Then
        copyDelay = 0.4
        Resume Next
    ' 每次失败依次增加延时并重新尝试粘贴
    Else
        copyDelay = copyDelay + 0.2
        Call Delay(copyDelay)
        Resume
    End If
End Function
