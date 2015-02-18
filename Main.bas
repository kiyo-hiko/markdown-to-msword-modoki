''''
'''' Markdown to Wordコンバーター 試食品
'''' スタイルは予め文書に組み込まれているつもりとする。
'''' Author kiyo-hiko
'''' Since 2015. 2.17
''''
''' Markdownてきなテキストファイルを読んで書式付きWord文書に変換。
Sub ReadMDLikeText()
    Dim dlg: Set dlg = Application.FileDialog(msoFileDialogOpen)
    With dlg.Filters
        .Clear
        .Add "テキストファイル", "*.txt", 1
    End With
    If dlg.Show = -1 Then
        Dim f: For Each f In dlg.SelectedItems
            Open f For Input As #1
            Do Until EOF(1)
                Dim l: Line Input #1, l
                l = ApplyFormat(l)
                Selection.TypeText Text:=l & vbCrLf
            Loop
            Close #1
        Next f
    End If
End Sub

''' 行単位の書式設定はこれでやる。変換メソッドを肥大化させたくないし。
Function ApplyFormat(l)
    With Selection
        If False Then
            MsgBox 1
        ElseIf Left(l, 4) = "### " Then
            .Paragraphs.Style = ActiveDocument.Styles("見出し3")
            ApplyFormat = Mid(l, 5)
        ElseIf Left(l, 3) = "## " Then
            .Paragraphs.Style = ActiveDocument.Styles("見出し2")
            ApplyFormat = Mid(l, 4)
        ElseIf Left(l, 2) = "# " Then
            .Paragraphs.Style = ActiveDocument.Styles("見出し1")
            ApplyFormat = Mid(l, 3)
        ElseIf Left(l, 1) = ">" Then
            .Paragraphs.Style = ActiveDocument.Styles("引用")
            ApplyFormat = Mid(l, 2)
        ElseIf Left(l, 3) = "** " Then ' 番号なしリスト：レベル2
            .Range.SetListLevel Level:=2
            ApplyFormat = Mid(l, 3)
        ElseIf Left(l, 2) = "* " Then ' 番号なしリスト：レベル1
            .Range.SetListLevel Level:=1
            .Range.ListFormat.ApplyListTemplateWithLevel _
                ListTemplate:=ListGalleries(wdBulletGallery).ListTemplates(1), _
                ContinuePreviousList:=False, _
                ApplyTo:=wdListApplyToWholeList, _
                DefaultListBehavior:=wdWord10ListBehavior
            ApplyFormat = Mid(l, 3)
        ElseIf Left(l, 3) = "1. " Then ' 番号付きリスト：今のところ連番にできない
            .Range.ListFormat.ApplyListTemplateWithLevel _
                ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(7), _
                ContinuePreviousList:=False, _
                ApplyTo:=wdListApplyToWholeList, _
                DefaultListBehavior:=wdWord10ListBehavior
            ApplyFormat = Mid(l, 4)
        ElseIf l = "***" Then ' 水平線：なんかエラー出るので実装できてない
            ' With .Paragraphs.Borders(wdBorderBottom)
            '     .LineStyle = wdLineStyleSingle
            '     .Color = Options.DefaultBorderColor
            ' End With
            ApplyFormat = ""
        Else
            .Paragraphs.Style = ActiveDocument.Styles("標準")
            ApplyFormat = l
        End If
    End With
End Function
