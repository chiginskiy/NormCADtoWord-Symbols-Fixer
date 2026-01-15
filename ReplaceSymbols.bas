Sub FindGreekFontCharacters()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim greekChars As String
    Dim charCount As Integer
    Dim foundAny As Boolean
    
    Set doc = ActiveDocument
    charCount = 0
    foundAny = False
    greekChars = ""
    
    ' Проходим по всем абзацам в документе
    For Each para In doc.Paragraphs
        Set rng = para.Range
        
        ' Проходим по каждому символу в абзаце
        For i = 1 To rng.Characters.Count
            Set character = rng.Characters(i)
            
            ' Проверяем, установлен ли шрифт Greek
            If character.Font.Name = "Greek" Then
                ' Добавляем символ в список (избегаем пустых символов)
                If character.Text <> "" And character.Text <> vbCr Then
                    greekChars = greekChars & character.Text & " "
                    charCount = charCount + 1
                    foundAny = True
                End If
            End If
        Next i
    Next para
    
    ' Также проверяем таблицы, если они есть в документе
    Dim tbl As Table
    Dim cell As Cell
    Dim cellRng As Range
    
    For Each tbl In doc.Tables
        For Each cell In tbl.Range.Cells
            Set cellRng = cell.Range
            For i = 1 To cellRng.Characters.Count
                Set character = cellRng.Characters(i)
                If character.Font.Name = "Greek" Then
                    If character.Text <> "" And character.Text <> vbCr Then
                        greekChars = greekChars & character.Text & " "
                        charCount = charCount + 1
                        foundAny = True
                    End If
                End If
            Next i
        Next cell
    Next tbl
    
    ' Добавляем результаты в конец документа
    Dim endPara As Paragraph
    Set endPara = doc.Range.Paragraphs(doc.Range.Paragraphs.Count).Range
    endPara.Collapse wdCollapseEnd
    
    ' Вставляем перевод строки перед результатами
    Selection.InsertParagraph
    
    ' Вставляем заголовок
    Set endPara = doc.Range.Paragraphs(doc.Range.Paragraphs.Count).Range
    endPara.Collapse wdCollapseEnd
    endPara.InsertParagraph
    endPara.Text = "=== РЕЗУЛЬТАТЫ ПОИСКА СИМВОЛОВ С ШРИФТОМ GREEK ==="
    endPara.Style = wdStyleHeading2
    
    ' Вставляем результаты
    If foundAny Then
        Set endPara = doc.Range.Paragraphs(doc.Range.Paragraphs.Count).Range
        endPara.Collapse wdCollapseEnd
        endPara.InsertParagraph
        Set endPara = doc.Range.Paragraphs(doc.Range.Paragraphs.Count).Range
        endPara.Text = "Найдено символов: " & charCount & vbCrLf & "Символы: " & greekChars
    Else
        Set endPara = doc.Range.Paragraphs(doc.Range.Paragraphs.Count).Range
        endPara.Collapse wdCollapseEnd
        endPara.InsertParagraph
        Set endPara = doc.Range.Paragraphs(doc.Range.Paragraphs.Count).Range
        endPara.Text = "Символов со шрифтом Greek не найдено."
    End If
    
    MsgBox "Поиск завершён! Найдено символов со шрифтом Greek: " & charCount, vbInformation
End Sub
