Sub FindGreekFontCharacters()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim greekChars As String
    Dim charCount As Integer
    Dim foundAny As Boolean
    Dim i As Long
    
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
    
    If doc.Tables.Count > 0 Then
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
    End If
    
    ' Добавляем результаты в конец документа
    Dim resultRange As Range
    Set resultRange = doc.Range
    resultRange.Collapse wdCollapseEnd
    
    ' Вставляем пустую строку
    resultRange.InsertParagraphAfter
    resultRange.Collapse wdCollapseEnd
    resultRange.InsertParagraphAfter
    
    ' Вставляем заголовок
    resultRange.Text = "=== РЕЗУЛЬТАТЫ ПОИСКА СИМВОЛОВ С ШРИФТОМ GREEK ===" & vbCrLf
    resultRange.Font.Bold = True
    resultRange.Font.Size = 14
    resultRange.Collapse wdCollapseEnd
    
    ' Вставляем результаты
    If foundAny Then
        resultRange.Text = "Найдено символов: " & charCount & vbCrLf & _
                          "Символы: " & greekChars & vbCrLf
    Else
        resultRange.Text = "Символов со шрифтом Greek не найдено." & vbCrLf
    End If
    
    MsgBox "Поиск завершён! Найдено символов со шрифтом Greek: " & charCount, vbInformation
End Sub
