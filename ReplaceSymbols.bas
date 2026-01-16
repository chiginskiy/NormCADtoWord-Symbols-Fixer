' =============================================================================
' NormCADtoWord-Symbols-Fixer
' =============================================================================
' Описание: Замена Greek и Math Light шрифтов на Unicode в Times New Roman
' Версия: 1.4.0
' Дата: 16.01.2026
' Автор: chiginskiy
' Лицензия: Apache 2.0
' Репозиторий: https://github.com/chiginskiy/NormCADtoWord-Symbols-Fixer
' =============================================================================

' =============================================================================
' МОДУЛЬ 1: Визуальное выделение проблемных символов
' =============================================================================
' Назначение: Находит все символы Greek и Math Light, делает их красными
'             и полужирными для визуальной проверки перед заменой
' =============================================================================

Sub HighlightGreekAndMathSymbols()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim highlightCount As Integer
    Dim greekCount As Integer
    Dim mathCount As Integer
    Dim i As Long
    
    Set doc = ActiveDocument
    highlightCount = 0
    greekCount = 0
    mathCount = 0
    
    ' Проходим по всем абзацам в документе
    For Each para In doc.Paragraphs
        Set rng = para.Range
        
        ' Проходим по каждому символу в абзаце
        For i = 1 To rng.Characters.Count
            Set character = rng.Characters(i)
            
            ' Проверяем шрифт Greek
            If character.Font.Name = "Greek" Then
                If character.Text <> "" And character.Text <> vbCr Then
                    character.Font.Color = wdColorRed
                    character.Font.Bold = True
                    highlightCount = highlightCount + 1
                    greekCount = greekCount + 1
                End If
            
            ' Проверяем шрифт Math Light
            ElseIf character.Font.Name = "Math Light" Then
                If character.Text <> "" And character.Text <> vbCr Then
                    character.Font.Color = wdColorRed
                    character.Font.Bold = True
                    highlightCount = highlightCount + 1
                    mathCount = mathCount + 1
                End If
            End If
        Next i
    Next para
    
    ' Проверяем таблицы
    Dim tbl As Table
    Dim cell As Cell
    Dim cellRng As Range
    
    If doc.Tables.Count > 0 Then
        For Each tbl In doc.Tables
            For Each cell In tbl.Range.Cells
                Set cellRng = cell.Range
                For i = 1 To cellRng.Characters.Count
                    Set character = cellRng.Characters(i)
                    
                    ' Greek
                    If character.Font.Name = "Greek" Then
                        If character.Text <> "" And character.Text <> vbCr Then
                            character.Font.Color = wdColorRed
                            character.Font.Bold = True
                            highlightCount = highlightCount + 1
                            greekCount = greekCount + 1
                        End If
                    
                    ' Math Light
                    ElseIf character.Font.Name = "Math Light" Then
                        If character.Text <> "" And character.Text <> vbCr Then
                            character.Font.Color = wdColorRed
                            character.Font.Bold = True
                            highlightCount = highlightCount + 1
                            mathCount = mathCount + 1
                        End If
                    End If
                Next i
            Next cell
        Next tbl
    End If
    
    ' Сообщение о результате
    MsgBox "Выделение завершено!" & vbCrLf & vbCrLf & _
           "Всего выделено: " & highlightCount & vbCrLf & _
           "- Greek: " & greekCount & vbCrLf & _
           "- Math Light: " & mathCount & vbCrLf & vbCrLf & _
           "Символы выделены КРАСНЫМ и ПОЛУЖИРНЫМ", vbInformation, "Результат"
    
End Sub

' =============================================================================
' МОДУЛЬ 2: Замена символов на Unicode (с исправлением форматирования)
' =============================================================================
' Назначение: Заменяет символы Greek и Math Light на соответствующие
'             Unicode-символы в Times New Roman.
'             Сбрасывает форматирование (цвет черный, не полужирный)
' =============================================================================

Sub ReplaceGreekAndMathFontsToTimesNewRoman()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim replacementCount As Integer
    Dim i As Long
    
    ' Таблица соответствия для Greek
    Dim greekMap As Object
    Set greekMap = CreateObject("Scripting.Dictionary")
    
    greekMap("a") = 945   ' α (alpha)
    greekMap("b") = 946   ' β (beta)
    greekMap("g") = 947   ' γ (gamma)
    greekMap("d") = 948   ' δ (delta)
    greekMap("e") = 949   ' ε (epsilon)
    greekMap("z") = 950   ' ζ (zeta)
    greekMap("h") = 951   ' η (eta)
    greekMap("q") = 952   ' θ (theta)
    greekMap("i") = 953   ' ι (iota)
    greekMap("k") = 954   ' κ (kappa)
    greekMap("l") = 955   ' λ (lambda)
    greekMap("m") = 956   ' μ (mu)
    greekMap("n") = 957   ' ν (nu)
    greekMap("x") = 958   ' ξ (xi)
    greekMap("o") = 959   ' ο (omicron)
    greekMap("p") = 960   ' π (pi)
    greekMap("r") = 961   ' ρ (rho)
    greekMap("s") = 963   ' σ (sigma)
    greekMap("t") = 964   ' τ (tau)
    greekMap("u") = 965   ' υ (upsilon)
    greekMap("f") = 966   ' φ (phi)
    greekMap("c") = 967   ' χ (chi)
    greekMap("y") = 968   ' ψ (psi)
    greekMap("w") = 969   ' ω (omega)
    
    ' Таблица соответствия для Math Light
    Dim mathMap As Object
    Set mathMap = CreateObject("Scripting.Dictionary")
    
    mathMap("r") = 8804   ' ≤ (меньше или равно) U+2264
    mathMap("t") = 8805   ' ≥ (больше или равно) U+2265
    mathMap(";") = 8730   ' √ (квадратный корень) U+221A
    
    Set doc = ActiveDocument
    replacementCount = 0
    
    ' Проходим по всем абзацам в документе
    For Each para In doc.Paragraphs
        Set rng = para.Range
        
        ' Проходим по каждому символу в абзаце (в обратном порядке)
        For i = rng.Characters.Count To 1 Step -1
            Set character = rng.Characters(i)
            Dim charLower As String
            charLower = LCase(character.Text)
            
            ' Проверяем шрифт Greek
            If character.Font.Name = "Greek" Then
                If greekMap.Exists(charLower) Then
                    character.Text = ChrW(greekMap(charLower))
                    character.Font.Name = "Times New Roman"
                    ' Сброс форматирования после визуальной проверки
                    character.Font.Color = wdColorAutomatic  ' Черный
                    character.Font.Bold = False              ' Не полужирный
                    replacementCount = replacementCount + 1
                End If
            
            ' Проверяем шрифт Math Light
            ElseIf character.Font.Name = "Math Light" Then
                If mathMap.Exists(character.Text) Then
                    character.Text = ChrW(mathMap(character.Text))
                    character.Font.Name = "Times New Roman"
                    ' Сброс форматирования после визуальной проверки
                    character.Font.Color = wdColorAutomatic  ' Черный
                    character.Font.Bold = False              ' Не полужирный
                    replacementCount = replacementCount + 1
                End If
            End If
        Next i
    Next para
    
    ' Проверяем таблицы
    Dim tbl As Table
    Dim cell As Cell
    Dim cellRng As Range
    
    If doc.Tables.Count > 0 Then
        For Each tbl In doc.Tables
            For Each cell In tbl.Range.Cells
                Set cellRng = cell.Range
                For i = cellRng.Characters.Count To 1 Step -1
                    Set character = cellRng.Characters(i)
                    Dim charLower2 As String
                    charLower2 = LCase(character.Text)
                    
                    ' Greek
                    If character.Font.Name = "Greek" Then
                        If greekMap.Exists(charLower2) Then
                            character.Text = ChrW(greekMap(charLower2))
                            character.Font.Name = "Times New Roman"
                            ' Сброс форматирования
                            character.Font.Color = wdColorAutomatic
                            character.Font.Bold = False
                            replacementCount = replacementCount + 1
                        End If
                    
                    ' Math Light
                    ElseIf character.Font.Name = "Math Light" Then
                        If mathMap.Exists(character.Text) Then
                            character.Text = ChrW(mathMap(character.Text))
                            character.Font.Name = "Times New Roman"
                            ' Сброс форматирования
                            character.Font.Color = wdColorAutomatic
                            character.Font.Bold = False
                            replacementCount = replacementCount + 1
                        End If
                    End If
                Next i
            Next cell
        Next tbl
    End If
    
    ' Сообщение о результате
    MsgBox "Замена завершена!" & vbCrLf & _
           "Заменено символов: " & replacementCount & vbCrLf & vbCrLf & _
           "- Greek → греческие буквы" & vbCrLf & _
           "- Math Light → математические символы" & vbCrLf & _
           "- Форматирование сброшено (черный, обычный)", vbInformation, "Результат"
    
End Sub

' =============================================================================
' МОДУЛЬ 3: Простая смена шрифта без замены символов
' =============================================================================
' Назначение: Меняет шрифт Greek/Math Light на Times New Roman
'             БЕЗ замены самих символов, сбрасывает форматирование
' =============================================================================

Sub ChangeFontToTimesNewRomanOnly()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim changeCount As Integer
    Dim greekCount As Integer
    Dim mathCount As Integer
    Dim i As Long
    
    Set doc = ActiveDocument
    changeCount = 0
    greekCount = 0
    mathCount = 0
    
    ' Проходим по всем абзацам в документе
    For Each para In doc.Paragraphs
        Set rng = para.Range
        
        ' Проходим по каждому символу в абзаце
        For i = 1 To rng.Characters.Count
            Set character = rng.Characters(i)
            
            ' Проверяем шрифт Greek
            If character.Font.Name = "Greek" Then
                If character.Text <> "" And character.Text <> vbCr Then
                    ' Меняем только шрифт и форматирование
                    character.Font.Name = "Times New Roman"
                    character.Font.Bold = False
                    character.Font.Color = wdColorAutomatic  ' Черный
                    changeCount = changeCount + 1
                    greekCount = greekCount + 1
                End If
            
            ' Проверяем шрифт Math Light
            ElseIf character.Font.Name = "Math Light" Then
                If character.Text <> "" And character.Text <> vbCr Then
                    ' Меняем только шрифт и форматирование
                    character.Font.Name = "Times New Roman"
                    character.Font.Bold = False
                    character.Font.Color = wdColorAutomatic  ' Черный
                    changeCount = changeCount + 1
                    mathCount = mathCount + 1
                End If
            End If
        Next i
    Next para
    
    ' Проверяем таблицы
    Dim tbl As Table
    Dim cell As Cell
    Dim cellRng As Range
    
    If doc.Tables.Count > 0 Then
        For Each tbl In doc.Tables
            For Each cell In tbl.Range.Cells
                Set cellRng = cell.Range
                For i = 1 To cellRng.Characters.Count
                    Set character = cellRng.Characters(i)
                    
                    ' Greek
                    If character.Font.Name = "Greek" Then
                        If character.Text <> "" And character.Text <> vbCr Then
                            character.Font.Name = "Times New Roman"
                            character.Font.Bold = False
                            character.Font.Color = wdColorAutomatic
                            changeCount = changeCount + 1
                            greekCount = greekCount + 1
                        End If
                    
                    ' Math Light
                    ElseIf character.Font.Name = "Math Light" Then
                        If character.Text <> "" And character.Text <> vbCr Then
                            character.Font.Name = "Times New Roman"
                            character.Font.Bold = False
                            character.Font.Color = wdColorAutomatic
                            changeCount = changeCount + 1
                            mathCount = mathCount + 1
                        End If
                    End If
                Next i
            Next cell
        Next tbl
    End If
    
    ' Сообщение о результате
    MsgBox "Смена шрифта завершена!" & vbCrLf & vbCrLf & _
           "Всего изменено: " & changeCount & vbCrLf & _
           "- Greek: " & greekCount & vbCrLf & _
           "- Math Light: " & mathCount & vbCrLf & vbCrLf & _
           "Шрифт изменён на Times New Roman" & vbCrLf & _
           "Форматирование: обычное, черное" & vbCrLf & vbCrLf & _
           "⚠️ ВНИМАНИЕ: Символы НЕ заменены на Unicode!", vbInformation, "Результат"
    
End Sub

' =============================================================================
' МОДУЛЬ 4: Очистка пунктуации и пробелов
' =============================================================================
' Назначение: Удаляет лишние пробелы перед знаками препинания и двойные пробелы.
'             Выполняет замены:
'             - "  " -> " " (два пробела -> один)
'             - " ;" -> ";" (пробел+точка с запятой -> точка с запятой)
'             - " :" -> ":" (пробел+двоеточие -> двоеточие)
'             - " ." -> "." (пробел+точка -> точка)
'             - " ," -> "," (пробел+запятая -> запятая)
' =============================================================================

Sub CleanupPunctuationAndSpaces()
    Dim doc As Document
    Dim rng As Range
    Dim findObj As Find
    Dim cleanupCount As Integer
    
    Set doc = ActiveDocument
    Set rng = doc.Content
    Set findObj = rng.Find
    cleanupCount = 0
    
    ' Настройка параметров поиска
    findObj.ClearFormatting
    findObj.Replacement.ClearFormatting
    findObj.Forward = True
    findObj.Wrap = wdFindContinue
    findObj.Format = False
    findObj.MatchCase = False
    findObj.MatchWholeWord = False
    findObj.MatchWildcards = False
    findObj.MatchSoundsLike = False
    findObj.MatchAllWordForms = False
    
    ' 1. Замена двойных пробелов на одинарные
    ' Выполняем в цикле, чтобы убрать тройные и более пробелы
    findObj.Text = "  "
    findObj.Replacement.Text = " "
    
    Do While findObj.Execute(Replace:=wdReplaceAll)
        ' Продолжаем пока есть двойные пробелы
        cleanupCount = cleanupCount + 1 ' Счетчик циклов, а не замен
    Loop
    
    ' 2. Замена пробел + точка с запятой -> точка с запятой
    findObj.Text = " ;"
    findObj.Replacement.Text = ";"
    findObj.Execute Replace:=wdReplaceAll
    
    ' 3. Замена пробел + двоеточие -> двоеточие
    findObj.Text = " :"
    findObj.Replacement.Text = ":"
    findObj.Execute Replace:=wdReplaceAll
    
    ' 4. Замена пробел + точка -> точка
    findObj.Text = " ."
    findObj.Replacement.Text = "."
    findObj.Execute Replace:=wdReplaceAll
    
    ' 5. Замена пробел + запятая -> запятая
    findObj.Text = " ,"
    findObj.Replacement.Text = ","
    findObj.Execute Replace:=wdReplaceAll
    
    MsgBox "Очистка пунктуации завершена!" & vbCrLf & _
           "Удалены двойные пробелы и лишние отступы перед знаками препинания.", _
           vbInformation, "Результат очистки"
    
End Sub
