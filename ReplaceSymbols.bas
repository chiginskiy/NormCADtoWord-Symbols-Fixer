' =============================================================================
' NormCADtoWord-Symbols-Fixer
' =============================================================================
' Описание: Замена Greek и Math Light шрифтов на Unicode в Times New Roman
' Версия: 1.1.0
' Дата: 15.01.2026
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
' МОДУЛЬ 2: Замена символов на Unicode
' =============================================================================
' Назначение: Заменяет символы Greek и Math Light на соответствующие
'             Unicode-символы в Times New Roman
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
                    replacementCount = replacementCount + 1
                End If
            
            ' Проверяем шрифт Math Light
            ElseIf character.Font.Name = "Math Light" Then
                If mathMap.Exists(character.Text) Then
                    character.Text = ChrW(mathMap(character.Text))
                    character.Font.Name = "Times New Roman"
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
                            replacementCount = replacementCount + 1
                        End If
                    
                    ' Math Light
                    ElseIf character.Font.Name = "Math Light" Then
                        If mathMap.Exists(character.Text) Then
                            character.Text = ChrW(mathMap(character.Text))
                            character.Font.Name = "Times New Roman"
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
           "- Math Light → математические символы", vbInformation, "Результат"
    
End Sub
