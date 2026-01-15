Sub ReplaceGreekFontWithTimesNewRoman()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim replacementCount As Integer
    Dim i As Long
    
    ' Таблица соответствия символов Greek -> греческие символы Times New Roman
    Dim greekMap As Object
    Set greekMap = CreateObject("Scripting.Dictionary")
    
    ' Соответствие символов
    greekMap("g") = "γ"  ' gamma
    greekMap("d") = "δ"  ' delta
    greekMap("e") = "ε"  ' epsilon
    greekMap("z") = "ζ"  ' zeta
    greekMap("h") = "η"  ' eta
    greekMap("q") = "θ"  ' theta
    greekMap("i") = "ι"  ' iota
    greekMap("k") = "κ"  ' kappa
    greekMap("l") = "λ"  ' lambda
    greekMap("m") = "μ"  ' mu
    greekMap("n") = "ν"  ' nu
    greekMap("x") = "ξ"  ' xi
    greekMap("o") = "ο"  ' omicron
    greekMap("p") = "π"  ' pi
    greekMap("r") = "ρ"  ' rho
    greekMap("s") = "σ"  ' sigma
    greekMap("t") = "τ"  ' tau
    greekMap("u") = "υ"  ' upsilon
    greekMap("f") = "φ"  ' phi
    greekMap("c") = "χ"  ' chi
    greekMap("y") = "ψ"  ' psi
    greekMap("w") = "ω"  ' omega
    greekMap("a") = "α"  ' alpha
    greekMap("b") = "β"  ' beta
    
    Set doc = ActiveDocument
    replacementCount = 0
    
    ' Проходим по всем абзацам в документе
    For Each para In doc.Paragraphs
        Set rng = para.Range
        
        ' Проходим по каждому символу в абзаце
        For i = rng.Characters.Count Down To 1
            Set character = rng.Characters(i)
            
            ' Проверяем, установлен ли шрифт Greek
            If character.Font.Name = "Greek" Then
                ' Получаем символ в нижнем регистре для поиска в таблице
                Dim charLower As String
                charLower = LCase(character.Text)
                
                ' Проверяем наличие в таблице соответствия
                If greekMap.Exists(charLower) Then
                    ' Заменяем символ
                    character.Text = greekMap(charLower)
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
                For i = cellRng.Characters.Count Down To 1
                    Set character = cellRng.Characters(i)
                    If character.Font.Name = "Greek" Then
                        Dim charLower2 As String
                        charLower2 = LCase(character.Text)
                        
                        If greekMap.Exists(charLower2) Then
                            character.Text = greekMap(charLower2)
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
           "Заменено символов: " & replacementCount, vbInformation, "Результат"
    
End Sub
