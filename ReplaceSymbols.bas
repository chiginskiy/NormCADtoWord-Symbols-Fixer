Sub ReplaceGreekFontWithTimesNewRoman()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim character As Range
    Dim replacementCount As Integer
    Dim i As Long
    
    ' Таблица соответствия символов Greek -> Unicode коды
    Dim greekMap As Object
    Set greekMap = CreateObject("Scripting.Dictionary")
    
    ' Соответствие: латинская буква -> Unicode код греческой буквы
    greekMap("a") = 945   ' α (alpha) U+03B1
    greekMap("b") = 946   ' β (beta) U+03B2
    greekMap("g") = 947   ' γ (gamma) U+03B3
    greekMap("d") = 948   ' δ (delta) U+03B4
    greekMap("e") = 949   ' ε (epsilon) U+03B5
    greekMap("z") = 950   ' ζ (zeta) U+03B6
    greekMap("h") = 951   ' η (eta) U+03B7
    greekMap("q") = 952   ' θ (theta) U+03B8
    greekMap("i") = 953   ' ι (iota) U+03B9
    greekMap("k") = 954   ' κ (kappa) U+03BA
    greekMap("l") = 955   ' λ (lambda) U+03BB
    greekMap("m") = 956   ' μ (mu) U+03BC
    greekMap("n") = 957   ' ν (nu) U+03BD
    greekMap("x") = 958   ' ξ (xi) U+03BE
    greekMap("o") = 959   ' ο (omicron) U+03BF
    greekMap("p") = 960   ' π (pi) U+03C0
    greekMap("r") = 961   ' ρ (rho) U+03C1
    greekMap("s") = 963   ' σ (sigma) U+03C3
    greekMap("t") = 964   ' τ (tau) U+03C4
    greekMap("u") = 965   ' υ (upsilon) U+03C5
    greekMap("f") = 966   ' φ (phi) U+03C6
    greekMap("c") = 967   ' χ (chi) U+03C7
    greekMap("y") = 968   ' ψ (psi) U+03C8
    greekMap("w") = 969   ' ω (omega) U+03C9
    
    Set doc = ActiveDocument
    replacementCount = 0
    
    ' Проходим по всем абзацам в документе
    For Each para In doc.Paragraphs
        Set rng = para.Range
        
        ' Проходим по каждому символу в абзаце (в обратном порядке)
        For i = rng.Characters.Count To 1 Step -1
            Set character = rng.Characters(i)
            
            ' Проверяем, установлен ли шрифт Greek
            If character.Font.Name = "Greek" Then
                Dim charLower As String
                charLower = LCase(character.Text)
                
                ' Проверяем наличие в таблице соответствия
                If greekMap.Exists(charLower) Then
                    ' Используем ChrW() для правильной вставки Unicode
                    character.Text = ChrW(greekMap(charLower))
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
                    If character.Font.Name = "Greek" Then
                        Dim charLower2 As String
                        charLower2 = LCase(character.Text)
                        
                        If greekMap.Exists(charLower2) Then
                            character.Text = ChrW(greekMap(charLower2))
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
