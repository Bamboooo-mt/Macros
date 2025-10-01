Option Explicit
Attribute VB_Name = "LinksToFootnotes"
Sub LinksToFootnotes()

    Dim rngSel As Range
    Set rngSel = Selection.Range
    
    ' Проверка: выделен ли какой-то непустой фрагмент
    ' Checking: Is some non-resistant fragment allocated
    If rngSel Is Nothing Then Exit Sub
    If rngSel.text = "" Or Selection.Type = wdSelectionIP Then
        MsgBox "Выделите текст, содержащий гиперссылки (стиль Hyperlink).", vbInformation
        Exit Sub
    End If
    
    Dim colLinks As New Collection
    
    ' Создадим Range для поиска внутри выделения
    ' Create Range to search inside the discharge
    Dim rngFind As Range
    Set rngFind = rngSel.Duplicate
    rngFind.Collapse Direction:=wdCollapseStart
    
    ' Настраиваем Find на поиск стиля Hyperlink
    ' We set up find to search for Hyperlink style
    With rngFind.Find
        .ClearFormatting
        .style = ActiveDocument.Styles(wdStyleHyperlink) ' Используем встроенный стиль
        .text = ""           
        .Forward = True      
        .Wrap = wdFindStop   
        .Format = True
    End With
    
    Do While rngFind.Find.Execute = True
        
        If rngFind.Start > rngSel.End Then
            Exit Do
        End If

        If rngFind.End <= rngSel.End Then
            Dim rngStore As Range
            Set rngStore = rngFind.Duplicate
            colLinks.Add rngStore
        Else
            Exit Do
        End If

        rngFind.Collapse Direction:=wdCollapseEnd
    Loop
    
    If colLinks.Count = 0 Then
        MsgBox "Не найдено ни одного фрагмента со стилем Hyperlink в выделенном тексте."
        Exit Sub
    End If
    
    ' Теперь обрабатываем найденные гиперссылки с конца к началу
    ' Now we process the found hyperlinks from the end to the beginning
    Dim footCount As Long
    footCount = 0
    
    Dim i As Long
    For i = colLinks.Count To 1 Step -1
        
        Dim rngLink As Range
        Set rngLink = colLinks(i)
        
        Dim linkText As String
        linkText = rngLink.text
        
        rngLink.text = ""
        
        ActiveDocument.Footnotes.Add _
            Range:=rngLink, _
            text:=linkText
        
        
        ' Удаляем все пробелы / табуляции / неразрывные пробелы перед только что вставленной сноской
        ' Remove all gaps / tabulation / inextricable gaps in front of the already inserted footnote
        Dim rngSpace As Range
        Set rngSpace = rngLink.Duplicate
        
        Do While (rngSpace.Start > rngSel.Start)
            ' Сдвигаемся на один символ назад
            ' We move to one symbol back
            rngSpace.MoveStart Unit:=wdCharacter, Count:=-1
            

            Select Case rngSpace.Characters.Last
                Case " ", Chr(160), vbTab
                    rngSpace.Characters.Last.Delete
                Case Else

                    Exit Do
            End Select
        Loop

        
        footCount = footCount + 1
    Next i
    

End Sub



