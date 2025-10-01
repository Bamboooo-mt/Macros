Option Explicit

Attribute VB_Name = "findRepeated"

Sub findRepeated()
    Dim doc As Document
    Dim sentence As Range
    Dim words() As String
    Dim phraseDict As Object ' Dictionary
    Dim phrase As String
    Dim i As Long, j As Long
    
    Set doc = ActiveDocument
    Set phraseDict = CreateObject("Scripting.Dictionary")
    
    For Each sentence In doc.Sentences
        words = Split(sentence.Text, " ")
        ' Сначала собираем все двухсловные фразы в словарь с подсчетом
        ' First we collect all two -word phrases in the dictionary with the calculation
        For i = LBound(words) To UBound(words) - 1
            phrase = Trim(RemoveSpecialCharacters(words(i) & " " & words(i + 1)))
            If phrase <> "" Then
                If phraseDict.Exists(LCase(phrase)) Then
                    phraseDict(LCase(phrase)) = phraseDict(LCase(phrase)) + 1
                Else
                    phraseDict.Add LCase(phrase), 1
                End If
            End If
        Next i
        
        ' Если найдены повторения, подсвечиваем их и добавляем комментарии
        ' If repetitions are found, we highlight them and add comments
        Dim key As Variant
        For Each key In phraseDict.Keys
            If phraseDict(key) > 1 Then
                HighlightWordInSentence sentence, key, wdTurquoise
                AddCommentToFirstWord sentence, "Повтор фразы: " & key
            End If
        Next key
        
        ' Очищаем словарь для следующего предложения
        ' We clean the dictionary for the next sentence
        phraseDict.RemoveAll
    Next sentence
End Sub

Sub HighlightWordInSentence(ByRef sentence As Range, ByVal wordToHighlight As String, ByVal colorIndex As Long)
    Dim foundRange As Range
    Set foundRange = sentence.Duplicate
    With foundRange.Find
        .Text = wordToHighlight
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .Wrap = wdFindStop
        Do While .Execute
            If foundRange.InRange(sentence) Then
                foundRange.HighlightColorIndex = colorIndex
            End If
            foundRange.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Function RemoveSpecialCharacters(inputStr As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .Pattern = "[(),.!?:;]" ' набор символов для удаления
    End With
    RemoveSpecialCharacters = regEx.Replace(inputStr, "")
End Function

Sub AddCommentToFirstWord(sentence As Range, commentText As String)
    If sentence.Words.Count > 0 Then
        Dim firstWord As Range
        Set firstWord = sentence.Words(1)
        ' Добавляем комментарий, если его ещё нет
        ' Add a comment if it is not yet
        If firstWord.Text <> "" And firstWord.Comments.Count = 0 Then
            firstWord.Comments.Add firstWord, commentText
        End If
    End If
End Sub
