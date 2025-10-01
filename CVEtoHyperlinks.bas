Option Explicit
Attribute VB_Name = "CVEtoHyperlinks"
Sub CVEtoHyperlinks()
    Dim regEx As Object
    ' Создаём объект регулярного выражения для поиска шаблонов в тексте документа.
    ' We create an object of regular expression to search for templates in the text of the document.
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Устанавливаем глобальный режим, чтобы находить все совпадения в документе.
    ' We set the global mode to find all the coincidences in the document.
    regEx.Global = True

    ' Для адаптации под иные нужды можно изменить данный паттерн.
    ' To adapt to other needs, you can change this pattern.
    ' Паттерн ищет строки вида "CVE-YYYY-ZZZZ", где YYYY — четыре цифры, а ZZZZ может состоять из 4–7 цифр.
    ' The pattern is looking for lines of the type "CVE-YYYY-ZZZZ", where YYYY is four digits, and ZZZZ can consist of 4-7 digits.
    regEx.Pattern = "CVE-\d{4}-\d{4,7}"
    
    Dim docRange As Range
    Set docRange = ActiveDocument.Range
    
    Dim matches As Object
    ' Выполняем поиск по всему тексту документа с использованием регулярного выражения.
    ' We search for the entire text of the document using regular expression.
    Set matches = regEx.Execute(docRange.Text)
    
    ' Если найдены совпадения, начинаем обработку каждого из них.
    ' If coincidences are found, we begin to process each of them.
    If matches.Count > 0 Then
        Dim match As Object
        For Each match In matches
            ' Используем метод Find для поиска конкретного вхождения найденного текста в документе.
            ' We use the Find method to search for a specific entry of the found text in the document.
            With ActiveDocument.Content.Find
                .Text = match.Value      ' Задаём искомый текст (например, "CVE-2021-12345")
                .Forward = True          ' Поиск осуществляется вперёд
                .Wrap = wdFindStop       ' Поиск прекращается, когда достигается конец документа
                
                ' Выполняем цикл поиска всех вхождений данного текста.
                ' We carry out the search cycle for all the entries of this text.
                Do While .Execute
                    Dim foundRange As Range
                    Set foundRange = .Parent
                    
                    ' Если в найденном диапазоне ещё нет гиперссылки, то добавляем её.
                    ' If the found range does not yet have a hyperlink, then add it.
                    If Not (foundRange.Hyperlinks.Count > 0) Then
                        ActiveDocument.Hyperlinks.Add Anchor:=foundRange, _
                                                      ' Для адаптации под иные нужды можно изменить данный адрес.
                                                      ' To adapt to other needs, you can change this address.
                                                      ' Здесь формируется URL с использованием найденного идентификатора CVE.
                                                      ' The URL is formed here using the found identifier CVE.
                                                      Address:="https://www.cve.org/CVERecord?id=" & match.Value, _
                                                      TextToDisplay:=match.Value
                    End If
                    
                    ' Сдвигаем начало диапазона поиска за конец текущего найденного диапазона,
                    ' We move the beginning of the search range for the end of the current found range,
                    ' чтобы избежать зацикливания на одном и том же вхождении.
                    ' To avoid bouncing on the same entry.
                    .Parent.Start = foundRange.End
                    If .Parent.Start >= ActiveDocument.Content.End Then Exit Do
                Loop
            End With
        Next match
    End If

    Set regEx = Nothing
End Sub
