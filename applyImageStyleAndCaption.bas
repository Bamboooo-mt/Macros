Option Explicit
Attribute VB_Name = "applyImageStyleAndCaption"
Sub applyImageStyleAndCaption()
    Dim shape As InlineShape      
    Dim captionText As String       
    Dim paragraph As Paragraph      
    Dim style As Style              
    Dim foundPTImage As Boolean     
    Dim foundPTImageName As Boolean 

    ' Инициализация флагов
    ' Initialization of flags
    foundPTImage = False
    foundPTImageName = False

    ' Перебираем все стили, доступные в документе, чтобы проверить, импортированы ли необходимые стили.
    ' We sort through all the styles available in the document to check whether the necessary styles are imported.
    For Each style In ActiveDocument.Styles
        ' Проверка наличия стиля для рисунка.
        ' Checking the availability of style for the picture.
        ' Здесь замените "<image style>" на фактическое имя стиля, который вы хотите применить к рисункам.
        ' Here, replace "<Image style>" with the actual name of the style that you want to apply to the drawings.
        If style.NameLocal = "<image style>" Then
            foundPTImage = True
        End If

        ' Проверка наличия стиля для подписи рисунка.
        ' Checking the availability of style for signing the picture.
        ' Здесь замените "<image name style>" на фактическое имя стиля, который вы хотите применить к подписи рисунка.
        ' Here, replace "<Image Name Style>" with the actual name of the style that you want to apply to the signature of the picture.
        If style.NameLocal = "<image name style>" Then
            foundPTImageName = True
        End If
    Next style

    ' Если ни один из нужных стилей не найден, выводим сообщение и завершаем выполнение процедуры.
    ' If none of the necessary styles is found, we display the message and complete the procedure.
    ' Это помогает избежать ошибок, если стили не импортированы из коллекции Normal.
    ' This helps to avoid mistakes if the styles are not imported from the Normal collection.
    If Not foundPTImage And Not foundPTImageName Then
        MsgBox "Импортируй стили из Normal"
        Exit Sub
    End If
    
    ' Перебираем все встроенные рисунки (InlineShapes) в документе.
    ' We sort out all the built -in drawings (Inlineshapes) in the document.
    For Each shape In ActiveDocument.InlineShapes
        ' Проверяем, что текущий объект является изображением.
        ' We check that the current object is an image.
        If shape.Type = wdInlineShapePicture Then

            ' Применяем стиль к рисунку.
            ' We use the style to the picture.
            ' Замените "<image style>" на фактическое имя стиля, который должен быть применён к изображению.
            ' Replace "<Image Style>" with the actual name of the style that should be applied to the image.
            shape.Range.Style = "<image style>"

            ' Добавляем подпись к изображению.
            ' Add the signature to the image.
            ' Метод InsertCaption вставляет подпись под изображением с указанными параметрами:
            ' The InsertCaption method inserts a signature under the image with the specified parameters:
            ' Label:="Рисунок" — метка подписи,
            ' Label: = "drawing" - signature mark,
            ' Title:="." — текст подписи (можно изменить на нужный),
            ' Title: = "."- the text of the signature (can be changed to the desired),
            ' Position:=wdCaptionPositionBelow — размещение подписи ниже изображения.
            ' Position: = WDCAPATIONPOSITIONBELOW - placement of the signature below the image.
            shape.Range.InsertCaption Label:="Рисунок", Title:=".", Position:=wdCaptionPositionBelow
        End If
    Next shape
    
    ' Перебираем все параграфы документа для корректировки стиля подписи рисунка.
    ' We sort through all the paragraphs of the document to adjust the style of the signing of the drawing.
    For Each paragraph In ActiveDocument.Paragraphs
        ' Если стиль параграфа равен "Название объекта", то предполагается, что данный параграф является подписью рисунка.
        ' If the paragraph style is "the name of the object", then it is assumed that this paragraph is the signature of the drawing.
        If paragraph.Style = "Название объекта" Then

            ' Применяем стиль для подписи рисунка.
            ' We use the style for signing the picture.
            ' Замените "<image name style>" на фактическое имя стиля, которое должно применяться к подписи.
            ' Replace "<Image Name Style>" with the actual name of the style, which should be applied to the signature.
            paragraph.Style = "<image name style>"
        End If
    Next paragraph
End Sub
