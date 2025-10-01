Option Explicit
Attribute VB_Name = "replaceImagesFromFolder"
Sub replaceImagesFromFolder()
    Dim folderPath As String
    Dim dlg As FileDialog
    Dim picIndex As Integer
    Dim shp As InlineShape
    Dim newShp As InlineShape
    Dim picPath As String
    Dim caption As Range
    Dim captionText As String
    Dim picName As String
    Dim doc As Document

    ' Открываем диалог выбора папки
    ' Open the dialogue for choosing a folder
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = "Выберите папку с изображениями"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "Папка не выбрана. Макрос завершен.", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
        ' Добавляем завершающий обратный слеш, если его нет
        ' Add the final reverse slash if it is not
        If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    End With

    ' Устанавливаем активный документ
    ' Install an active document
    Set doc = ActiveDocument

    ' Проходим по всем картинкам в документе
    ' We pass through all the pictures in the document
    For Each shp In doc.InlineShapes
        ' Проверяем наличие подписи под картинкой
        ' Check the presence of a signature under the picture
        If shp.Range.Paragraphs.Last.Next Is Nothing Then
            GoTo NextImage
        Else
            Set caption = shp.Range.Paragraphs.Last.Next.Range
            captionText = caption.text
        End If
        
        ' Проверяем формат подписи "Рисунок X." или "Figure X."
        ' We check the format of the signature "Figure X."Or "Figure X."
        If InStr(1, captionText, "Рисунок") > 0 Or InStr(1, captionText, "Figure") > 0 Then
            If InStr(1, captionText, "Рисунок") > 0 Then
                picIndex = Val(Replace(Replace(captionText, "Рисунок", ""), ".", ""))
            ElseIf InStr(1, captionText, "Figure") > 0 Then
                picIndex = Val(Replace(Replace(captionText, "Figure", ""), ".", ""))
            End If
            ' Формируем имя файла картинки
            ' Form the name of the picture file
            picName = picIndex & ".png"
            picPath = folderPath & picName
            
            ' Если файл существует, заменяем картинку
            ' If the file exists, replace the picture
            If Dir(picPath) <> "" Then
                Set newShp = shp.Range.InlineShapes.AddPicture(FileName:=picPath, _
                                                               LinkToFile:=False, _
                                                               SaveWithDocument:=True)
                ' Копируем свойства старого изображения на новое
                ' Copy the properties of the old image for a new
                With newShp
                    .Shadow.Visible = True
                    .Shadow.Transparency = 0.75
                    .Shadow.Blur = 4
                    .Shadow.Size = 100
                    .Shadow.OffsetX = -0.001
                    .Shadow.OffsetY = 0.1
                    .LockAspectRatio = True
                    If shp.Width > 465 Then
                        .Width = 465
                    Else
                        .Height = shp.Height
                        .Width = shp.Width
                    End If
                End With
                ' Удаляем старое изображение
                ' We remove the old image
                shp.Delete
            End If
        End If
NextImage:
    Next shp
    MsgBox "Если ничего не произошло, то проверьте путь к папке с картинками.", vbExclamation
End Sub


