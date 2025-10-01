Option Explicit
Attribute VB_Name = "applyImageStyleAndCaption"
Sub applyImageStyleAndCaption()
    Dim shape As InlineShape      
    Dim captionText As String       
    Dim paragraph As Paragraph      
    Dim style As Style              
    Dim foundPTImage As Boolean     
    Dim foundPTImageName As Boolean 

    ' ������������� ������
    ' Initialization of flags
    foundPTImage = False
    foundPTImageName = False

    ' ���������� ��� �����, ��������� � ���������, ����� ���������, ������������� �� ����������� �����.
    ' We sort through all the styles available in the document to check whether the necessary styles are imported.
    For Each style In ActiveDocument.Styles
        ' �������� ������� ����� ��� �������.
        ' Checking the availability of style for the picture.
        ' ����� �������� "<image style>" �� ����������� ��� �����, ������� �� ������ ��������� � ��������.
        ' Here, replace "<Image style>" with the actual name of the style that you want to apply to the drawings.
        If style.NameLocal = "<image style>" Then
            foundPTImage = True
        End If

        ' �������� ������� ����� ��� ������� �������.
        ' Checking the availability of style for signing the picture.
        ' ����� �������� "<image name style>" �� ����������� ��� �����, ������� �� ������ ��������� � ������� �������.
        ' Here, replace "<Image Name Style>" with the actual name of the style that you want to apply to the signature of the picture.
        If style.NameLocal = "<image name style>" Then
            foundPTImageName = True
        End If
    Next style

    ' ���� �� ���� �� ������ ������ �� ������, ������� ��������� � ��������� ���������� ���������.
    ' If none of the necessary styles is found, we display the message and complete the procedure.
    ' ��� �������� �������� ������, ���� ����� �� ������������� �� ��������� Normal.
    ' This helps to avoid mistakes if the styles are not imported from the Normal collection.
    If Not foundPTImage And Not foundPTImageName Then
        MsgBox "���������� ����� �� Normal"
        Exit Sub
    End If
    
    ' ���������� ��� ���������� ������� (InlineShapes) � ���������.
    ' We sort out all the built -in drawings (Inlineshapes) in the document.
    For Each shape In ActiveDocument.InlineShapes
        ' ���������, ��� ������� ������ �������� ������������.
        ' We check that the current object is an image.
        If shape.Type = wdInlineShapePicture Then

            ' ��������� ����� � �������.
            ' We use the style to the picture.
            ' �������� "<image style>" �� ����������� ��� �����, ������� ������ ���� ������� � �����������.
            ' Replace "<Image Style>" with the actual name of the style that should be applied to the image.
            shape.Range.Style = "<image style>"

            ' ��������� ������� � �����������.
            ' Add the signature to the image.
            ' ����� InsertCaption ��������� ������� ��� ������������ � ���������� �����������:
            ' The InsertCaption method inserts a signature under the image with the specified parameters:
            ' Label:="�������" � ����� �������,
            ' Label: = "drawing" - signature mark,
            ' Title:="." � ����� ������� (����� �������� �� ������),
            ' Title: = "."- the text of the signature (can be changed to the desired),
            ' Position:=wdCaptionPositionBelow � ���������� ������� ���� �����������.
            ' Position: = WDCAPATIONPOSITIONBELOW - placement of the signature below the image.
            shape.Range.InsertCaption Label:="�������", Title:=".", Position:=wdCaptionPositionBelow
        End If
    Next shape
    
    ' ���������� ��� ��������� ��������� ��� ������������� ����� ������� �������.
    ' We sort through all the paragraphs of the document to adjust the style of the signing of the drawing.
    For Each paragraph In ActiveDocument.Paragraphs
        ' ���� ����� ��������� ����� "�������� �������", �� ��������������, ��� ������ �������� �������� �������� �������.
        ' If the paragraph style is "the name of the object", then it is assumed that this paragraph is the signature of the drawing.
        If paragraph.Style = "�������� �������" Then

            ' ��������� ����� ��� ������� �������.
            ' We use the style for signing the picture.
            ' �������� "<image name style>" �� ����������� ��� �����, ������� ������ ����������� � �������.
            ' Replace "<Image Name Style>" with the actual name of the style, which should be applied to the signature.
            paragraph.Style = "<image name style>"
        End If
    Next paragraph
End Sub
