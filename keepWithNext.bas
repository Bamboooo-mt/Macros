
Attribute VB_Name = "keepWithNext"
Sub keepWithNext()
    With Selection.ParagraphFormat
        .keepWithNext = True
    End With
End Sub