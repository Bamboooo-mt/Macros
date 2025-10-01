Option Explicit
Attribute VB_Name = "maskingPass"
Sub maskingPass()


Dim mask As String
Dim stars As String
Dim count1 As Long
Dim count2 As Long


    mask = Selection.text
    '��� ������� �����, ��� ���� ������������� �������� ����������� ���������, ������� ����������� �� ���������� ���������� ��������.
    ' This is the cycle counter, for it the value of the selected fragment is calculated, which decreases by the number of uncovered characters.
    count2 = Len(mask) - 2
    
    '���� ��� �������� ������ �� ��������� (�� ������ �������� �� ���� ������� �������)
    ' A cycle for creating a string of stars (you can replace the symbol with your version)
    For count1 = 1 To count2 Step 1
        stars = stars & "*"
        Next count1
    
    '���� ������ ������ �� �������� ������, ������������ ���������, ����������� � ����� ��������� ������ �� �������� ������
    ' We take the first symbol from the original line, attach stars, glue the last symbol from the original line to the end
    Options.ReplaceSelection = True
    Selection.text = Left(mask, 1) & stars & Right(mask, 1)
    
End Sub
