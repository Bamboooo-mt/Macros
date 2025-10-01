Option Explicit
Attribute VB_Name = "CVEtoHyperlinks"
Sub CVEtoHyperlinks()
    Dim regEx As Object
    ' ������ ������ ����������� ��������� ��� ������ �������� � ������ ���������.
    ' We create an object of regular expression to search for templates in the text of the document.
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' ������������� ���������� �����, ����� �������� ��� ���������� � ���������.
    ' We set the global mode to find all the coincidences in the document.
    regEx.Global = True

    ' ��� ��������� ��� ���� ����� ����� �������� ������ �������.
    ' To adapt to other needs, you can change this pattern.
    ' ������� ���� ������ ���� "CVE-YYYY-ZZZZ", ��� YYYY � ������ �����, � ZZZZ ����� �������� �� 4�7 ����.
    ' The pattern is looking for lines of the type "CVE-YYYY-ZZZZ", where YYYY is four digits, and ZZZZ can consist of 4-7 digits.
    regEx.Pattern = "CVE-\d{4}-\d{4,7}"
    
    Dim docRange As Range
    Set docRange = ActiveDocument.Range
    
    Dim matches As Object
    ' ��������� ����� �� ����� ������ ��������� � �������������� ����������� ���������.
    ' We search for the entire text of the document using regular expression.
    Set matches = regEx.Execute(docRange.Text)
    
    ' ���� ������� ����������, �������� ��������� ������� �� ���.
    ' If coincidences are found, we begin to process each of them.
    If matches.Count > 0 Then
        Dim match As Object
        For Each match In matches
            ' ���������� ����� Find ��� ������ ����������� ��������� ���������� ������ � ���������.
            ' We use the Find method to search for a specific entry of the found text in the document.
            With ActiveDocument.Content.Find
                .Text = match.Value      ' ����� ������� ����� (��������, "CVE-2021-12345")
                .Forward = True          ' ����� �������������� �����
                .Wrap = wdFindStop       ' ����� ������������, ����� ����������� ����� ���������
                
                ' ��������� ���� ������ ���� ��������� ������� ������.
                ' We carry out the search cycle for all the entries of this text.
                Do While .Execute
                    Dim foundRange As Range
                    Set foundRange = .Parent
                    
                    ' ���� � ��������� ��������� ��� ��� �����������, �� ��������� �.
                    ' If the found range does not yet have a hyperlink, then add it.
                    If Not (foundRange.Hyperlinks.Count > 0) Then
                        ActiveDocument.Hyperlinks.Add Anchor:=foundRange, _
                                                      ' ��� ��������� ��� ���� ����� ����� �������� ������ �����.
                                                      ' To adapt to other needs, you can change this address.
                                                      ' ����� ����������� URL � �������������� ���������� �������������� CVE.
                                                      ' The URL is formed here using the found identifier CVE.
                                                      Address:="https://www.cve.org/CVERecord?id=" & match.Value, _
                                                      TextToDisplay:=match.Value
                    End If
                    
                    ' �������� ������ ��������� ������ �� ����� �������� ���������� ���������,
                    ' We move the beginning of the search range for the end of the current found range,
                    ' ����� �������� ������������ �� ����� � ��� �� ���������.
                    ' To avoid bouncing on the same entry.
                    .Parent.Start = foundRange.End
                    If .Parent.Start >= ActiveDocument.Content.End Then Exit Do
                Loop
            End With
        Next match
    End If

    Set regEx = Nothing
End Sub
