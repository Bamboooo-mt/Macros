Option Explicit
Attribute VB_Name = "exportHotKeys"
Sub exportHotKeys()
    Dim kb As KeyBinding         ' ���������� ��� �������� ���������� ������
    Dim output As String         ' ��������� ���������� ��� ������������ ������ ����
    Dim fileNum As Integer       ' ���������� ��� �������� ������ ����� (��� �������� �����/������)
    Dim desktopPath As String    ' ������, ���������� ���� � ����� �� ������� �����

    ' ���������� ���� � ����� �� ������� �����: ������������ ���������� ��������� USERPROFILE
    ' Determine the path to the file on the desktop: Userprofile environment is used
    desktopPath = Environ("USERPROFILE") & "\Desktop\hotkeys.bas"
    
    ' �������� ��������� ����� ����� ��� �������� ����� �� ������
    ' We get a free file number to open a file for recording
    fileNum = FreeFile
    Open desktopPath For Output As #fileNum
    
    ' ���������� ������ ��������� SetHotkeys, ������� ����� �������� ������� �������
    ' Record the beginning of the Sethotkeys procedure, which will set the hot keys
    Print #fileNum, "Sub SetHotkeys()"
    Print #fileNum, "    Dim keyBindings As KeyBindings"
    Print #fileNum, "    Set keyBindings = Application.KeyBindings"
    Print #fileNum, "    CustomizationContext = NormalTemplate"
    
    ' ���������� ��� ���������� ������ � ����������
    ' We sort out all the assignments of the keys in the application
    For Each kb In Application.keyBindings
        ' ���� ��������� ���������� ��������� � ��������, �� ������������ � �������������� ����������� ��������� AddMacroHotkey
        ' If the destination category refers to the macros, then we export using the Special Addmacrohotkey procedure
        If kb.KeyCategory = wdKeyCategoryMacro Then
            Dim keyCode2Part As String
            ' ���� � ���������� ������������ �������������� ��� ������� (KeyCode2) � ��� ��� �� ����� "wdKey�",
            ' If the purpose of the Keycode2 is present in the purpose and his name is not equal to "Wdkey",
            ' �� ��������� �������������� ����� ��� ��������.
            ' then we form an additional part for export.
            If kb.KeyCode2 <> 0 And GetKeyCodeName(kb.KeyCode2) <> "wdKey�" Then
                keyCode2Part = ", " & GetKeyCodeName(kb.KeyCode2)
            Else
                keyCode2Part = ""
            End If

            ' ��������� ������, ���������� ��������� AddMacroHotkey � �����������:
            ' We form a string that causes the Addmacrohotkey procedure with parameters:
            ' - ��� ������� (� ��������� ���������� ����� �������)
            ' - Macro name (with the removal of the constant part of the team)
            ' - ��������� ������������� ��������� ���� ������� (KeyCode)
            ' - string representation of the main key code (Keycode)
            ' - �������������� ��� ������� (���� �������)
            ' - additional key code (if any)
            Print #fileNum, "    AddMacroHotkey """ & RemoveConstantPart(kb.command) & """, " & _
                            GetKeyCodeString(kb.KeyCode) & keyCode2Part
        Else
            ' ��� ����������, �� ����������� � ��������, ��������� ������ ������ ������ Add � �����������:
            ' For appointments that are not related to macros, we form a line of calling the Add method with parameters:
            ' ��������� ����������, ������� � ��� �������.
            ' Purpose category, team and key code.
            output = "    keyBindings.Add KeyCategory:=" & KeyCategoryName(kb.KeyCategory) & _
                     ", Command:=""" & kb.command & """"

            ' ��������� ��� �������, ��������� ������� BuildKeyCode � ��������� ������������� ����
            ' Add the key code using the BuildkeyCode function and the stringent show of code
            output = output & ", KeyCode:=BuildKeyCode(" & GetKeyCodeString(kb.KeyCode) & ")"

            ' ���� ������������ �������������� ��� ������� (KeyCode2), ��������� ��� � ������
            ' If there is an additional key code (keycode2), add it to the line
            If kb.KeyCode2 <> 0 And GetKeyCodeName(kb.KeyCode2) <> "" And GetKeyCodeName(kb.KeyCode2) <> "wdKey�" Then
                output = output & ", KeyCode2:=" & GetKeyCodeName(kb.KeyCode2)
            End If

            ' ���������� �������������� ������ � ����
            ' Record the formed line in the file
            Print #fileNum, output
        End If
    Next kb
    ' ���������� ��������� ��������� SetHotkeys
    ' Record the end of the Sethotkeys procedure
    Print #fileNum, "End sub"
    
    ' ������������ ��������� AddMacroHotkey, ������� ��������� ���������� ������� ������� ��� �������.
    ' We export the Addmacrohotkey procedure, which adds a hot key for the macros.
    Print #fileNum, "Sub AddMacroHotkey(baseCommand As String, KeyCode As Long, Optional KeyCode1 As Long = 0, Optional KeyCode2 As Long = 0)"
    Print #fileNum, "    On Error Resume Next"
    Print #fileNum, "    Dim combinedKeyCode As Long"
    ' ����� ����� ���� �����������, ������� ����� �������� ��� �������������.
    ' There may be a comment that can be added if necessary.
    Print #fileNum, "    "
    ' ���� ����� �������� ��� ������� ������ wdKeyControl, �������������� KeyCode1 ����� wdKeyAlt � ���� KeyCode2,
    ' If the main key code is equal to WDKEYCONTROL, the additional KEYCODE1 is WDKEYALT and there is KEYCODE2,
    ' �� ����������� ��� ��� ���� �������.
    ' then we combine all three key code code.
    Print #fileNum, "    If KeyCode = wdKeyControl And KeyCode1 = wdKeyAlt And KeyCode2 <> 0 Then"
    Print #fileNum, "        combinedKeyCode = BuildKeyCode(KeyCode, KeyCode1, KeyCode2)"
    Print #fileNum, "        Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, command:=baseCommand, KeyCode:=combinedKeyCode"
    ' ���� ������������ �������������� ��� ������� KeyCode2, ����������� �������� ���� � ��������� ���������� � �������������� �����.
    ' If there is an additional keycode2 key code, we combine the main codes and add the purpose with additional code.
    Print #fileNum, "    ElseIf KeyCode2 <> 0 Then"
    Print #fileNum, "        combinedKeyCode = BuildKeyCode(KeyCode, KeyCode1)"
    Print #fileNum, "        Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, command:=baseCommand, KeyCode:=combinedKeyCode, KeyCode2:=KeyCode2"
    ' ����� ����������� ������ �������� ��� � KeyCode1.
    ' Otherwise, we combine only the main code and Keycode1.
    Print #fileNum, "    Else"
    Print #fileNum, "        combinedKeyCode = BuildKeyCode(KeyCode, KeyCode1)"
    Print #fileNum, "        Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, command:=baseCommand, KeyCode:=combinedKeyCode"
    Print #fileNum, "    End If"
    Print #fileNum, "    On Error GoTo 0"
    Print #fileNum, "End Sub"
    
    ' ������������ ������� BuildKeyCode, ������� ��������� ����� ����� ������ � ���������� �� ��������� ��������.
    ' We export the BuildkeyCode function, which accepts the set of keys codes and returns their total value.
    Print #fileNum, "Function BuildKeyCode(ParamArray keys() As Variant) As Long"
    Print #fileNum, "    Dim i As Integer, code As Long"
    Print #fileNum, "    code = 0"
    Print #fileNum, "    For i = LBound(keys) To UBound(keys)"
    Print #fileNum, "        code = code + keys(i)"
    Print #fileNum, "    Next i"
    Print #fileNum, "    BuildKeyCode = code"
    Print #fileNum, "End Function"
    
    ' ��������� ���� ����� ������ ���� ������.
    ' Close the file after recording all the data.
    Close #fileNum
End Sub

' ������� ���������� ��������� ������������� ��������� ���������� ������ �� ������ � ��������� ��������.
' The function returns the string representation of the key assignment category based on its numerical value.
Function KeyCategoryName(category As WdKeyCategory) As String
    Select Case category
        Case wdKeyCategoryDisable: KeyCategoryName = "wdKeyCategoryDisable"
        Case wdKeyCategoryCommand: KeyCategoryName = "wdKeyCategoryCommand"
        Case wdKeyCategoryMacro: KeyCategoryName = "wdKeyCategoryMacro"
        Case wdKeyCategoryFont: KeyCategoryName = "wdKeyCategoryFont"
        Case wdKeyCategoryAutoText: KeyCategoryName = "wdKeyCategoryAutoText"
        Case wdKeyCategoryStyle: KeyCategoryName = "wdKeyCategoryStyle"
        Case wdKeyCategorySymbol: KeyCategoryName = "wdKeyCategorySymbol"
        Case wdKeyCategoryPrefix: KeyCategoryName = "wdKeyCategoryPrefix"
        Case wdKeyCategoryBookmark: KeyCategoryName = "wdKeyCategoryBookmark"
        Case wdKeyCategoryField: KeyCategoryName = "wdKeyCategoryField"
        Case wdKeyCategoryMailMerge: KeyCategoryName = "wdKeyCategoryMailMerge"
        Case wdKeyCategoryFormField: KeyCategoryName = "wdKeyCategoryFormField"
        Case wdKeyCategoryList: KeyCategoryName = "wdKeyCategoryList"
        Case Else: KeyCategoryName = "wdKeyCategoryUnknown"
    End Select
End Function

' ������� ���������� ��������� ��� ��� ��������� ���� �������.
' The function returns the string name for the given key code.
Function GetKeyCodeName(KeyCode As Long) As String
    Select Case KeyCode
        Case vbKeyUp: GetKeyCodeName = "vbKeyUp"
        Case vbKeyDown: GetKeyCodeName = "vbKeyDown"
        Case vbKeyLeft: GetKeyCodeName = "vbKeyLeft"
        Case vbKeyRight: GetKeyCodeName = "vbKeyRight"
        Case vbKeyReturn: GetKeyCodeName = "vbKeyReturn"
        Case vbKeyTab: GetKeyCodeName = "vbKeyTab"
        Case vbKeyEscape: GetKeyCodeName = "vbKeyEscape"
        Case vbKeyBack: GetKeyCodeName = "vbKeyBack"
        Case vbKeyDelete: GetKeyCodeName = "vbKeyDelete"
        Case vbKeyInsert: GetKeyCodeName = "vbKeyInsert"
        Case vbKeyHome: GetKeyCodeName = "vbKeyHome"
        Case vbKeyEnd: GetKeyCodeName = "vbKeyEnd"
        Case vbKeyPageUp: GetKeyCodeName = "vbKeyPageUp"
        Case vbKeyPageDown: GetKeyCodeName = "vbKeyPageDown"
        Case vbKeyHyphen, 189, 109: GetKeyCodeName = "wdKeyHyphen"
        Case 188: GetKeyCodeName = "wdKeyComma"
        Case 190: GetKeyCodeName = "wdKeyPeriod"
        Case 191: GetKeyCodeName = "wdKeySlash"
        Case Else: GetKeyCodeName = "wdKey" & Chr(KeyCode)
    End Select
End Function

' ������� ��������� ��������� ������������� ���� ������� � ������ ������������� (Ctrl, Alt, Shift)
' The function forms the string representation of the key code taking into account the modifiers (CTRL, ALT, Shift)
Function GetKeyCodeString(KeyCode As Long) As String
    Dim result As String
    result = ""
    
    ' ���������, ���������� �� ���� ������� Control, � ���� ��, ��������� ��������������� �����
    ' Check if the Control key flag is set, and if so, add the corresponding text
    If (KeyCode And wdKeyControl) <> 0 Then result = result & "wdKeyControl, "
    
    ' ���������, ���������� �� ���� ������� Alt
    ' Check if the flag of the Alt key
    If (KeyCode And wdKeyAlt) <> 0 Then result = result & "wdKeyAlt, "

    ' ���������, ���������� �� ���� ������� Shift
    ' Check if the shift key flag is set
    If (KeyCode And wdKeyShift) <> 0 Then result = result & "wdKeyShift, "
    
    ' ���� ������������ ���� ���������, ������� ����������� ������� � ������
    ' If the modifiers were added, remove the final comma and the gap
    If result <> "" Then result = Left(result, Len(result) - 2)
    
    ' ��������� �������� ��� ���� �������, ����������� ��� ������� 8 ����� ����
    ' Add the main name of the key code calculated for the younger 8 bits code
    result = result & ", " & GetKeyCodeName(KeyCode And &HFF)
    
    GetKeyCodeString = result
End Function

' ������� RemoveConstantPart ������� �� ������� ���������� �����, �������� ������ �� �����
' The RemoVeContantPart function deleys the constant part from the team, sharing the line by point
' � ��������� ��������� ������� �������, ��� ��������� �������� ������ ��� �������.
' And returning the last element of the array, which allows you to leave only the name of the macros.
Function RemoveConstantPart(command As String) As String
    Dim parts() As String
    parts = Split(command, ".")
    RemoveConstantPart = parts(UBound(parts))
End Function
