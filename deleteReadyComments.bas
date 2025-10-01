Option Explicit
Attribute VB_Name = "deleteReadyComments"
Sub deleteReadyComments()
    Dim comment As comment
    Dim checkedScopes As New Collection
    Dim currentScope As String
    Dim allReady As Boolean

    ' Проходим по всем комментариям в документе.
    ' We pass through all the comments in the document.
    ' Если комментарий отмечен как выполненный (Done = True), то обрабатываем его область (scope).
    ' If the comment is noted as executed (Done = True), then we process its area (Scope).
    For Each comment In ActiveDocument.Comments
        If comment.Done = True Then
            currentScope = comment.scope
            ' Если область еще не проверялась, то проводим проверку выполненности всех комментариев в этой области.
            ' If the region has not yet been checked, then we are checking the execution of all comments in this area.
            If Not IsInCollection(checkedScopes, currentScope) Then
                allReady = AreAllCommentsReadyByScope(currentScope)
                ' Если все комментарии для данной области выполнены или область не задана, удаляем комментарии этой области.
                ' If all the comments for this area are executed or the region is not set, we delete the comments of this area.
                If allReady Or currentScope = "" Then
                    DeleteCommentsByScope currentScope
                End If  
                ' Добавляем область в коллекцию обработанных, чтобы не проверять ее повторно.
                ' Add the area to the processed collection so as not to check it again.
                checkedScopes.Add currentScope
            End If
        End If
    Next comment
End Sub

' Функция проверяет, что для заданной области (scope) все комментарии помечены как выполненные.
' The function checks that for a given area (Scope), all comments are marked as executed.
Function AreAllCommentsReadyByScope(scope As String) As Boolean
    Dim comment As comment
    Dim foundNotReady As Boolean
    
    ' Перебираем все комментарии и ищем хотя бы один, который относится к указанной области и не выполнен.
    ' We sort through all the comments and look for at least one that refers to the specified area and is not executed.
    For Each comment In ActiveDocument.Comments
        If comment.scope = scope Then
            If Not comment.Done Then
                foundNotReady = True
                Exit For
            End If
        End If
    Next comment
    
    ' Если не найден ни один не выполненный комментарий, возвращаем True.
    ' If not a single comment is found, we return True.
    AreAllCommentsReadyByScope = Not foundNotReady
End Function

' Функция проверяет, содержится ли заданный ключ (например, область комментария) в коллекции.
' The function checks whether the specified key (for example, the commentary area) is contained in the collection.
Function IsInCollection(coll As Collection, key As Variant) As Boolean
    On Error Resume Next
    IsInCollection = Not coll(key) Is Nothing
    On Error GoTo 0
End Function

' Процедура удаляет все комментарии, принадлежащие указанной области (scope), если они отмечены как выполненные.
' The procedure removes all the comments belonging to the specified area (Scope) if they are marked as executed.
Sub DeleteCommentsByScope(scope As String)
    Dim comment As comment
    
    ' Проходим по всем комментариям и удаляем те, у которых совпадает область и свойство Done установлено в True.
    ' We pass through all the comments and delete those in which the area and the Done property coincide are installed in True.
    For Each comment In ActiveDocument.Comments
        If comment.Done = True And comment.scope = scope Then
            comment.Delete
        End If
    Next comment
End Sub
