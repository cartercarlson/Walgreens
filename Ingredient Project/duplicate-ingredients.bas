'Look for an indredient listed twice - Walgreens website does not allow duplicate ingredients
'(algorithm - "Triangular number")
Sub DuplicateIngredients(ingredientString As Collection)
    Dim a as Long, b As Long
    For a = 1 To ingredientString.Count
        For b = a + 1 To ingredientString.Count
            'check for too long of a string or duplicate ingredients
            If Len(ingredientString(a)) > 80 Then
                MainSheet.Range("T" & sku.Row) = "Item ingredients are missing commas for proper separation"
                issue = True
            ElseIf ingredientString(a) = ingredientString(b) Then
                MainSheet.Range("T" & sku.Row) = "Item has the ingredient '" & ingredientString(a) & "' listed twice"
                issue = True
            End If
            If issue Then Exit Sub
        Next b
    Next a
End Sub
