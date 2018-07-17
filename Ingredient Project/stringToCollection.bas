'Separates ingredient string into array of ingredients, to test for duplicate ingredients
'Note: ingredient exception is "1,2 Hexanadrol"
Function stringToCollection(inactive_ingredient_string As String) As Collection

    Dim StrSplit As Variant
    StrSplit = Split(inactive_ingredient_string, ",")
    Exception = False
    Dim ingredient As Variant, newString As New Collection
    For Each ingredient In StrSplit
        ingredient = Trim(ingredient)
        If ingredient = "1" Then: Exception = True: GoTo resumeLoop
        If Len(ingredient) < 3 Then GoTo resumeLoop
        ingredient = Replace(ingredient, "'", "\'")
        If Exception Then: ingredient = "1," & ingredient: Exception = False
        If Right(ingredient, 1) = "." Then ingredient = Trim(Left(ingredient, Len(ingredient) - 1))
        newString.Add ingredient
resumeLoop:
    Next
    Set stringToCollection = newString
    
End Function
