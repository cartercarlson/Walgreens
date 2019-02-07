'Cleans ingredient string, by factoring in:
'   a. Human error of ingredient entry by supplier
'   b. Constraints of entering ingredient information into Walgreens' website
'   c. Standardization of ingredient string to better locate duplicate ingredients
Function cleanString(stringToCheck) As String

    If InStr(stringToCheck, ",") = 0 Then stringToCheck = Replace(stringToCheck, ".", ",")
    stringToCheck = WorksheetFunction.Proper(stringToCheck)
    stringToCheck = Replace(stringToCheck, vbLf & vbLf, vbLf)
    stringToCheck = Replace(stringToCheck, "   ", " ")
    stringToCheck = Replace(stringToCheck, "  ", " ")
    stringToCheck = Replace(stringToCheck, vbLf, ",")
    stringToCheck = Replace(stringToCheck, "`", "'")
    stringToCheck = Replace(stringToCheck, "’", "'")
    stringToCheck = Replace(stringToCheck, "''", "'")
    stringToCheck = Replace(stringToCheck, "\", "")
    stringToCheck = Replace(stringToCheck, " /", "/")
    stringToCheck = Replace(stringToCheck, "/ ", "/")
    stringToCheck = Replace(stringToCheck, " )", ")")
    stringToCheck = Replace(stringToCheck, "( ", "(")
    stringToCheck = Replace(stringToCheck, "", "")
    stringToCheck = Replace(stringToCheck, " : ", "")
    stringToCheck = Replace(stringToCheck, Chr(63), "") ' heart
    stringToCheck = Replace(stringToCheck, Chr(134), "") ' cross
    stringToCheck = Replace(stringToCheck, Chr(174), "") '  ® symbol
    stringToCheck = Replace(stringToCheck, "®", "")
    stringToCheck = Replace(stringToCheck, "•", ",")
    stringToCheck = Replace(stringToCheck, "(And)", ",")
    stringToCheck = Replace(stringToCheck, "Contains ", "")
    stringToCheck = Replace(stringToCheck, ", and ", ",")
    stringToCheck = Replace(stringToCheck, ", And ", ",")
    stringToCheck = Replace(stringToCheck, ", AND ", ",")
    stringToCheck = Replace(stringToCheck, "[+/-", ",[+/-")
    stringToCheck = Replace(stringToCheck, "[May Contain", ",[May Contain")
    stringToCheck = Replace(stringToCheck, ", Oil", " Oil")
    stringToCheck = Replace(stringToCheck, ", Seed", " Seed")
    stringToCheck = Replace(stringToCheck, ", Extract", " Extract")
    stringToCheck = Replace(stringToCheck, ", Root", " Root")
    stringToCheck = Replace(stringToCheck, ", Flower", " Flower")
    cleanString = stringToCheck

End Function
