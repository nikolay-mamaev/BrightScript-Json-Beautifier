Function BeautifyJsonString(jsonString As String, indentSpacesNumber = 4) As String
    dictionaryStartChar = "{"
    dictionaryEndChar = "}"
    arrayStartChar = "["
    arrayEndChar = "]"
    commaChar = ","
    colonChar = ":"
    newLineChar = Chr(10)
    carriageReturnChar = Chr(13)
    spaceChar = " "
    quoteChar = Chr(34)
    tabChar = Chr(9)
    indentString = String(indentSpacesNumber, spaceChar)
    indentLevel = 0
    beautifiedJsonString = ""

    i = 1
    While i <= Len(jsonString)
        curChar = Mid(jsonString, i, 1)
        If curChar = quoteChar Then
            indexOfClosingQuote = Instr(i + 1, jsonString, quoteChar)
            If indexOfClosingQuote > 0 Then
                beautifiedJsonString = beautifiedJsonString + Mid(jsonString, i, indexOfClosingQuote - i + 1)
                i = indexOfClosingQuote
            Else
                beautifiedJsonString = beautifiedJsonString + curChar
            End If
        Else
            If curChar = dictionaryStartChar OR curChar = arrayStartChar Then
                indentLevel = indentLevel + 1
                beautifiedJsonString = beautifiedJsonString + curChar + newLineChar + String(indentLevel, indentString)
            Else If curChar = dictionaryEndChar OR curChar = arrayEndChar Then
                indentLevel = indentLevel - 1
                If indentLevel < 0 Then
                    indentLevel = 0
                End If
                beautifiedJsonString = beautifiedJsonString + newLineChar + String(indentLevel, indentString) + curChar
            Else If curChar = commaChar Then
                beautifiedJsonString = beautifiedJsonString + curChar + newLineChar + String(indentLevel, indentString)
            Else If curChar = colonChar Then
                beautifiedJsonString = beautifiedJsonString + spaceChar + curChar + spaceChar
            Else If NOT(curChar = spaceChar OR curChar = tabChar OR curChar = newLineChar OR curChar = carriageReturnChar) Then
                beautifiedJsonString = beautifiedJsonString + curChar
            End If
        End If

        i = i + 1
    End While

    return beautifiedJsonString
End Function
