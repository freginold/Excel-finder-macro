Sub finder()

' finder Macro
' v1.2.1
    
    ' if cell is currently yellow, change it to turqoise (8)
    ' only counts number of columns based on first row
    ' only searches columns over to ZZ -- if need to go farther, change cellRange
    
    Dim word, numRows, cellRange, newColor, prevColor, firstAddress, c, NotBold, numFound, loc, locStr, FindNext
    
    word = InputBox("Word to Find?", "")
    
    If word = "" Then MsgBox ("Nothing to find"): Exit Sub
    
    numFound = 0
    numRows = ActiveSheet.Range("A1048576").End(xlUp).Row
    cellRange = "a1:zz" & numRows
    
    With ActiveSheet.Range(cellRange)
        Set c = .Find(word)
        If Not c Is Nothing Then
            ' found an occurrence
            firstAddress = c.Address
            Do
                numFound = numFound + 1
                newColor = 6
                prevColor = Range(c.Address).Interior.ColorIndex
                If prevColor = 6 Then newColor = 8
                Range(c.Address).Interior.ColorIndex = newColor
                If Range(c.Address).Font.Bold <> True Then
                    NotBold = False
                    Range(c.Address).Font.Bold = True
                Else
                    NotBold = True
                End If
                Range(c.Address).Select
                loc = Split(c.Address, "$")
                locStr = ""
                For Each x In loc
                    locStr = locStr + x
                Next
                locStr = "        #" & numFound & ":  " & locStr
                If .FindNext(c).Address <> firstAddress Then
                    ' if there's still another match to show
                    FindNext = MsgBox(locStr & vbCrLf & vbCrLf & "        Find Next?", 4)
                    ' Yes = 6, No = 7
                Else
                    ' this is the last match
                    MsgBox locStr, 0
                End If
                Range(c.Address).Interior.ColorIndex = prevColor
                If Not NotBold Then Range(c.Address).Font.Bold = False
                If FindNext = 7 Then Exit Sub
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> firstAddress
        End If
    End With
    
    If numFound = 0 Then MsgBox ("""" & word & """ not found in this sheet.")
    
End Sub
