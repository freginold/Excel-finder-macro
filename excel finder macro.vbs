Sub finder()

' finder Macro
    
    ' if cell is currently yellow, change it to turqoise (8)
    ' only counts number of columns based on first row
    ' only searches columns over to ZZ -- if need to go farther, change cellRange
    
    Dim word, numRows, cellRange, newColor, prevColor, firstAddress, c, NotBold
    
    word = InputBox("Word to Find?", "")
    
    If word = "" Then MsgBox ("Nothing to find"): Exit Sub
    
    numRows = ActiveSheet.Range("A1048576").End(xlUp).Row
    cellRange = "a1:z" & numRows
    
    With Worksheets(1).Range(cellRange)
        Set c = .Find(word)
        If Not c Is Nothing Then
            ' found an occurrence
            firstAddress = c.Address
            Do
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
                MsgBox (c.Address)
                Range(c.Address).Interior.ColorIndex = prevColor
                If Not NotBold Then Range(c.Address).Font.Bold = False
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> firstAddress
        End If
    End With
End Sub