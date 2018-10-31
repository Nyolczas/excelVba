Attribute VB_Name = "Module1"
Sub Refresh()

    Application.DisplayAlerts = False
    '-------------------------------------------------------------
    
    Dim mt4Path As String, file As String, path As String
    
    mt4Path = Environ("UserProfile") & _
    "\Documents\Tozsde\MT4ek\XM1\MQL4\Files\csvStatement_19061180\"
    
    '-------------------------------------------------------------
    file = "tickBalance.csv"
    path = mt4Path & file
    Call dataCopy(path, "A1:F12000", "tickData", "A2", file)
    Call rowFill(1, 8, "G", "K")

    '-------------------------------------------------------------
    file = "robot_daily.csv"
    path = mt4Path & "robot\" & file
    Call dataCopy(path, "A1:C1000", "dData", "A2", file)
    
    '-------------------------------------------------------------
    file = "manual_daily.csv"
    path = mt4Path & "manual\" & file
    Call dataCopy(path, "A1:C1000", "dData", "D2", file)
    
    '-------------------------------------------------------------
    file = "depoHistory.csv"
    path = mt4Path & file
    Call dataCopy(path, "A1:C1000", "depoHistory", "A11", file)
    Call rowFill(1, 6, "E", "H")
    
    '-------------------------------------------------------------
    file = "usdhufPrices.csv"
    path = Environ("UserProfile") & _
            "\Documents\Tozsde\MT4ek\XM1\MQL4\Files\" & file
    Call dataCopy(path, "A1:B6000", "usdhuf", "A2", file)
    
    '-------------------------------------------------------------
    Sheets("Balance").Select
    
    Dim Ido As Date, Y As Integer
     
        Do While Ido < Now()
            Y = Cells(Rows.Count, 2).End(xlUp).Row
            Ido = Cells(Y, 2)
            Cells(Y + 1, 2) = Ido + 1
        Loop
        
     Call rowFill(2, 3, "C", "M")
     Call rowFill(2, 1, "A", "A")
    '-------------------------------------------------------------
    Application.DisplayAlerts = True

End Sub

Sub dataCopy(ByVal path As String, rng As String, sht As String, ByVal place As String, ByVal file As String)
    ' csv fileokat másol a megadott helyre
    
    If (Dir(path) > "") Then
        
        Workbooks.OpenText Filename:=path, DataType:=xlDelimited, Comma:=True, Local:=True
        Range(rng).Select
        Selection.Copy
        ThisWorkbook.Activate
        Sheets(sht).Select
        Range(place).Select
        ActiveSheet.Paste
        Workbooks(file).Close SaveChanges:=False
        
    End If
    
End Sub

Sub rowFill(refColumn As Integer, goalColumn As Integer, rowFrom As String, rowTo)
    ' kitölti a hiányzó sorokat
    
    Dim current As String, goal As String, adatSor As Integer, sumSor As Integer
    
    adatSor = Cells(Rows.Count, refColumn).End(xlUp).Row
    sumSor = Cells(Rows.Count, goalColumn).End(xlUp).Row
    current = rowFrom & sumSor & ":" & rowTo & sumSor
    goal = rowFrom & sumSor & ":" & rowTo & adatSor
    
    If (adatSor > sumSor) Then
        Range(current).Select
        Selection.AutoFill Destination:=Range(goal), Type:=xlFillDefault
    End If

End Sub
