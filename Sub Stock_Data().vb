Sub Stock_Data()

Dim Stock_Price_Start As Double
Dim Stock_Price_End As Double
Dim Stock_Volume As Long
Dim Symbol As String


Range("K1").Value = "Stock Symbol"
Range("L1").Value = "Dollar Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "End Stock Volume"

Range("A1").Activate

Stock_Price_Start = ActiveCell.Offset(1, 5).Value

MsgBox ("Starting Price is now " & Stock_Price_Start)

'Need to find a way to identify max end range instead of hard-coding the last cell each time

For i = 2 To 797711

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    
    Symbol = Cells(i, 1).Value
    Stock_Price_End = Cells(i, 1).Offset(0, 5).Value
    Stock_Volume = Cells(i, 1).Offset(0, 6).Value
    
'Hi Grading Powers-that-be!  I'm not sure if the below solution to select and enter the table data was ideal / appropriate;  Is there a better method I should utilize in the future?  Thank you for your time and thoughts!

    Range("K1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = Symbol
         
    ActiveCell.Offset(0, 1).Value = (Stock_Price_End - Stock_Price_Start)
    ActiveCell.Offset(0, 1).NumberFormat = "#,##0.00_);-#,##0.00_)"
    
    If (Stock_Price_End - Stock_Price_Start) = 0 Then
    ActiveCell.Offset(0, 2).Value = 0
    End If
    
    If Stock_Price_Start = "0" Then
    ActiveCell.Offset(0, 2).Value = "-"
    End If
        
    If (Stock_Price_End - Stock_Price_Start) <> 0 Then
    ActiveCell.Offset(0, 2).Value = ((Stock_Price_End - Stock_Price_Start) / Stock_Price_Start)
    End If
    
    ActiveCell.Offset(0, 2).NumberFormat = "0.00%"
    ActiveCell.Offset(0, 3).Value = Stock_Volume
    ActiveCell.Offset(0, 3).NumberFormat = "#,##0"
    
    Cells(i, 1).Select

    Stock_Price_Start = ActiveCell.Offset(1, 5).Value
    
    End If
    
   ' MsgBox ("Your New Stock start price is " & Stock_Price_Start)

Next i

    Range("L2", Range("L2").End(xlDown)).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False



End Sub


