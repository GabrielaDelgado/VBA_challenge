Attribute VB_Name = "Module1"
Sub Challenge2()
    ' Set an initial variable for the worksheets
    Dim ws As Worksheet

    ' Select and activate the worksheets to run the code on each sheet
    For Each ws In Worksheets
        ws.Activate
        
        ' Define an initial variable for the ticker symbol
        Dim ticker_symbol As String
        
        ' Set an initial variable for the total stock volume
        Dim total_volume As Double
        total_volume = 0
        
        ' Keep track of the location for each ticker number in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Set opening and closing stock values as variables
        Dim opening As Double
        Dim closing As Double
        Dim yearly_change As Double ' Changed from String to Double
            
        ' Yearly change starts with Cells C2 value
        opening = Range("C2").Value
        
        ' Loop through all stock data
        For i = 2 To 753001
        
            ' Check for the ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' Set the ticker symbol
                ticker_symbol = Cells(i, 1).Value
                
                ' Add to the total_volume
                total_volume = total_volume + Cells(i, 7).Value
                
                ' Print the ticker symbol in the Summary table
                Range("I" & Summary_Table_Row).Value = ticker_symbol
                
                ' Print the total stock volume in the Summary table
                Range("L" & Summary_Table_Row).Value = total_volume
                
                ' When the ticker symbol changes, grab closing value i and i +1 for opening value for the next symbol
                closing = Cells(i, 6).Value
                
                ' Subtract closing and opening values
                Cells(i, 10).Value = closing - opening
                
                ' Print the yearly change
                yearly_change = Range("J" & Summary_Table_Row).Value
                
                    ' Check if the yearly change is greater than 0
                    If yearly_change > 0 Then
                    
                        ' Color the cell green
                        Cells(i, 10).Interior.ColorIndex = 4
                    Else
                        ' Color the cell red
                        Cells(i, 10).Interior.ColorIndex = 3
                    End If
                
                 ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the total stock volume
                total_volume = 0
                
                ' Reset opening to i + 1
                opening = Cells(i + 1, 3).Value
            Else
                ' If the cells have the same ticker symbol, add to total_volume
                total_volume = total_volume + Cells(i, 7).Value
                
            End If
        Next i
        
        ' Set the range variables for each worksheet
        Dim range18 As Range
        Dim greatest_volume18 As Range
        Dim range19 As Range
        Dim greatest_volume19 As Range
        Dim range20 As Range
        Dim greatest_volume20 As Range
        
        Set range18 = Worksheets("2018").Range("J2:J3001")
        Set greatest_volume18 = Worksheets("2018").Range("L2:L3001")
        
        Set range19 = Worksheets("2019").Range("J2:J3001")
        Set greatest_volume19 = Worksheets("2019").Range("L2:L3001")
        
        Set range20 = Worksheets("2020").Range("J2:J3001")
        Set greatest_volume20 = Worksheets("2020").Range("L2:L3001")
        
        ' Calculate the maximum, minimum, and greatest volume for each year
        Dim max_increase18 As Double
        Dim min_increase18 As Double
        Dim greatest_increase18 As Double
        
        max_increase18 = Application.WorksheetFunction.Max(range18)
        min_increase18 = Application.WorksheetFunction.Min(range18)
        greatest_increase18 = Application.WorksheetFunction.Max(greatest_volume18)
        
        Dim max_increase19 As Double
        Dim min_increase19 As Double
        Dim greatest_increase19 As Double
        
        max_increase19 = Application.WorksheetFunction.Max(range19)
        min_increase19 = Application.WorksheetFunction.Min(range19)
        greatest_increase19 = Application.WorksheetFunction.Max(greatest_volume19)
        
        Dim max_increase20 As Double
        Dim min_increase20 As Double
        Dim greatest_increase20 As Double
        
        max_increase20 = Application.WorksheetFunction.Max(range20)
        min_increase20 = Application.WorksheetFunction.Min(range20)
        greatest_increase20 = Application.WorksheetFunction.Max(greatest_volume20)
    Next ws
End Sub
