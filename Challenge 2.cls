VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Challenge2()

' Select multiple worksheets code and range to max values was gathered from EverydayVBA YouTube videos

    ' Set an initial variable for the worksheets
    Dim ws As Worksheet
    
    ' Select and activate the owrksheets to run in the code at the same time
    For Each ws In Worksheets
        ws.Activate
        
        ' Define an initial variable for the ticker symbol
        Dim ticker_symbol As String
        
        ' Set an initial variable for the toal vstock volume
        Dim total_volume As Double
        total_volume = 0
        
        ' Keep track of the location for each ticker number in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Loop through all stock data
        For i = 2 To 753001
        
            ' Check for the ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' Set the ticker symbol
                ticker_symbol = Cells(i, 1).Value
                
                ' Add to the total_volume
                total_volume = total_volume + Cells(i, 7)
                
                ' Print the ticker symbol in the SUmmary table
                Range("I" & Summary_Table_Row).Value = ticker_symbol
                
                ' Printthe toal stock volume in the Summary table
                Range("L" & Summary_Table_Row).Value = total_volume
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the total stock volume
                total_volume = 0
                
            ' If the cells have the same tocker symbol
            Else
            
                ' Add to total_volume
                total_volume = total_volume + Cells(i, 7).Value
                
            ' Set opening and closing stock values as variables
            Dim opening As Double
            Dim closing As Double
            
                
            ' Check if the yearly change is greater than 0
            If Cells(i, 10).Value > 0 Then
            
                ' Color the cell green
                Cells(i, 10).Interior.ColorIndex = 4
                
            Else
            
                ' Color the cell Red
                Cells(i, 10).Interior.ColorIndex = 3
                
            End If
            
        Next i
        
        Set range18 = Worksheets("2018").Range("J")
        Set greatest_volume18 = Worksheets("2018").Range("L")
        
        Set range19 = Worksheets("2019").Range("J")
        Set greatest_volume19 = Worksheets("2019").Range("L")
        
        Set range18 = Worksheets("2020").Range("J")
        Set greatest_volume18 = Worksheets("2020").Range("L")
        
        max_increase18 = Application.WorksheetFunction.Max(range18)
        min_increase18 = Application.WorksheetFunction.Min(range18)
        greatest_increase18 = Application.WorksheetFunction.Max(greatest_volume18)
        
        max_increase19 = Application.WorksheetFunction.Max(range19)
        min_increase19 = Application.WorksheetFunction.Min(range19)
        greatest_increase19 = Application.WorksheetFunction.Max(greatest_volume19)
        
        max_increase20 = Application.WorksheetFunction.Max(range20)
        min_increase20 = Application.WorksheetFunction.Min(range20)
        greatest_increase20 = Application.WorksheetFunction.Max(greatest_volume20)
        
    Next
    
End Sub
