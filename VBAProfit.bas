Attribute VB_Name = "VBAProfit"

' Creating my profit  loss calculation function
Public Sub profitLoss()
    'Setting up my variables
    Dim profitSheet As Worksheet
    Dim priceUnit As Double
    Dim costUnit As Double
    Dim units As Double

    ' setting up the worksheet
    Set profitSheet = Worksheets("profit")

    ' working of the sheet
    With profitSheet
        ' initializing variables
        priceUnit = Range("B5")
        costUnit = Range("C5")
        units = Range("D5")

        ' profit and loss calculations
        revenue = priceUnit * units
        cost = costUnit * units
        profitAndLoss = revenue - cost

        ' add results to table
        Range("B8") = revenue
        Range("B9") = cost
        Range("B10") = profitAndLoss

    End With

End Sub


