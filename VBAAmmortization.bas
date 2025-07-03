Attribute VB_Name = "VBAAmmortization"

Public Sub CreateAmmortizationIOTable()
    ' define sheet
    Dim ammortizationSheet as Worksheet

    ' set it up
    Set ammortizationSheet = Worksheets("ammortization")

    ' designing table 
    With ammortizationSheet.Range("A1:B1")
        .Merge
        .Value = "Inputs"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
    End With

    With ammortizationSheet.Range("A7:B7")
        .Merge
        .Value = "Outputs"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
    End With

    With ammortizationSheet
        Range("A2") = "Loan"
        Range("A3") = "Nominal interest (p.a)"
        Range("A4") = "Frequency of payments per year"
        Range("A5") = "Term in Years"
        Range("A2:A5,B2:B5").Borders.LineStyle = xlContinuous ' add borders
        Range("A2:A5").Interior.ColorIndex = 23
        Range("A2:A5").Font.ColorIndex = 2
        Range("B2:B5").Interior.ColorIndex = 20
        Range("B2:B5").Font.ColorIndex = 1

        Range("A8") = "Total number of payments"
        Range("A9") = "Effective interest rate"
        Range("A10") = "Installment repayment"
        Range("A8:A10,B8:B10").Borders.LineStyle = xlContinuous ' add borders
        Range("A8:A10").Interior.ColorIndex = 23
        Range("A8:A10").Font.ColorIndex = 2
        Range("B8:B10").Interior.ColorIndex = 20
        Range("B8:B10").Font.ColorIndex = 1
    End With
    ' autofit the data
    ammortizationSheet.Range("A:B").EntireColumn.AutoFit
End Sub

Public Sub createAmmortizationTable()
    ' setting up input variables
    Dim LoanAmount as Variant
    Dim nominalInterestRate as Variant
    Dim frequencyOfPayments as Variant
    Dim term as Integer

    ' define sheet
    Dim ammortizationSheet as Worksheet

    ' set it up
    Set ammortizationSheet = Worksheets("ammortization")

    ' retrieve input data
    With ammortizationSheet
        LoanAmount = Range("B2").Value
        nominalInterestRate = Range("B3").Value
        frequencyOfPayments = Range("B4").Value
        term = Range("B5").Value
    End With

    ' clear ammortization table
    ammortizationSheet.Range("D1:J" & ammortizationSheet.Rows.Count).Clear

    ' designing  ammortization table 
    With ammortizationSheet.Range("D1:J1")
        .Merge
        .Value = "Ammortization Table"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
    End With

    With ammortizationSheet
        .Range("D2") = "Year"
        .Range("E2") = "Payment number"
        .Range("F2") = "Installment repayment"
        .Range("G2") = "Interest payment"
        .Range("H2") = "Principal repayment"
        .Range("I2") = "Principal outstanding start"
        .Range("J2") = "Principal outstanding end"
        .Range("D2:J2").Borders.LineStyle = xlContinuous ' add borders
        .Range("D2:J2").Interior.ColorIndex = 50 ' add background color
        .Range("D2:J2").Font.ColorIndex = 2 ' add font color
    End With

    'Compute outputs
    With ammortizationSheet
        .Range("B8") = frequencyOfPayments * term ' number of payments
        .Range("B9") = (nominalInterestRate  / 100) / frequencyOfPayments ' effective interest
        .Range("B10") = LoanAmount * ((.Range("B9") * (1 + .Range("B9"))^(.Range("B8"))) / ((1 + .Range("B9"))^(.Range("B8")) - 1))
        .Range("B10").NumberFormat = """R"" #,##0.00"
        .Range("B9").NumberFormat = "0.00%"
    End With

    ' working calculations for the table
    Dim steps As Long: steps = frequencyOfPayments * term 
    Dim stepSize As Double : stepSize = 1 / frequencyOfPayments
    Dim startVal As Double: startVal = stepSize
    Dim i As Long
    Dim counter  As Integer : counter = 1

    ' dynamically calculate the  values and update the table
    Dim principalOutstanding As Double
    principalOutstanding = LoanAmount

    For i = 0 To steps - 1
        ammortizationSheet.Range("D3").Offset(i, 0).Value = startVal + i * stepSize ' year calculation
        ammortizationSheet.Range("E3").Offset(i, 0).Value = counter ' payment number
        ammortizationSheet.Range("F3").Offset(i, 0).Value = ammortizationSheet.Range("B10").Value ' installment
        ammortizationSheet.Range("I3").Offset(i, 0).Value = principalOutstanding ' principal outstanding start
        ammortizationSheet.Range("G3").Offset(i, 0).Value = principalOutstanding * ammortizationSheet.Range("B9").Value ' interest payment
        ammortizationSheet.Range("H3").Offset(i, 0).Value = ammortizationSheet.Range("F3").Offset(i, 0).Value - ammortizationSheet.Range("G3").Offset(i, 0).Value ' principal repayment
        ammortizationSheet.Range("J3").Offset(i, 0).Value = principalOutstanding - ammortizationSheet.Range("H3").Offset(i, 0).Value ' principal outstanding end
        principalOutstanding = ammortizationSheet.Range("J3").Offset(i, 0).Value ' update for next period
        counter = counter + 1
    Next i

     ' dynamically add border lines to the table
    For i = 0 To steps - 1
        ammortizationSheet.Range("D3").Offset(i, 0).Borders.LineStyle = xlContinuous
        ammortizationSheet.Range("E3").Offset(i, 0).Borders.LineStyle = xlContinuous
        ammortizationSheet.Range("F3").Offset(i, 0).Borders.LineStyle = xlContinuous
        ammortizationSheet.Range("G3").Offset(i, 0).Borders.LineStyle = xlContinuous
        ammortizationSheet.Range("H3").Offset(i, 0).Borders.LineStyle = xlContinuous
        ammortizationSheet.Range("I3").Offset(i, 0).Borders.LineStyle = xlContinuous
        ammortizationSheet.Range("J3").Offset(i, 0).Borders.LineStyle = xlContinuous
    Next i

    ' dynamically add background color to the table
    For i = 0 To steps - 1
        ammortizationSheet.Range("D3").Offset(i, 0).Interior.ColorIndex = 15
        ammortizationSheet.Range("E3").Offset(i, 0).Interior.ColorIndex = 15
        ammortizationSheet.Range("F3").Offset(i, 0).Interior.ColorIndex = 15
        ammortizationSheet.Range("G3").Offset(i, 0).Interior.ColorIndex = 15
        ammortizationSheet.Range("H3").Offset(i, 0).Interior.ColorIndex = 15
        ammortizationSheet.Range("I3").Offset(i, 0).Interior.ColorIndex = 15
        ammortizationSheet.Range("J3").Offset(i, 0).Interior.ColorIndex = 15
    Next i

    ' autofit the data
    ammortizationSheet.Range("D:J").EntireColumn.AutoFit
    ammortizationSheet.Range("H:J").NumberFormat = """R"" #,##0.00"
    ammortizationSheet.Columns("F").NumberFormat = """R"" #,##0.00"
End Sub

Public Sub ClearAmmortization()
    ' define sheet
    Dim ammortizationSheet as Worksheet

    ' set it up
    Set ammortizationSheet = Worksheets("ammortization")

    With ammortizationSheet
        .Cells.Clear
        .Cells.UnMerge
    End With
End Sub