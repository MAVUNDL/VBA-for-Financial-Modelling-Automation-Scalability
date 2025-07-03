Attribute VB_Name = "VBAAnnuity"

Public sub createIOTable()
    ' define sheet
    Dim annuitySheet as Worksheet

    ' set it up
    Set annuitySheet = Worksheets("annuity")

    ' designing table 
    With annuitySheet.Range("A1:B1")
        .Merge
        .Value = "Inputs"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
    End With

    With annuitySheet.Range("A7:B7")
        .Merge
        .Value = "Outputs"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
    End With

    With annuitySheet
        Range("A2") = "Amount to be paid"
        Range("A3") = "Nominal interest (p.a)"
        Range("A4") = "Frequency of payments per year"
        Range("A5") = "Term in Years"
        Range("D1") = "Annuity Type"
        Range("A2:A5,B2:B5,D1:E1").Borders.LineStyle = xlContinuous ' add borders
        Range("A2:A5").Interior.ColorIndex = 23
        Range("D1").Interior.ColorIndex = 10
        Range("E1").Interior.ColorIndex = 44
        Range("A2:A5").Font.ColorIndex = 2
        Range("B2:B5").Interior.ColorIndex = 20
        Range("B2:B5").Font.ColorIndex = 1

        Range("A8") = "Total number of payments"
        Range("A9") = "Effective interest rate"
        Range("A10") = "Present value (Payements in Advance)"
        Range("A11") = "Present value (Payments in Arrears)"
        Range("A12") = "Perpetuity (in Advance)"
        Range("A13") = "Perpetuity (in Arrears)"
        Range("A8:A13,B8:B13").Borders.LineStyle = xlContinuous ' add borders
        Range("A8:A13").Interior.ColorIndex = 23
        Range("A8:A13").Font.ColorIndex = 2
        Range("B8:B13").Interior.ColorIndex = 20
        Range("B8:B13").Font.ColorIndex = 1
        Range("E1").Validation.Delete
        Range("E1").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:= "Advance,Arrears" ' add drop down list
    End With
    ' autofit the data
    annuitySheet.Range("A:E").EntireColumn.AutoFit
End Sub


Public Sub CreateAutoAnnuityTable()
    ' setting up input variables
    Dim amountToBePaid as Variant
    Dim nominalInterest as Variant
    Dim frequency as Variant
    Dim term as Integer

    ' setting up variables for calculations
    Dim totalPayments as Integer
    Dim effectiveRate as Variant
    Dim paymentsAdvance as Double
    Dim paymentsArrears as Double

    ' define sheet
    Dim annuitySheet as Worksheet

    ' set it up
    Set annuitySheet = Worksheets("annuity")

    ' retrieve input data
    With annuitySheet
        amountToBePaid = Range("B2").Value
        nominalInterest = Range("B3").Value
        frequency = Range("B4").Value
        term = Range("B5").Value
    End With

    ' clear annnuity table
    annuitySheet.Range("H1:M" & annuitySheet.Rows.Count).Clear

    ' designing table 
    With annuitySheet.Range("H1:M1")
        .Merge
        .Value = "Annuity Table"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
    End With

    With annuitySheet
        .Range("H2") = "Year"
        .Range("I2") = "Amount paid"
        .Range("J2") = "Payment number"
        .Range("K2") = "Discount factor"
        .Range("L2") = "PV of amount Paid"
        .Range("M2") = "Total PV"
        .Range("H2:M2").Borders.LineStyle = xlContinuous ' add borders
        .Range("H2:M2").Interior.ColorIndex = 50 ' add background color
        .Range("H2:M2").Font.ColorIndex = 2 ' add font color
    End With

    'Compute outputs
    With annuitySheet
        .Range("B8") = frequency * term ' number of payments
        .Range("B9") = (nominalInterest  / 100) / frequency ' effective interest
        .Range("B10") = amountToBePaid * ((1 - (1 + Range("B9").Value)^(-Range("B8"))) / Range("B9").Value) * (1 + Range("B9").Value) ' present value - annuity due
        .Range("B11") = amountToBePaid * ((1 - (1 + Range("B9").Value)^(-Range("B8"))) / Range("B9").Value) ' present value - annuity ordinal
        .Range("B12") = (Range("B2").Value / Range("B9").Value) * (1 + Range("B9").Value) ' perpertuity of present value in advance
        .Range("B13") = Range("B2").Value / Range("B9").Value ' perpertuity of present value in arrears
    End With

    ' retrieve data
    With annuitySheet
        amountToBePaid = Range("B2").value
        effectiveRate = Range("B9").Value
    End With

    ' working calculations for the table
    Dim steps As Long: steps = frequency * term 
    Dim startVal As Double: startVal = 0
    Dim stepSize As Double : stepSize = 1 / frequency
    Dim iterator as Integer
    Dim i As Long
    Dim counter  As Integer : counter = 1
    Dim sumOfPayments As Variant : sumOfPayments = 0


    '
    If annuitySheet.Range("E1").Value = "Advance" Then
        startVal = 0
    Else
        startVal = stepSize
    End If

    ' dynamically calculate the  values and update the table
    For i = 0 To steps - 1
        annuitySheet.Range("H3").Offset(i, 0).Value = startVal + i * stepSize
        annuitySheet.Range("I3").Offset(i,0).Value = amountToBePaid
        annuitySheet.Range("J3").Offset(i,0).Value = counter 
        If annuitySheet.Range("E1").Value = "Advance" Then
            annuitySheet.Range("K3").Offset(i,0).Value = (1 / ((1 + effectiveRate)^counter)) * (1 + effectiveRate)
        Else
            annuitySheet.Range("K3").Offset(i,0).Value = (1 / ((1 + effectiveRate)^counter))
        End If
        annuitySheet.Range("L3").Offset(i,0).Value = amountToBePaid * annuitySheet.Range("K3").Offset(i,0).Value
        ' update counter
        counter = counter + 1
        ' sum values
        sumOfPayments = sumOfPayments +  annuitySheet.Range("L3").Offset(i,0).Value
    Next i

    ' update total present value in table
    annuitySheet.Range("M3").Value = sumOfPayments

    ' dynamically add border lines to the table
    For i = 0 To steps - 1
        annuitySheet.Range("H3").Offset(i, 0).Borders.LineStyle = xlContinuous
        annuitySheet.Range("I3").Offset(i,0).Borders.LineStyle = xlContinuous
        annuitySheet.Range("J3").Offset(i,0).Borders.LineStyle = xlContinuous
        annuitySheet.Range("K3").Offset(i,0).Borders.LineStyle = xlContinuous
        annuitySheet.Range("L3").Offset(i,0).Borders.LineStyle = xlContinuous
        annuitySheet.Range("M3").Borders.LineStyle = xlContinuous
    Next i

    ' dynamically add background color to the table
    For i = 0 To steps - 1
        annuitySheet.Range("H3").Offset(i, 0).Interior.ColorIndex = 15
        annuitySheet.Range("I3").Offset(i,0).Interior.ColorIndex = 15
        annuitySheet.Range("J3").Offset(i,0).Interior.ColorIndex = 15
        annuitySheet.Range("K3").Offset(i,0).Interior.ColorIndex = 15
        annuitySheet.Range("L3").Offset(i,0).Interior.ColorIndex = 15
        annuitySheet.Range("M3").Interior.ColorIndex = 15
    Next i

    ' autofit the data
    annuitySheet.Range("H:L").EntireColumn.AutoFit
End Sub

Public Sub cleanUp()
    ' define sheet
    Dim annuitySheet as Worksheet

    ' set it up
    Set annuitySheet = Worksheets("annuity")

    With annuitySheet
        .Cells.Clear
        .Cells.UnMerge
    End With
End Sub