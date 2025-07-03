Attribute VB_Name = "VBACompound"

Public Sub compoundInterest()
    ' Declaring variables
    Dim usedCurrency as Variant
    Dim startDate as Date
    Dim endDate as Date
    Dim adjastedStartDate as Variant
    Dim adjastedEndDate as Variant
    Dim spotDate as Variant
    Dim dayDifference as Integer
    Dim dayCount as Double
    Dim principalAmount as Variant
    Dim annualInterest as Variant
    Dim compoundingType as Variant
    Dim frequencyType as Variant
    Dim frequency as Variant
    Dim interestAmount as Double
    Dim futureValue as Double
    Dim presentValue as Double
    Dim compoundSheet as Worksheet

    ' setting up my work sheet
    Set compoundSheet = Worksheets("compound")

    ' start working
    With compoundSheet

        ' creating drop-down menu
        Range("B3").Validation.Delete
        Range("B16").Validation.Delete
        Range("B17").Validation.Delete
        Range("B3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:= "EUR,ZAR"
        Range("B16").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:= "Discrete,Continuous"
        Range("B17").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:= "Daily,Weekly,Monthly,Quarterly,Semi-annual,Annual"

        ' initialize variables
        startDate = Range("B4").Value
        endDate = Range("B5").Value
        spotDate = Range("B8").Value
        principalAmount = Range("B14").Value
        annualInterest = Range("B15").Value / 100
        usedCurrency = Range("B3").Value
        compoundingType = Range("B16").Value
        frequencyType = Range("B17").Value

        ' formatting
        Range("B4:B8").NumberFormat = "yyyy/mm/dd"
        

        '  adjusting the start date
        Select Case Weekday(startDate, 2)
            Case Is = 5:
                adjastedStartDate = startDate + 8 - Weekday(startDate, 2)
            Case 6:
                adjastedStartDate = startDate + 8 - Weekday(startDate, 2)
            Case 7:
                adjastedStartDate = startDate + 8 - Weekday(startDate, 2)
            Case Else:
            adjastedStartDate = ""
        End Select

        ' justing the end date 
        Select Case Weekday(endDate, 2)
            Case Is = 5:
                adjastedEndDate = endDate + 8 - Weekday(endDate, 2)
            Case 6:
                adjastedEndDate = endDate + 8 - Weekday(endDate, 2)
            Case 7:
                adjastedEndDate = endDate + 8 - Weekday(endDate, 2)
            Case Else:
            adjastedEndDate = ""
        End Select

        ' calculating the day difference
        If adjastedStartDate <> "" And adjastedEndDate <> "" Then
            dayDifference = adjastedEndDate - adjastedStartDate
        ElseIf adjastedStartDate <> "" And adjastedEndDate = "" Then
            dayDifference = endDate - adjastedStartDate  
        ElseIf adjastedEndDate <> "" And adjastedStartDate = "" Then
            dayDifference = adjastedEndDate - startDate
        Else
            dayDifference = endDate - startDate
        End If

        ' update spot date acccordingly
        If adjastedStartDate <> "" Then
            spotDate = adjastedStartDate
        Else
            spotDate = startDate
        End If

        ' adjausting the denominator based on the currency
        Select Case usedCurrency
            Case Is = "EUR":
                denominator = 360
            Case Else:
                denominator = 365
        End Select

        ' update day count
        dayCount = dayDifference / denominator

        'adjasting the frequency
        Select Case frequencyType
            Case Is = "Daily" :
                frequency = 365
            Case "Weekly":
                frequency = 52
            Case "Monthly":
                frequency = 12
            Case "Quarterly":
                frequency = 4
            Case "Semi-annual":
                frequency = 2
            Case "Annual":
                frequency = 1
            Case Else:
                frequency = 0
        End Select
        
        ' working on calculations
        Select Case compoundingType
            Case Is = "Discrete":
                futureValue = principalAmount * (1 + (annualInterest / frequency))^(frequency * dayCount)
                presentValue = futureValue / ((1 + (annualInterest / frequency))^(frequency * dayCount))
                interestAmount = futureValue - presentValue
            Case Else:
                futureValue = principalAmount * exp(annualInterest * dayCount)
                presentValue = futureValue / (exp(annualInterest * dayCount))
                interestAmount = principalAmount * ((exp(annualInterest * dayCount)) - 1)
        End Select

        ' updating the sheet
        Range("B6") = adjastedStartDate
        Range("B7") = adjastedEndDate
        Range("B8") = spotDate
        Range("C4") = Format(startDate, "DDDD")
        Range("C5") = Format(endDate,  "DDDD")
        Range("C6") = Format(adjastedStartDate, "DDDD")
        Range("C7") = Format(adjastedEndDate, "DDDD")
        Range("C8") = Format(spotDate, "DDDD")
        Range("B10") = dayDifference
        Range("B11") = denominator
        Range("B12") = dayDifference / denominator
        Range("G3") = interestAmount
        Range("G4") = futureValue
        Range("G5") = presentValue

        ' Color formatting
        Select Case usedCurrency
            Case Is = "EUR":
            ' right side
                Range("A3").Interior.ColorIndex = 44
                Range("A17").Interior.ColorIndex = 44
                Range("A16").Interior.ColorIndex = 44
                Range("A3").Font.ColorIndex = 1
                Range("A17").Font.ColorIndex = 1
                Range("A16").Font.ColorIndex = 1
                Range("A4:A8").Interior.ColorIndex = 27
                Range("A10:A12").Interior.ColorIndex = 27
                Range("A14:A15").Interior.ColorIndex = 27

                ' left side
                Range("B3").Interior.ColorIndex = 45
                Range("B17").Interior.ColorIndex = 45
                Range("B16").Interior.ColorIndex = 45
                Range("B3").Font.ColorIndex = 1
                Range("B17").Font.ColorIndex = 1
                Range("B16").Font.ColorIndex = 1
                Range("B4:B8").Interior.ColorIndex = 35
                Range("B10:B12").Interior.ColorIndex = 35
                Range("B14:B15").Interior.ColorIndex = 35

                Range("C4:C8").Interior.ColorIndex = 40
                Range("C4:C8").Font.ColorIndex = 1

            Case Else:
            ' right side
            Range("A3").Interior.ColorIndex = 16
            Range("A17").Interior.ColorIndex = 16
            Range("A16").Interior.ColorIndex = 16
            Range("A3").Font.ColorIndex = 1
            Range("A17").Font.ColorIndex = 1
            Range("A16").Font.ColorIndex = 1
            Range("A4:A8").Interior.ColorIndex = 33
            Range("A10:A10").Interior.ColorIndex = 33
            Range("A14:A15").Interior.ColorIndex = 33
            Range("F3:F5").Interior.ColorIndex = 33

            ' left side
            Range("B3").Interior.ColorIndex = 15
            Range("B17").Interior.ColorIndex = 15
            Range("B16").Interior.ColorIndex = 15
            Range("B3").Font.ColorIndex = 1
            Range("B17").Font.ColorIndex = 1
            Range("B16").Font.ColorIndex = 1
            Range("B4:B8").Interior.ColorIndex = 20
            Range("B10:B12").Interior.ColorIndex = 20
            Range("B14:B15").Interior.ColorIndex = 20
            Range("G3:G5").Interior.ColorIndex = 20

            Range("C4:C8").Interior.ColorIndex = 48
            Range("C4:C8").Font.ColorIndex = 1
        End Select
    End With
End Sub

Public Sub createTables()
    ' worksheet
    Dim compoundSheet as Worksheet

    ' setting up my work sheet
    Set compoundSheet = Worksheets("compound")
    
    With compoundSheet
        ' Creating my headers
        Range("A3") = "Currency"
        Range("A4") = "Start Date"
        Range("A5") = "Maturity Date"
        Range("A6") = "Adjasted Start Date"
        Range("A7") = "Adjasted End Date"
        Range("A8") = "Spot Date"
        Range("A10") = "Day Difference"
        Range("A11") = "Denominator"
        Range("A12") = "Day Count in Years"
        Range("A14") = "Principal Amount"
        Range("A15") = "Annual Rate (%)"
        Range("A16") = "Compounding Type"
        Range("A17") = "Frequency"
        Range("F3") = "Amount with Interest"
        Range("F4") = "Future Value"
        Range("F5") = "Present Value"

        ' adding borders
        Range("A3:A17,B3:B17,C4:C8,F3:F5,G3:G5").Borders.LineStyle = xlContinuous

        ' setting coloring
        Range("A3").Interior.ColorIndex = 10
        Range("A17").Interior.ColorIndex = 10
        Range("A16").Interior.ColorIndex = 10
        Range("A3").Font.ColorIndex = 2
        Range("A17").Font.ColorIndex = 2
        Range("A16").Font.ColorIndex = 2

        Range("A4:A8").Interior.ColorIndex = 43
        Range("A10:A12").Interior.ColorIndex = 43
        Range("A14:A15").Interior.ColorIndex = 43
        Range("F3:F5").Interior.ColorIndex = 43

        compoundSheet.Range("A:G").EntireColumn.AutoFit
    End With
    
End Sub

Public Sub clearCompoundSheet()
    Range("A3:A17,B3:B17,C4:C8,F3:F5,G3:G5").ClearFormats
    Range("A3:A17,B3:B17,C4:C8,F3:F5,G3:G5").ClearContents
End Sub