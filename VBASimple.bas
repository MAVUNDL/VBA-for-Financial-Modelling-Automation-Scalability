Attribute VB_Name = "VBASimple"

Public Sub SimpleInterest()
    ' setting up variables
    Dim simpleSheet as Worksheet
    Dim difference  as Integer
    Dim startDate as Date
    Dim endDate as Date
    Dim adjastedStart as Variant
    Dim adjastedEnd as Variant
    Dim countryCurrency as Variant
    Dim dayCount as Double
    Dim principalAmount as Variant
    Dim interestRate as Double
    Dim simpleInt as Double
    Dim futureValue as Double
    Dim presentValue as Double

    ' setting up sheet
    Set simpleSheet = Worksheets("simple")

    ' starting mods
    With simpleSheet
        ' formating these cells to be of a date format
        Range("B4:B7").NumberFormat = "yyyy/mm/dd"

        ' formating to percentage format
        Range("B9").NumberFormat = "0.00%"

        ' setting colors
        Range("A3").Interior.ColorIndex = 10
        Range("A3").Font.ColorIndex = 2
        Range("A15").Interior.ColorIndex = 10
        Range("A15").Font.ColorIndex = 2

        'setting borders
        Range("A3:A18,B3:B18,C4:C7").Borders.LineStyle = xlContinuous
        
        ' clear drop bar
        Range("B10").Validation.Delete

        ' adding drop-down list
        Range("B10").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:= "EUR,ZAR"

        ' Adding calculations
        'Range("B12") = Range("B6").Value - Range("B4").Value

        ' initialize variables
        startDate = Range("B4").Value
        adjastedStart = Range("B5").Value
        endDate = Range("B6").Value
        adjastedEnd =  Range("B7").Value
        countryCurrency = Range("B10").Value
        principalAmount = Range("B8").Value
        interestRate = Range("B9").Value


        '  start Conditionals
        Select Case Weekday(startDate, 2)
            Case Is = 5:
                adjastedStart = startDate + 8 - Weekday(startDate, 2)
            Case 6:
                adjastedStart = startDate + 8 - Weekday(startDate, 2)
            Case 7:
                adjastedStart = startDate + 8 - Weekday(startDate, 2)
            Case Else:
            adjastedStart = ""
        End Select

        ' end date conditional
        Select Case Weekday(endDate, 2)
            Case Is = 5:
                adjastedEnd = endDate + 8 - Weekday(endDate, 2)
            Case 6:
                adjastedEnd = endDate + 8 - Weekday(endDate, 2)
            Case 7:
                adjastedEnd = endDate + 8 - Weekday(endDate, 2)
            Case Else:
            adjastedEnd = ""
        End Select

        ' calculate the difference
        If adjastedStart <> "" And adjastedEnd <> "" Then
            difference = adjastedEnd - adjastedStart
        ElseIf adjastedStart <> "" And adjastedEnd = "" Then
            difference = endDate - adjastedStart  
        ElseIf adjastedEnd <> "" And adjastedStart = "" Then
            difference = adjastedEnd - startDate
        Else
            difference = endDate - startDate
        End If

        ' computing day count
        Select Case countryCurrency
            Case Is = "EUR" :
                dayCount = difference / 360
            Case "ZAR":
                dayCount = difference / 365
            Case Else:
            dayCount = 0
        End Select

        ' working out the calculations for the output
        simpleInt = principalAmount * interestRate * dayCount
        futureValue = principalAmount * (1 + interestRate * dayCount)
        presentValue = futureValue / (1 + interestRate * dayCount)



        ' update sheet
        Range("B12") = difference
        Range("B5") = adjastedStart
        Range("B7") = adjastedEnd
        Range("B13") = dayCount
        Range("B16") = simpleInt
        Range("B17") = futureValue
        Range("B18") = presentValue
        Range("C4") = WorksheetFunction.Text(startDate, "DDD")
        Range("C5") = WorksheetFunction.Text(adjastedStart, "DDD")
        Range("C6") = WorksheetFunction.Text(endDate, "DDD")
        Range("C7") = WorksheetFunction.Text(adjastedEnd, "DDD")

        ' conditions for color change
        If countryCurrency = "ZAR" Then
            Range("A4:A13").Interior.ColorIndex = 35
            Range("A16:A18").Interior.ColorIndex = 35
            Range("B4:B13").Interior.ColorIndex = 34
            Range("B16:B18").Interior.ColorIndex = 34
            Range("C4:C7").Interior.ColorIndex = 24
        ElseIf countryCurrency = "EUR" Then
            Range("A4:A13").Interior.ColorIndex = 44
            Range("A16:A18").Interior.ColorIndex = 44
            Range("B4:B13").Interior.ColorIndex = 19
            Range("B16:B18").Interior.ColorIndex = 19
            Range("C4:C7").Interior.ColorIndex = 48
        Else
            Range("A4:A13").Interior.ColorIndex = 2
            Range("A16:A18").Interior.ColorIndex = 2
            Range("B4:B13").Interior.ColorIndex = 2
            Range("B16:B18").Interior.ColorIndex = 2
            Range("C4:C7").Interior.ColorIndex = 2
        End If


    End With

End Sub

Public Sub clearContent()
    Range("B3:B18,C3:C7").ClearContents 
End Sub

Public Sub ClearFormating()
    Range("A3:A18,B3:B18,C3:C7").ClearFormats
End Sub