Sub priceSeperated()
Dim rng As Range
Dim r As Long
Dim arrParts() As String
Dim pricePartsGBP() As String
Dim pricePartssEUR() As String
Dim pricePartsUSD() As String
Dim partNum As Long
'## In my example i use columns A:E, and column D contains the Corresponding Parts ##

Set rng = Range("A1:BM10050") '## Modify as needed ##'

r = 2
Do While r <= rng.Rows.Count
    '## Split the value in column F (6) by commas, store in array ##
    arrParts = Split(rng(r, 4).Value, ",")

    '## Split the value in column P (16) by commas, store in array ##
    pricePartsUSD = Split(rng(r, 14).Value, ",")
    pricePartsEUR = Split(rng(r, 13).Value, ",")
    pricePartsGBP = Split(rng(r, 12).Value, ",")
    

    '## If there's more than one item in the array, add new lines ##
    If UBound(arrParts) >= 1 Then '## corrected this logic for base 0 array

        rng(r, 4).Value = arrParts(0) '# Size Seperate #
        rng(r, 14).Value = pricePartsUSD(0) '# USD Seperate #
        rng(r, 13).Value = pricePartsEUR(0) '# EUR Seperate #
        rng(r, 12).Value = pricePartsGBP(0) '# GBP Seperate #

        '## Iterate over the items in the array ##
        For partNum = 1 To UBound(arrParts)

            '## Insert a new row ##'
            '## increment the row counter variable ##
            r = r + 1
            rng.Rows(r).Insert Shift:=xlDown

            '## Copy the row above ##'
            rng.Rows(r).Value = rng.Rows(r - 1).Value

            '## update the part number in the new row ##'
            rng(r, 4).Value = Trim(arrParts(partNum))
            rng(r, 14).Value = Trim(pricePartsUSD(partNum))
            rng(r, 13).Value = Trim(pricePartsEUR(partNum))
            rng(r, 12).Value = Trim(pricePartsGBP(partNum))

            '## resize our range variable as needed ##
            Set rng = rng.Resize(rng.Rows.Count + 1, rng.Columns.Count)

        Next

    End If
'## increment the row counter variable ##
r = r + 1
Loop

End Sub
