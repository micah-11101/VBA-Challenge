Attribute VB_Name = "Module1"

Sub stockOutcome()
    Dim i As Long
    Dim Tcounter As Long
    Dim stockVolume As Double
    Dim Qstart As Double
    Dim Qend As Double
    Dim lastrow As Long
    Dim Olastrow As Integer
    Dim Gincrease As Double
    Dim Gdecrease As Double
    Dim Gvolume As Double
    Dim Tincrease As String
    Dim Tdecrease As String
    Dim Tvolume As String
    Dim ws As Integer
    
    
    ' vba script would run, but would not loop through my sheets, unless I defined a variable to count the sheets
    ws = Application.Worksheets.Count
    For x = 1 To ws
        Worksheets(x).Activate
    
    
    
        'setting variables/setting up the layout of spreadsheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        Tcounter = 1
        Gincrease = 0
        Gdecrease = 0
        Gvolume = 0
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To lastrow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            'Outputting Ticker based on name in separate row'
            Tcounter = Tcounter + 1
            Cells(Tcounter, 9).Value = Cells(i, 1).Value
            
            ' conditional statement for the first ticker
            If Qstart = 0 Then
            Qstart = Cells(2, 3)
            End If
            
            'quarter change value
            Qend = Cells(i, 6).Value
            Cells(Tcounter, 10).Value = Qend - Qstart
            
            'percent change value
            Cells(Tcounter, 11).Value = ((Qend - Qstart) / Qstart)
            'Convert the cell format to percentage
            Range("K" & Tcounter).NumberFormat = "0.00%"
            
            'new quarter start value
            Qstart = Cells(i + 1, 3).Value
            
            'Conditional Formatting for Quarter change
            If Cells(Tcounter, 10).Value > 0 Then
                Cells(Tcounter, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf Cells(Tcounter, 10).Value < 0 Then
                Cells(Tcounter, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            
            'Adds last row of volume to the variable before outputting it.
            'then resetting it to zero for the next Ticker.
            stockVolume = stockVolume + Cells(i, 7).Value
            Cells(Tcounter, 12).Value = stockVolume
            stockVolume = 0
        Else
            'if cell values are equal add the stock to total
            stockVolume = stockVolume + Cells(i, 7).Value
        End If
        
        
        
        Next i
        'Find last row of Outcomes
        Olastrow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'second for loop to find greatest values
        
        For i = 2 To Olastrow
        
        If Cells(i, 11) > Gincrease Then
            Gincrease = Cells(i, 11)
            Tincrease = Cells(i, 9)
            
        End If
        If Cells(i, 11) < Gdecrease Then
            Gdecrease = Cells(i, 11)
            Tdecrease = Cells(i, 9)
            
        End If
        
        If Cells(i, 12) > Gvolume Then
            Gvolume = Cells(i, 12)
            Tvolume = Cells(i, 9)
            
        End If
        
        Next i
        
        'greatest increase output
        Cells(2, 16).Value = Tincrease
        Cells(2, 17).Value = Gincrease
        Range("Q2").NumberFormat = "0.00%"
        'greatest decrease output
        Cells(3, 16).Value = Tdecrease
        Cells(3, 17).Value = Gdecrease
        Range("Q3").NumberFormat = "0.00%"
        
        'greatest volume output
        Cells(4, 16).Value = Tvolume
        Cells(4, 17).Value = Gvolume

    Next x
    
End Sub
