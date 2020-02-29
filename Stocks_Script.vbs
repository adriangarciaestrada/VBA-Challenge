Sub Button1_Click()

For Each ws In Worksheets

    ws.Activate

    Tabcont = 2
    
    TikVol = 0

    TikOpen = 0
    
    TikClose = 0
    
    Cells(1, 10) = "Ticker"

    Cells(1, 11) = "Yearly Change"

    Cells(1, 12) = "Percentage Change"

    Cells(1, 13) = "Total Stock Volume"

    FRCount = Cells(Rows.Count, 1).End(xlUp).Row


    For a = 2 To (FRCount + 1)

        If TikOpen = 0 Then
    
            TikOpen = Cells(a, 3)
    
        End If
    
        TikVol = TikVol + Cells(a, 7)
        
        If Cells(a, 1) <> Cells(a + 1, 1) Then
    
            TikClose = Cells(a, 6)
    
            Cells(Tabcont, 10) = Cells(a, 1)
        
            Cells(Tabcont, 11) = TikClose - TikOpen
        
            If Cells(Tabcont, 11) < 0 Then
        
                Cells(Tabcont, 11).Interior.ColorIndex = 3
            
            Else
        
                Cells(Tabcont, 11).Interior.ColorIndex = 4
        
            End If
            
            If TikOpen <> 0 Then
                    
                Cells(Tabcont, 12) = (TikClose - TikOpen) / TikOpen
                
            Else
            
                Cells(Tabcont, 12) = 0
            
            End If
        
            Cells(Tabcont, 12).Style = "Percent"
        
            Cells(Tabcont, 13) = TikVol
        
            Cells(Tabcont, 13).Style = "Currency"
        
            TikOpen = 0
        
            TikClose = 0
    
            TikVol = 0
        
            Tabcont = Tabcont + 1
    
        End If

    Next a

    SRCount = Cells(Rows.Count, 10).End(xlUp).Row

    MinPerChange = Cells(2, 12)

    MaxPerChange = Cells(2, 12)

    MaxTotVolume = Cells(2, 13)

    For b = 3 To SRCount

        If MinPerChange > Cells(b, 12) Then
    
            MinPerChange = Cells(b, 12)
        
            MinPerTik = Cells(b, 10)
                        
        End If
    
        If MaxPerChange < Cells(b, 12) Then
    
            MaxPerChange = Cells(b, 12)
        
            MaxPerTik = Cells(b, 10)
    
        End If
    
        If MaxTotVolume < Cells(b, 13) Then
    
            MaxTotVolume = Cells(b, 13)
        
            MaxTotTik = Cells(b, 10)
    
        End If

    Next b

    Cells(1, 16) = "Ticker"

    Cells(1, 17) = "Value"

    Cells(2, 15) = "Greatest % Increase"

    Cells(2, 16) = MaxPerTik

    Cells(2, 17) = MaxPerChange

    Cells(2, 17).Style = "Percent"

    Cells(3, 15) = "Greatest % Decrease"

    Cells(3, 16) = MinPerTik

    Cells(3, 17) = MinPerChange

    Cells(3, 17).Style = "Percent"

    Cells(4, 15) = "Greatest Total Volume"

    Cells(4, 16) = MaxTotTik

    Cells(4, 17) = MaxTotVolume

    Cells(4, 17).Style = "Currency"
    
Next ws

End Sub
