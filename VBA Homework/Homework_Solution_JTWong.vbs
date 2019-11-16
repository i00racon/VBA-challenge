Sub stock_volumn()

' Author: JoTing Wong


    For Each ws In Worksheets

        'Dim WorksheetName As String
        Dim Table1_LastRow As Long
        Dim Table2_LastRow As Long
        
        Dim total As Double
        Dim ticker As String
        Dim table_Row As Integer
        Dim OpenValue As Double
        Dim CloseValue As Double
        Dim Yr_Change As Double
        Dim Pct_Change As Double
        Dim GInc As Double
        Dim GInc_Ticker As String
        Dim GDec As Double
        Dim GDec_Ticker As String
        Dim GTotal As Double
        Dim GTotal_Ticker As String
        
        
            ' assign header
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volumn"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volumn"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            ' Determine the Last Row for data table
            Table1_LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
         'Initialize variables declared
            total = 0
            table_Row = 2
    
            ticker = ""
            OpenValue = 0
            CloseValue = 0
            Yr_Change = 0
            Pct_Change = 0
            GInc = 0
            GInc_Ticker = ""
            GDec = 0
            GDec_Ticker = ""
            GTotal = 0
            GTotal_Ticker = ""
        
            
            For i = 2 To Table1_LastRow
                ' find first open value
                If i = 2 Then
                    OpenValue = ws.Cells(i, 3).Value
                Else
                    OpenValue = OpenValue
                    
                End If
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        
                        ticker = ws.Cells(i, 1).Value
                        total = total + ws.Cells(i, 7).Value
                        ws.Range("I" & table_Row).Value = ticker
                        ws.Range("L" & table_Row).Value = total
                        
                        CloseValue = ws.Cells(i, 6).Value
                        Yr_Change = CloseValue - OpenValue
                        
                        'fix divide 0 error
                        If OpenValue = 0 Then
                            Pct_Change = 0
                        Else
                            Pct_Change = Yr_Change / OpenValue
                        End If
                        
                        ws.Range("J" & table_Row).Value = Yr_Change
                        ws.Range("K" & table_Row).Value = Pct_Change
                        ws.Range("K" & table_Row).NumberFormat = "0.00%"
                        
                        'Conditional format if change is positive green, negative red
                        If ws.Range("J" & table_Row).Value >= 0 Then
                            ws.Range("J" & table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & table_Row).Interior.ColorIndex = 3
                        End If
                        
                        OpenValue = ws.Cells(i + 1, 3).Value
                        table_Row = table_Row + 1
                        total = 0
                        ticker = ""
                        Pct_Change = 0
                        
                     Else
                        total = total + ws.Cells(i, 7).Value
                        
                        'if it's not the last value, keep the first open value
                        'OpenValue = OpenValue + 0
        
                
                    End If
                
                 
                Next i
            
        
            ' Determine the Last Row for summary table
            Table2_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            
            For x = 2 To Table2_LastRow
            
                ' get the greatest increase
                If ws.Cells(x, 11).Value > GInc Then
                   
                   GInc = ws.Cells(x, 11).Value
                   GInc_Ticker = ws.Cells(x, 9).Value
                           
                Else
                    
                   GInc = GInc
                   GInc_Ticker = GInc_Ticker
                    
                End If
            
                ' get the greatest decrease
                If ws.Cells(x, 11).Value > GDec Then
                   
                   GDec = GDec
                   GDec_Ticker = GDec_Ticker
                           
                Else
                    
                   GDec = ws.Cells(x, 11).Value
                   GDec_Ticker = ws.Cells(x, 9).Value
                    
                End If
                
                
                ' get the greatest total volumn
                If ws.Cells(x, 12).Value > GTotal Then
                   
                   GTotal = ws.Cells(x, 12).Value
                   GTotal_Ticker = ws.Cells(x, 9).Value
                           
                Else
                    
                   GTotal = GTotal
                   GTotal_Ticker = GTotal_Ticker
                    
                End If
            
            
            
            Next x
            
            ws.Range("P2").Value = GInc_Ticker
            ws.Range("Q2").Value = GInc
            ws.Range("P3").Value = GDec_Ticker
            ws.Range("Q3").Value = GDec
            ws.Range("P4").Value = GTotal_Ticker
            ws.Range("Q4").Value = GTotal
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
        
    Next ws
    
End Sub
