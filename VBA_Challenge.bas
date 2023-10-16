VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_ticker():
Dim ws As Worksheet
Dim ticker As String
Dim vol As Variant
Dim i As Variant
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
Dim Greatest_Percent_Increase_Ticker As String
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease_Ticker As String
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Total_Volume_Ticker As String
Dim Greatest_Total_Volume As Double



       
       




For Each ws In Worksheets
   
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Yearly Change"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"


    Summary_Table_Row = 2
Greatest_Percent_Increase = 0
Greatest_Percent_Decrease = 99
Greatest_Total_Volume = 0

        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value

            year_open = ws.Cells(i, 3).Value
            year_close = ws.Cells(i, 6).Value

            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close

            
            ws.Cells(Summary_Table_Row, 8).Value = ticker
            ws.Cells(Summary_Table_Row, 9).Value = yearly_change
            ws.Cells(Summary_Table_Row, 10).Value = percent_change
            ws.Cells(Summary_Table_Row, 11).Value = vol
             If ws.Cells(Summary_Table_Row, 9).Value >= 0 Then
            ws.Cells(Summary_Table_Row, 9).Interior.ColorIndex = 4
             Else
            ws.Cells(Summary_Table_Row, 9).Interior.ColorIndex = 3
     
     
     
     
     End If
     
     
     
     If ws.Cells(Summary_Table_Row, 10).Value > Greatest_Percent_Increase Then
     Greatest_Percent_Increase = ws.Cells(Summary_Table_Row, 10).Value

            Greatest_Percent_Increase_Ticker = ws.Cells(Summary_Table_Row, 8).Value


                End If

            

            If ws.Cells(Summary_Table_Row, 10).Value < Greatest_Percent_Decrease Then
     Greatest_Percent_Decrease = ws.Cells(Summary_Table_Row, 10).Value
            
             Greatest_Percent_Decrease_Ticker = ws.Cells(Summary_Table_Row, 8).Value


                End If
'
            

            If ws.Cells(Summary_Table_Row, 11).Value > Greatest_Total_Volume Then
            Greatest_Total_Volume = ws.Cells(Summary_Table_Row, 11).Value

            Greatest_Total_Volume_Ticker = ws.Cells(Summary_Table_Row, 8).Value

                   
             End If
 
      
            Summary_Table_Row = Summary_Table_Row + 1
            
        ws.Cells(2, 13).Value = "Greatest_Percent_Increase"
        ws.Cells(3, 13).Value = "Greatest_Percent_Decrease"
        ws.Cells(4, 13).Value = "Greatest_Total_Volume"
        ws.Cells(1, 14).Value = "Ticker"
        ws.Cells(1, 15).Value = "Value"
       
            
'
        
        End If
        vol = vol + ws.Cells(i, 7).Value
        


    Next i
    
ws.Columns("g").NumberFormat = "0.00%"


     
     
     
      
     
     
  
   
   

    
        
        
                
        
                   
 
      

     
        ws.Cells(2, 14).Value = Greatest_Percent_Increase_Ticker
        ws.Cells(3, 14).Value = Greatest_Percent_Decrease_Ticker
        ws.Cells(4, 14).Value = Greatest_Total_Volume_Ticker
        ws.Cells(2, 15).Value = Greatest_Percent_Increase
        ws.Cells(3, 15).Value = Greatest_Percent_Decrease
        ws.Cells(4, 15).Value = Greatest_Total_Volume
        


Next ws





End Sub
'
'
