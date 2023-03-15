Attribute VB_Name = "StockData"
Sub StockData()

Dim Ticker As String
Dim Tickertotal, OpenCount, EndCount, Percent As Double
Dim Summary_Table_Row As Integer


Summary_Table_Row = 2
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Tickertotal = 0
VolumeCount = 0

     For i = 2 To Lastrow
 
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
        VolumeCount = Cells(i, 7).Value + VolumeCount
        EndCount = Cells(i, 6).Value
        
        Ticker = Cells(i, 1).Value
        Cells(Summary_Table_Row, 11).Value = Ticker
        
        Tickertotal = EndCount - OpenCount
        Percent = Tickertotal / OpenCount
    
        Cells(Summary_Table_Row, 12).Value = Tickertotal
        Cells(Summary_Table_Row, 13).Value = Percent
        Cells(Summary_Table_Row, 14).Value = VolumeCount
        
           If Tickertotal > 0 Then
           Cells(Summary_Table_Row, 12).Interior.ColorIndex = 4
        
           Else
           Cells(Summary_Table_Row, 12).Interior.ColorIndex = 3
        
           End If
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Tickertotal = 0
        VolumeCount = 0
        
        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1) Then
        
        OpenCount = Cells(i, 3).Value
        
        VolumeCount = Cells(i, 7).Value + VolumeCount
        
        Else
        
        VolumeCount = Cells(i, 7).Value + VolumeCount
        
        End If
        
        
    Next i



End Sub
