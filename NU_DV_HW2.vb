Sub WallStreet()
Dim k As Double
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim lastCol As Long
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
Dim x As Long
x = 2
k = 0
Dim last As Double
Dim start As Double
Dim i As Long
Dim max As Long
Dim min As Long
Dim total As Long
Dim most As Double
most = 0
start = Cells(2, 6)
    For i = 2 To lastRow
        If Cells(i, 1) = Cells(i + 1, 1) Then
        k = k + Cells(i, 7)
        Else
        Cells(x, 12) = k
        k = 0
        Cells(x, 9) = Cells(i, 1)
        last = Cells(i, 6)
        Cells(x, 10) = last - start
            If start = 0 Then
            Cells(x, 11) = 0
            Else
            Cells(x, 11) = (last - start) / start
            End If
        last = 0
        x = x + 1
        start = Cells(i + 1, 6)
        End If
    Next i
    For i = 2 To 2836
        If Cells(i, 11) < Cells(i + 1, 11) Then
        max = Cells(i + 1, 11)
        End If
        Cells(2, 16) = max
    Next i
    For i = 2 To 2835
        If Cells(i, 11) > Cells(i + 1, 11) Then
        min = Cells(i + 1, 11)
        End If
        Cells(3, 16) = min
    Next i
    For i = 2 To 2836
        If Cells(i, 12) < Cells(i + 1, 12) Then
        most = Cells(i + 1, 12)
        End If
        Cells(4, 16) = most
    Next i
    For i = 2 To 2836
        If Cells(i, 11) = Cells(2, 16) Then
            Cells(2, 15) = Cells(i, 9)
        End If
    Next i
    For i = 2 To 2836
        If Cells(i, 11) = Cells(3, 16) Then
            Cells(3, 15) = Cells(i, 9)
        End If
    Next i
    For i = 2 To 2836
        If Cells(i, 12) = Cells(4, 16) Then
            Cells(4, 15) = Cells(i, 9)
        End If
    Next i
        
End Sub
