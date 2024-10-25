Option Explicit

Sub 練習2_1()
    
    Dim syurui As String
    syurui = Range("G5").Value
    
    Range("B11:E16").Select
    
    Select Case syurui
        Case "縦棒"
            ActiveSheet.Shapes.AddChart(51, xlColumnClustered).Select
        Case "横棒"
            ActiveSheet.Shapes.AddChart(57, xlBarClustered).Select
        Case "折れ線"
            ActiveSheet.Shapes.AddChart(4, xlLine).Select
        Case "折れ線"
            ActiveSheet.Shapes.AddChart(1, xlArea).Select
    End Select
    
    ActiveChart.SetSourceData Source:=Range("練習問題2!$B$11:$E$16")
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="グラフ"
    
End Sub

Sub 練習2_2()
    If Sheets("グラフ").Name = "グラフ" Then
        Sheets("グラフ").Select
        ActiveWindow.SelectedSheets.Delete
    End If
    
End Sub
