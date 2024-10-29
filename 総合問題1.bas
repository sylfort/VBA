Option Explicit
Sub メニューへ()
    Sheets("メニュー").Select
End Sub

Sub 第1四半期へ()
    Sheets("第1四半期").Select
End Sub

Sub 第2四半期へ()
    Sheets("第2四半期").Select
End Sub

Sub 上半期計へ()
    Sheets("上半期計").Select
End Sub

Sub 印刷プレビュー()
    Range("B4").CurrentRegion.Select
    Selection.Borders(xlBottom).Weight = xlThin
    Selection.Borders(xlLeft).Weight = xlThin
    Selection.Borders(xlRight).Weight = xlThin
    
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = 1
    End With
    ActiveSheet.PrintPreview
End Sub

Sub グラフ()
    ActiveSheet.Range("$B$4:$E$19").AutoFilter Field:=4, Criteria1:="3", _
        Operator:=xlTop10Items
    Range("C4:C18,E4:E18").Select
    Range("E4").Activate
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("上半期計!$C$4:$C$18,上半期計!$E$4:$E$18")
    Range("A1").Select
End Sub

Sub グラフ削除()
    ActiveSheet.ChartObjects.Delete
    ActiveSheet.Range("$B$4:$E$19").AutoFilter Field:=4
    Range("A1").Select
    'ActiveSheet.ChartObjects("上半期計").Activate
End Sub


