Option Explicit

Sub 顧客名入力()
    Range("B4").Value = InputBox("顧客名を入力してください", "顧客名", , 200, 200) & "御中"
End Sub

Sub データ入力()
    
    If Range("C15").Value = "" Then
        Range("C15").Select
    Else
        Range("C14").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
    End If
    
    Selection.Value = InputBox("型番を入力します" & Chr(10) _
    & "終了する場合は*（アスタリスク）を入力してください", "型番", , 200, 200)
    
    If Selection.Value = "*" Or Selection.Value = "" Then
        Exit Sub
    End If
    
    Selection.Offset(0, 3).Select
    Selection.Value = InputBox("数量を入力します", "数量", , 200, 200)
    
    Selection.Offset(0, 2).Select
End Sub

Sub プレビュー()
    Range("B14").CurrentRegion.Select
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

Sub データ削除()
    Range("B4").ClearContents
    Range("C15:C29").ClearContents
    Range("F15:F29").ClearContents
    Range("A1").Select
End Sub
