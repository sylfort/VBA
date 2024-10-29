Option Explicit

Sub 未入荷リスト作成()

    Sheets("入荷待ちリスト").Range("B7").CurrentRegion.Clear

    Sheets("商品リスト").Range("B4:G50").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("入荷待ちリスト").Range("B3:B4"), CopyToRange:=Sheets("入荷待ちリスト").Range("B7"), Unique:=False
        
        Sheets("入荷待ちリスト").Select
End Sub
