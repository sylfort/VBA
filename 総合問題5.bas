Option Explicit

Function 金種計算(kingaku As Long, moneyCell As Range) As Long
        Dim prev As Variant
        Dim money As Variant
        
        prev = moneyCell.Offset(0, -1)
        money = moneyCell.Value
        
        Select Case money
        Case 10000
            金種計算 = Int((kingaku) / money)
        Case 5000
            金種計算 = Int((kingaku Mod prev) / money)
        Case 1000
            金種計算 = Int((kingaku Mod prev) / money)
        Case 500
            金種計算 = Int((kingaku Mod prev) / money)
        Case 100
            金種計算 = Int((kingaku Mod prev) / money)
        Case 50
            金種計算 = Int((kingaku Mod prev) / money)
        Case 10
            金種計算 = Int((kingaku Mod prev) / money)
        Case 5
            金種計算 = Int((kingaku Mod prev) / money)
        Case 1
            金種計算 = Int((kingaku Mod prev) / money)

        End Select

End Function
