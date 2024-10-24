Option Explicit

Sub 練習1()
    If Range("D10").Value = "" Then
        Range("E10").Value = "休み"
          
    ElseIf Range("C10").Value = "平日" Then
        Range("E10").Value = Range("C5").Value * Range("D10").Value
        
    ElseIf Range("C10").Value = "休日" Then
        Range("E10").Value = Range("C5").Value * Range("D10").Value

    End If
End Sub

Sub 練習3()

    Range("D6").Select
    
    Do Until Selection.Value = ""
        If Selection.Value >= 80 Then
            Selection.Offset(0, 1).Value = "合格"
        Else
            Selection.Offset(0, 1).Value = "不合格"
        End If
        Selection.Offset(1, 0).Select
    Loop
    
End Sub


Sub 練習4()
    Dim total As Long

    Range("C6").Select
    
    Do Until Selection.Value = ""
        total = total + ActiveCell.Value
        Selection.Offset(0, 1).Value = total
        Selection.Offset(1, 0).Select
    Loop
    
End Sub

