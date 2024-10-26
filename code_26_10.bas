Option Explicit

Sub kaisuu()
    Dim i As Integer
    For i = 1 To 3
        MsgBox i & "回目の実行です"
    Next
End Sub

Sub zouka()
    Dim num As Integer
    Dim i As Integer
    Dim cur_cell As String
    num = 100
    For i = 0 To 4
        cur_cell = Range("C15").Select
        Selection.Offset(0, i).Value = num
        num = num + 50
    Next
End Sub


Sub shozoku()
    Dim busho As String
    Select Case Range("C8").Value
        Case 100
            busho = "総務部"
        Case 200
            busho = "人事部"
        Case 300
            busho = "営業部"
        Case 400
            busho = "企画部"
        Case 500
            busho = "開発部"
        Case Else
            busho = "正しいコードを入力して下さい。"
    
    End Select
    MsgBox busho
End Sub
    
Sub point()
    Dim shouhin As String
    Select Case Range("C17").Value
        Case Is >= 1000
            shouhin = "商品券"
        Case 800 To 999
            shouhin = "ギフトカタログ"
        Case 500 To 799
            shouhin = "入浴剤"
        Case 200 To 499
            shouhin = "タオル"
        Case Else
            shouhin = "対応する商品はありません。"
    
    End Select
    MsgBox shouhin
End Sub
    
Sub kubun()
    Dim taipu As String
    Select Case Range("C25").Value
        Case 100, 110, 120
            taipu = "乗用車"
        Case 201, 211, 221
            taipu = "RV・4WD"
        Case 300, 305, 310
            taipu = "スポーツカー"
        Case Else
            taipu = "正しいコードを入力して下さい。"
    
    End Select
    MsgBox taipu
End Sub

Sub iro()
    Dim iro As String
    iro = Trim(StrConv(Range("C32").Value, vbUpperCase))
    
    Select Case iro
        Case "RED"
            Range("C32").Font.Color = vbRed
        Case "BLUE"
            Range("C32").Font.Color = vbBlue
        Case "PINK"
            Range("C32").Font.Color = vbMagenta
        Case "GREEN"
            Range("C32").Font.Color = vbGreen
        Case Else
            MsgBox "正しい色を入力して下さい。"
    
    End Select
    
End Sub

Sub loop1()
    Dim cur_cell As String
    cur_cell = ActiveCell.Select
    
    Dim i As Integer
    i = 1
    
    Do While ActiveCell.Value <> ""
        MsgBox i & "回目の実行です"
        i = i + 1
        ActiveCell.Offset(0, 1).Select
    Loop

End Sub

Sub loop2()
    Dim cur_cell As String
    cur_cell = ActiveCell.Select
    
    Dim i As Integer
    i = 1
    
    Do
        MsgBox i & "回目の実行です"
        i = i + 1
        ActiveCell.Offset(0, 1).Select
    
    Loop While ActiveCell.Value <> ""

End Sub

