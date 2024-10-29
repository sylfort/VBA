Option Explicit

Sub 日報入力()
    If Range("B5").Value = "" Then
        Range("B5").Select
    Else
        Range("B4").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
    End If

    Dim hidzuke As String
    
    Dim topics As Variant
    
    
    Dim message As String
    message = "を入力してください"
    
    topics = Application.Transpose(Application.Transpose(Range("B4:G4")))

    Dim i As Integer
    For i = 1 To 6
        
    Next i
    
    Dim element As Variant
    For Each element In topics
        If element = "曜日" Then
            ActiveCell.Value = WeekdayName(Weekday(hidzuke), True)
            
            ActiveCell.Offset(0, 1).Select
        GoTo MyLabel
        End If
    
        hidzuke = InputBox(element & message, , , 200, 200)
        If hidzuke = "end" Or hidzuke = "" Then
            Exit Sub
        End If
            
            ActiveCell.Value = hidzuke
            ActiveCell.Offset(0, 1).Select
               
MyLabel:
    Next element
    
    'Range("F4").Style = "Currency [0]"
    Range("F4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "\#,##0;[赤]\-#,##0"
    Range("B4").Select
    Selection.End(xlDown).Offset(1, 0).Select
    
End Sub


Sub 日報入力2()
    If Range("B5").Value = "" Then
        Range("B5").Select
    Else
        Range("B4").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
    End If

    Dim hidzuke As String
    
    Do
        hidzuke = InputBox("今日の日付を入力してください", , , 200, 200)
        If hidzuke = "end" Or hidzuke = "" Then
            Exit Sub
        Else
            ActiveCell.Value = hidzuke
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = WeekdayName(Weekday(hidzuke), True)
            
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = InputBox("天候を入力してください", , , 200, 200)
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = InputBox("来場者数を入力してください", , , 200, 200) & "人"
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = InputBox("売上金額を入力してください", , , 200, 200)
            Selection.Style = "Currency"
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = InputBox("担当者名を入力してください", , , 200, 200)
            ActiveCell.Offset(1, -5).Select
        End If
    Loop While ActiveCell.Offset(-1, 0).Value <> ""
    
End Sub
