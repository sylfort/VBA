Sub 入力()
'
' 入力 Macro
'
    Dim hiduke As String
    
    ActiveWindow.NewWindow
    Windows.Arrange ArrangeStyle:=xlVertical
    Windows("第6章_A.xlsm:1").Activate
    Sheets("得意先リスト").Select
    Windows("第6章_A.xlsm:2").Activate
    
    If Range("C5").Value = "" Then
        Range("C5").Select
    Else
        Range("C4").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
    End If
    
    Do While ActiveCell.Offset(0, -1).Value <> ""
        hiduke = InputBox("日付を入力してください" & Chr(10) & _
        "入力を終了する場合には半角でendと入力します", , , 200, 200)
        
        If hiduke = "end" Or hiduke = "" Then
            Exit Do
        Else
            ActiveCell.Value = hiduke
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = InputBox("得先別コードを入力してください", , , 200, 200)
            Windows("第6章_A.xlsm:1").Activate
            Sheets("商品リスト").Select
            Windows("第6章_A.xlsm:2").Activate
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = InputBox("商品コードを入力してください", , , 200, 200)
            Application.ScreenUpdating = True
            ActiveCell.Offset(0, 4).Select
            ActiveCell.Value = InputBox("数量を入力してください", , , 200, 200)
            ActiveCell.Offset(1, -7).Select
            Windows("第6章_A.xlsm:1").Activate
            Sheets("得意先リスト").Select
            Windows("第6章_A.xlsm:2").Activate
        End If
    Loop

    Windows("第6章_A.xlsm:1").Activate
    ActiveWindow.Close
'    ActiveWindow.WindowState = xlMaximized
    Range("A1").Select
End Sub

Sub 新規シート()
'
' 新規シート Macro
'
    Dim tuki As String
    ActiveSheet.Copy after:=ActiveSheet
    Range("C5:D14,F5:F14,J5:J14").Select
    Selection.ClearContents
    Range("A1").Select
    
    tuki = InputBox("得先別コードを入力してください", , , 200, 200)
    If tuki = "" Then
        Application.DisplayAlerts = False
        ActiveSheet.delete
        Application.DisplayAlerts = True
    Else
        On Error Resume Next
        ActiveSheet.Name = tuki & "月度"
        If Err.Number = 1004 Then
            MsgBox "シート名が重複します"
            Application.DisplayAlerts = False
            ActiveSheet.delete
            Application.DisplayAlerts = True
        End If
    End If
End Sub


Sub 月別販売データへ()
    Dim mysheet As String
    mysheet = InputBox("売上月を半角の数字でを入力してください")
    On Error Resume Next
    Worksheets(mysheet & "月度").Select
    Range("A1").Select
    If Err.Number = 9 Then
        MsgBox "シートがありません"
    End If
End Sub

Sub 終了()
    Dim endex As Integer
    endex = MsgBox("Excelを終了します。よろしいですか？", vbOKCancel, "終了")
    If endex = vbOK Then
        ActiveWorkbook.Save
        Application.Quit
    End If
End Sub
