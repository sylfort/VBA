Option Explicit

Sub シート保護解除()
    Dim true_pass As String
    true_pass = "password"
    Dim pass_try As String
    Dim chances As Integer
    chances = 3
    Dim i As Integer
    For i = 1 To chances
            pass_try = InputBox("パスワードを入力（大文字小文字を）", "パスワード入力", , 200, 200)
    
    If pass_try = true_pass Then
        MsgBox ("シート保護解除します")
        ActiveSheet.Unprotect (true_pass)
        Exit For
    Else
        If i = 3 Then
            Exit For
        End If
        MsgBox ("パスワードが違います" & Chr(10) & "残り" & (chances - i) & "回のチャンス")
    End If
    Next i

End Sub

Sub シート保護()
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
