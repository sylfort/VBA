Option Explicit
Public atari As Integer
Public msg1 As String
Public msg2 As String
Public msg3 As String

Private Sub tgl01_Click()
If tgl01.Value = True Then
    tgl01.ForeColor = vbRed
    CheckWinner
    MsgBox msg1
    tgl01.Value = False
Else
    tgl01.ForeColor = &H8000000D
End If
End Sub
Private Sub tgl02_Click()
If tgl02.Value = True Then
    tgl02.ForeColor = vbRed
    CheckWinner
    MsgBox msg2
    tgl02.Value = False
Else
    tgl02.ForeColor = &H8000000D
End If
End Sub
Private Sub tgl03_Click()
If tgl03.Value = True Then
    tgl03.ForeColor = vbRed
    CheckWinner
    MsgBox msg3
    tgl03.Value = False
Else
    tgl03.ForeColor = &H8000000D
End If
End Sub

Sub CheckWinner()
    atari = Int(Rnd() * 3) + 1
    
    If atari = 1 Then
        msg1 = "勝った！"
        msg2 = "負けった…"
        msg3 = "負けった…"
    ElseIf atari = 2 Then
        msg1 = "負けった…"
        msg2 = "勝った！"
        msg3 = "負けった…"
    Else
        msg1 = "負けった…"
        msg2 = "負けった…"
        msg3 = "勝った！"   
    End If
    
End Sub
