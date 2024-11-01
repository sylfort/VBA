Option Explicit

Dim danseiCell As String
Dim jyoseiCell As String
Dim danseiBlood As String
Dim jyoseiBlood As String
Dim selectedCell As String

Private Sub btn占う_Click()
    If OptionButton1 = True Then
        danseiCell = "3"
        danseiBlood = "A"
    End If
    If OptionButton2 = True Then
        danseiCell = "4"
        danseiBlood = "B"
    End If
    If OptionButton3 = True Then
        danseiCell = "5"
        danseiBlood = "O"
    End If
    If OptionButton4 = True Then
        danseiCell = "6"
        danseiBlood = "AB"
    End If
    
    If OptionButton5 = True Then
        jyoseiCell = "C"
        jyoseiBlood = "A"
    End If
    If OptionButton6 = True Then
        jyoseiCell = "D"
        jyoseiBlood = "B"
    End If
    If OptionButton7 = True Then
        jyoseiCell = "E"
        jyoseiBlood = "O"
    End If
    If OptionButton8 = True Then
        jyoseiCell = "F"
        jyoseiBlood = "AB"
    End If
    
    selectedCell = jyoseiCell & danseiCell
    
    Range(selectedCell).Select
       
    Range("C8").Value = danseiBlood & "型"
    Range("E8").Value = jyoseiBlood & "型"
    
    MsgBox ("男性" & danseiBlood & "型" & "女性" & jyoseiBlood & "型は" & Range(selectedCell).Value & "%ぐらいです")
    
    UserForm1.Hide
End Sub
