Option Explicit

Public tensuu(3, 10) As Integer
Public jyuniArr(1 To 10) As Integer
Sub hairetsu()
    Dim i As Integer
    
    For i = 1 To 10
        tensuu(1, i) = Range("A" & i + 1).Value
        tensuu(2, i) = Range("B" & i + 1).Value
        'tensuu(3, i) = Range("A" & i + 1).Value + Range("B" & i + 1).Value
         
        'Debug.Print (tensuu(1, i) & " " & tensuu(2, i))
    Next i
    
End Sub
Sub 行プリント()
    Dim lineKokugo As String
    Dim lineSuugaku As String
    Dim i As Integer
    
    For i = 10 To 1 Step -1
        lineKokugo = lineKokugo & " " & tensuu(1, i)
        lineSuugaku = lineSuugaku & " " & tensuu(2, i)
    Next i
    Debug.Print lineKokugo
    Debug.Print lineSuugaku
End Sub
Sub goukei()
    Dim i As Integer
    
    For i = 1 To 10
        tensuu(3, i) = Range("A" & i + 1).Value + Range("B" & i + 1).Value
        Range("C" & i + 1).Value = tensuu(3, i)
    Next i
End Sub

Sub copyArr()
    Dim i As Integer
    For i = 1 To 10
        jyuniArr(i) = tensuu(3, i)
    Next i
End Sub

Sub scratchRank()
    Dim rankArr(1 To 10) As Integer
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To 10
        rankArr(i) = 1
        Range("D" & i + 1).Value = rankArr(i)
    Next i
    
    For i = 1 To 9
        For j = i + 1 To 10
            Debug.Print (jyuniArr(i) & " " & jyuniArr(j))
            If jyuniArr(i) > jyuniArr(j) Then
                rankArr(j) = rankArr(j) + 1
                Range("D" & j + 1).Value = rankArr(j)
            Else
                rankArr(i) = rankArr(i) + 1
                Range("D" & i + 1).Value = rankArr(i)
            End If
        Next j
    Next i
    
    For i = 1 To 10
        Debug.Print rankArr(i)
    Next i
    
    
End Sub


Sub bubbleSort()
    
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    Dim line As String

    For i = 1 To 10
        jyuniArr(i) = tensuu(3, i)
    Next i
    
    For i = 1 To 9
        temp = tensuu(3, i)
            
        For j = 1 To 9
            If jyuniArr(j + 1) < jyuniArr(j) Then
                temp = jyuniArr(j)
                jyuniArr(j) = jyuniArr(j + 1)
                jyuniArr(j + 1) = temp

            End If
        Next j

    Next i
       
    For i = 10 To 1 Step -1
        line = line & " " & jyuniArr(i)
    Next i
    
    Debug.Print line
    
End Sub
