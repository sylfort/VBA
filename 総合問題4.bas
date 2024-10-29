Option Explicit

Function 表面利回り(kakaku As Long, yachin As Long) As Single

    表面利回り = Round(yachin * 12 / kakaku * 100, 1)

End Function
