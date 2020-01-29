Public Class AccCalc
    Public Function TopCalculation(ByRef hit300 As Double, ByRef hit100 As Double, ByRef hit50 As Double) As Double

        Dim result As Double

        result = (50 * hit50) + (100 * hit100) + (300 * hit300)

        TopCalculation = result

    End Function

    Public Function BottomCalculation(ByRef hit300 As Double, ByRef hit100 As Double, ByRef hit50 As Double, ByRef miss As Double) As Double

        Dim result As Double

        result = 300 * (miss + hit50 + hit100 + hit300)

        BottomCalculation = result


    End Function
End Class
