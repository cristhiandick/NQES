Public Class Stats
    Public Function Square(ByVal x As Double) As Double
        Dim a As Double
        a = x * x
        Return a

    End Function
    Public Function StDevArray(ByVal x As Double()) As Double
        Dim a As Double = 0

        a = Var(x)
        a = Math.Sqrt(a)
        Return a

    End Function
    Public Function Average(ByRef Returns As Double()) As Double

        Dim dblTotRet As Double
        Dim x As Integer
        Dim dblLength# = UBound(Returns, 1)

        For x = 0 To dblLength
            dblTotRet += Returns(x)
        Next x

        Return dblTotRet / (dblLength + 1)

    End Function

    Public Function Var(ByVal InArray As Double()) As Double

        Dim dblAvgReturn, dblDeviation As Double
        Dim x As Integer

        Dim dblLength = UBound(InArray)

        dblAvgReturn = Average(InArray)

        For x = 0 To dblLength

            dblDeviation += (InArray(x) - dblAvgReturn) ^ 2

        Next x

        Return dblDeviation / dblLength

    End Function

    Public Function MVHR(ByVal a As Double(), ByVal b As Double(), ByVal c As Double()) As Double

        Dim d As Double
        Dim x As Integer

        d = StDevArray(a) * StDevArray(b) * Correlation(a, b) / Var(b)

        Return d

    End Function
    Public Function Skew(ByRef InArray As Double()) As Double

        Dim dblSkewSumm, dblAvgReturn, dblStdDev As Double
        Dim x As Integer

        Dim dblLength# = UBound(InArray)

        dblAvgReturn = Average(InArray)
        dblStdDev = Math.Sqrt(Var(InArray))

        For x = 0 To dblLength

            dblSkewSumm += ((InArray(x) - dblAvgReturn) / dblStdDev) ^ 3

        Next x

        Return (dblLength + 1) / ((dblLength) * (dblLength - 1)) * dblSkewSumm

    End Function
    Public Function Kurtosis(ByRef InArray As Double()) As Double

        Dim dblKurtSumm, dblAvgReturn, dblStdDev As Double
        Dim x As Integer

        Dim dblLength# = UBound(InArray)

        dblAvgReturn = Average(InArray)
        dblStdDev = Math.Sqrt(Var(InArray))

        For x = 0 To dblLength

            dblKurtSumm += ((InArray(x) - dblAvgReturn) / dblStdDev) ^ 4

        Next x

        Return ((dblLength + 1) * (dblLength + 2)) / ((dblLength) * (dblLength - 1) * (dblLength - 2)) * _
        dblKurtSumm - (3 * (dblLength ^ 2) / ((dblLength - 1) * (dblLength - 2)))

    End Function


    Public Function Correlation(ByRef a As Double(), ByRef b As Double()) As Double


        Dim sumX, sumY, sumXsquare, sumYsquare, sumXY, corr1, corr2, corrcoef As Double
        Dim x As Integer

        Dim dblLength# = UBound(a)


        For x = 0 To dblLength
            sumX += a(x)
            sumY += b(x)
            sumXsquare += a(x) * a(x)
            sumYsquare += b(x) * b(x)
            sumXY += a(x) * b(x)
        Next x

        corr1 = ((dblLength * sumXY) - (sumX * sumY))
        corr2 = Math.Sqrt(((dblLength * sumXsquare) - (sumX * sumX)) * ((dblLength * sumYsquare) - (sumY * sumY)))
        corrcoef = corr1 / corr2


        Return corrcoef

    End Function

End Class