Module Module1

    Sub Main()
        Dim num1 As Integer = 66
        Dim num2 As Integer = 7
        Dim answer As Integer = Nothing
        answer = num1 + num2
        Console.WriteLine(answer)

        Dim num3 As Integer = 10
        Dim num4 As Integer = 13
        Dim answer1 As Integer = Nothing
        answer1 = num3 - num4
        Console.WriteLine(answer1)

        Dim num5 As Integer = 3
        Dim num6 As Integer = 9
        Dim answer2 As Integer = Nothing
        answer2 = num5 * num6
        Console.WriteLine(answer2)

        Dim num7 As Integer = 4
        Dim num8 As Integer = 8
        Dim answer3 As Double = Nothing
        answer3 = num7 / num8
        Console.WriteLine(answer3)

        Dim num9 As Integer = 2
        Dim num10 As Integer = 4
        Dim answer4 As Integer = Nothing
        answer4 = num9 ^ num10
        Console.WriteLine(answer4)

        Dim num11 As Integer = 31
        Dim num12 As Integer = 4
        Dim answer5 As Integer = Nothing
        answer5 = num11 Mod num12
        Console.WriteLine(answer5)

        Console.ReadLine()
    End Sub

End Module
