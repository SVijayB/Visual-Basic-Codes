Module Module1

    Sub Main()
        Console.WriteLine("Please enter the Base to power")
        Dim num1 As Double = Console.ReadLine()
        Console.WriteLine("The Base is : " & num1)
        Console.WriteLine("Please enter the power to the base")
        Dim num2 As Double = Console.ReadLine()
        Console.WriteLine("The power to the base is : " & num2)
        Console.WriteLine(num1 & " To the power " & num2 & " is : " & num1 ^ num2)
        Console.ReadLine()
    End Sub

End Module
