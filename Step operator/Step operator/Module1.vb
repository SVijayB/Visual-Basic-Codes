Module Module1
    Sub Main()
        Console.WriteLine("All the numbers bellow are normal which starts from 1 and ends with 20")
        For int1 = 1 To 20
            Console.WriteLine(int1)
        Next
        Console.WriteLine()

        Console.WriteLine("All the numbers bellow are odd numbers from 1 to 20.")
        For int1 = 1 To 20 Step 2
            Console.WriteLine(int1)
        Next
        Console.WriteLine()
        Console.WriteLine("All the numbers bellow are even numbers from 1 to 20.")
        For int1 = 0 To 20 Step 2
            Console.WriteLine(int1)
        Next
        Console.ReadLine()
    End Sub
End Module
