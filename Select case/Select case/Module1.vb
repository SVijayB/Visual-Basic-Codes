Module Module1

    Sub Main()
        Dim MyInt As String = Nothing
        Console.WriteLine("Please enter any number from 1 to 5")
        MyInt = Console.ReadLine()
        Select Case MyInt.ToLower()
            Case 1
                Console.WriteLine("Hello!")
            Case 2
                Console.WriteLine("Your going to face a great trap in your life so be carefull")
            Case 3
                Console.WriteLine("You are going to be rich in a very few days so be happy and enjoy!")
            Case 4
                Console.WriteLine("The world is going to laugh at you and you will be chased like a dog so be carefull")
            Case 5
                Console.WriteLine("Goodbye!")
            Case Else
                Console.WriteLine("Sorry! the number you typed is out of limit so please try again with numbers between 1 to 5")
        End Select
        Console.ReadLine()
    End Sub

End Module
