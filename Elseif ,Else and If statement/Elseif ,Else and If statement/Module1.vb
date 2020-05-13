Module Module1

    Sub Main()
        Console.WriteLine("What is your name? ")
        Dim Username As String = Console.ReadLine()
        Console.WriteLine("What is your password")
        Dim Password As String = Console.ReadLine()
        If Username = "Vijaybalaji" Then
            Console.WriteLine("We are great to you " & Username & " !")
        ElseIf Username = "Aishvarya" Then
            Console.WriteLine("We are great to you " & Username & " !")
        ElseIf Username = "Anbusudha" Then
            Console.WriteLine("We are great to you " & Username & " !")
        ElseIf Username = "Sankara narayanan" Then
            Console.WriteLine("We are great to you " & Username & " !")
        Else
            Console.WriteLine("I Do not no who you are " & Username & " !")
        End If

        If (Password = "Lets Rock" Or Password = "Lets rock" Or Password = "lets Rock" Or Password = "lets rock") Then
            Console.WriteLine("We Grand You Permission to access our top secrets ! " & Username)
        ElseIf Password = "123456789" Then
            Console.WriteLine("You have entered the number password so you can be a member of our secret Agency ! " & Username)
        Else
            Console.WriteLine("Sorry! You have entered a wrong answer , but don't worry you can try again! " & Username)
        End If
        Console.ReadLine()
    End Sub

End Module
