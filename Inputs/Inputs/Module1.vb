Module Module1

    Sub Main()
        Dim Username As String = Nothing
        Dim Userage As Integer = Nothing
        Dim Usergame As String = Nothing
        Dim Userhobbies As String = Nothing
        Dim Usersubject As String = Nothing
        Dim Userlike As String = Nothing

        Console.WriteLine("What is your name ? ")
        Username = Console.ReadLine()

        Console.WriteLine("What is your age ? ")
        Userage = Console.ReadLine()

        Console.WriteLine("What Game or Games do you like the most ? ")
        Usergame = Console.ReadLine()

        Console.WriteLine("What is or what are your hobbies ? ")
        Userhobbies = Console.ReadLine()

        Console.WriteLine("Which subject or subjects do you like the most ? ")
        Usersubject = Console.ReadLine()

        Console.WriteLine("Who do you like the most in your life ? ")
        Userlike = Console.ReadLine()

        Console.WriteLine("Your name is " & Username & "," & " Your age is " & Userage & "," & " The game or games you like the most is " & Usergame & "," & " Your Hobbie is to " & Userhobbies & "," & " The Subject you like the most is " & Usersubject & " and" & " The person you like the most is " & Userlike & ".")
        Console.ReadLine()
    End Sub

End Module
