Module Module1

    Sub Main()
        Dim Username As String = Nothing
        Dim Password As String = Nothing
        Console.WriteLine("What is your Name?")
        Username = Console.ReadLine()
        Console.WriteLine("What is your Password? " & Username)
        Password = Console.ReadLine()

        If (Username = "Vijaybalaji" Or Username = "vijaybalaji" Or Username = "Vijay" Or Username = "vijay" Or Username = "Balaji" Or Username = "balaji" Or Username = "Aishvarya" Or Username = "aishvarya" Or Username = "Anbusudha" Or Username = "anbusudha" Or Username = "Sankara Narayanan" Or Username = "Sankara narayanan" Or Username = "sankara narayanan" Or Username = "Sankar" Or Username = "sankar") And (Password = "Rock on" Or Password = "rock on" Or Password = "Rock On" Or Password = "rock On") Then
            Console.WriteLine("We are great to you " & Username & "!" & " Your Password is Obsoletely Correct! ")
        Else
            Console.WriteLine("I don't know who you are " & Username & " so i'am sorry i can't check your Password")
        End If
        Console.ReadLine()
    End Sub

End Module
