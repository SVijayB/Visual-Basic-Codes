Module Module1

    Sub Main()
        Dim Userstring As String = Nothing
        Dim Programstring As String = "Hello"
        Console.WriteLine("What is your Name")
        Userstring = Console.ReadLine()
        Console.WriteLine("Your name is : " & Userstring & ". Now i am adding some words to your name. So if you want to see those words with your name press Enter.")
        Console.ReadLine()
        Userstring = Programstring + " " + Userstring + " You are"
        Console.WriteLine(Userstring & " Rocking.")
        Console.ReadLine()
    End Sub

End Module
