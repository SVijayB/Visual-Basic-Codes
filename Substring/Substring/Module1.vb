Module Module1

    Sub Main()
        Dim Mystring As String = Nothing
        Console.WriteLine("Please type a word more then 10 letters without using space.")
        Mystring = Console.ReadLine()
        Console.WriteLine()
        Console.WriteLine("The word you typed is: " & Mystring)
        Console.WriteLine("The number of letters in the word you typed are: " & Mystring.Length.ToString())
        Console.WriteLine("The first three letters of the word you typed are: " & Mystring.Substring(0, 3))
        Console.WriteLine("The second three letters of the word you typed are: " & Mystring.Substring(3, 3))
        Console.WriteLine("The other letters left are: " & Mystring.Substring(6))
        Console.ReadLine()
    End Sub

End Module
