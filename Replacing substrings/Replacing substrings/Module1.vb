Module Module1

    Sub Main()
        Dim Mystring As String = Nothing
        Dim Finalstring As String = Nothing
        Console.WriteLine("Please enter a word or a sentence")
        Mystring = Console.ReadLine()
        Console.WriteLine("The word or sentence you typed is : " & Mystring)
        Finalstring = Mystring.Replace("a", "Y")
        Console.WriteLine("After replacing all the letter 'a' to 'y' the word or sentence you typed look like this " & Finalstring)
        Console.ReadLine()
    End Sub

End Module
