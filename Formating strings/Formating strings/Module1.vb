Module Module1

    Sub Main()
        Dim Mystring As String = Nothing
        Console.WriteLine("Please enter a word or a sentence to format")
        Mystring = Console.ReadLine()
        Console.WriteLine("The word or sentence you entered for formatting is: " & Mystring)
        Console.WriteLine("Please enter a decimal value to format")
        Dim Mydouble As Double = Console.ReadLine()
        Console.WriteLine("The decimal value you entered for formatting is: " & Mydouble)
        Console.WriteLine()
        Console.WriteLine("The formated form of the decimal value you gave is: " & String.Format("{0:n2}", Mydouble))
        Console.WriteLine("All the letters you typed in the word or sentence is now in Capital . Look at it: " & Mystring.ToUpper())
        Console.WriteLine("Each and every letters you typed in the word or sentence is now shown without a capital letter . Look at it: " & Mystring.ToLower())
        Console.ReadLine()
    End Sub

End Module
