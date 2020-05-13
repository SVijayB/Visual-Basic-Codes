Module Module1

    Sub Main()
        Dim String1 As String = Nothing
        Dim String2 As String = Nothing
        Console.WriteLine("Please enter a word to compare with other word which you will type next")
        String1 = Console.ReadLine()

        Console.WriteLine("The first word you typed to compare is : " & String1)
        Console.WriteLine("Please enter another word to compare with the first word you typed")
        String2 = Console.ReadLine()
        Console.WriteLine("The second word you typed to compare with the first one is : " & String2)

        Console.WriteLine("The number of Differences in the letters of the first one to the second one is : " & String.Compare(String1, String2, True))
        Console.ReadLine()
    End Sub

End Module
