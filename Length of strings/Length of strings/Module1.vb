Module Module1

    Sub Main()
        Dim Username As String = Nothing
        Console.WriteLine("Please Create your own username by typing it in the length of 10 or less.")
        Username = Console.ReadLine()
        If (Username = "Vijay" Or Username = "vijay" Or Username = "Balaji" Or Username = "balaji" Or Username = "Aishvarya" Or Username = "aishvarya" Or Username = "Anbusudha" Or Username = "anbusudha" Or Username = "Sankar" Or Username = "sankar") Then
            Console.WriteLine("Sorry , the username you wanted is already used . don't worry try it again with a new username.")
        Else
            If Username.Length.Equals(10) Or Username.Length <= 10 Then
                Console.WriteLine("You have writen the password at the correct length! So you are granded access to make your very own website! But remember don't forget your username!")
            Else
                Console.WriteLine("You haven't writen the password in the correct length. So please try again . ")
            End If
        End If
        Console.ReadLine()
    End Sub

End Module
