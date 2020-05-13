Module Module1

    Sub Main()
        Dim Kindergarden As Boolean = Nothing
        Dim Drive As Boolean = Nothing
        Dim Park As Boolean = Nothing
        Console.WriteLine("What is your name ? ")
        Dim Name As String = Console.ReadLine()
        Console.WriteLine("What is your age ? " & Name & ".")
        Dim age As Integer = Console.ReadLine()

        Dim OutputKindergarden As String = Nothing
        Dim Outputpark As String = Nothing
        Dim Outputdrive As String = Nothing

        If age <> 5 Then
            Kindergarden = False
            OutputKindergarden = Name & " You are not in Kindergarden."
        Else
            Kindergarden = True
            OutputKindergarden = Name & " You are in Kindergarden!"
        End If

        If age < 18 Then
            Park = True
            Outputpark = "You can play in the park!"
        Else
            Park = False
            Outputpark = "Sorry! You can't play in the park."
        End If

        If age > 18 Then
            Drive = True           
                Outputdrive = "You can drive legally!"
            Else
            Drive = False
                Outputdrive = "You can drive but don't get Screwed."
            End If
        Console.WriteLine(OutputKindergarden & " , " & Outputpark & " and " & Outputdrive)
        Console.ReadLine()
    End Sub

End Module
