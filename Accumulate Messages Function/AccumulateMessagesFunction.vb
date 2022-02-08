'Miranda Breves
'RCET0265
'Spring 2022
'Accumulate Messages Function
'url

Option Strict On
Option Explicit On

Module AccumulateMessagesFunction

    Function UserMessages(ByVal newMessage As String, ByVal clear As Boolean) As String

        'Declaring function variable
        Static allMessages As String

        If clear Then
            allMessages = ""
        Else
            allMessages += newMessage & vbNewLine
        End If

        Return allMessages

    End Function

    Function CheckInputsForBoolean(ByVal input As String) As Boolean

        Dim endLoop As Boolean = False
        Dim returnValue As Boolean

        Do Until endLoop
            If input.ToLower() = "y" Then
                returnValue = True
                endLoop = True
            ElseIf input.ToLower = "n" Then
                returnValue = False
                endLoop = True
            Else
                Console.WriteLine("Your input wasn't valid.")
                Console.WriteLine("Please enter a Y for yes or N for No.")
                Console.Write("> ")
                input = Console.ReadLine()
                endLoop = False
            End If
        Loop

        Return returnValue

    End Function

    Sub Main()

        'Declaring variables
        Dim message As String
        Dim messageList As String
        Dim userInput As String
        Dim emptyMessages As Boolean = False
        Dim stillCreatingMessages As Boolean = True
        Dim usingMessagesProgram As Boolean = True

        Do While usingMessagesProgram

            Do While stillCreatingMessages

                Console.WriteLine("-----------------------------------------------------------------------")
                Console.WriteLine("Please enter a message.")
                Console.Write("> ")
                message = Console.ReadLine()
                messageList = UserMessages(message, emptyMessages)

                Console.WriteLine()
                Console.WriteLine("Would you like to add another message? Y / N")
                Console.Write("> ")
                userInput = Console.ReadLine()

                stillCreatingMessages = CheckInputsForBoolean(userInput)
                Console.WriteLine()

            Loop

            stillCreatingMessages = True    'The variable is reset to true so that another iteration of the loop can happen later

            Console.WriteLine("Here are your messages:")
            Console.WriteLine(messageList)

            Console.WriteLine("Would you like to clear your messages? Y / N")
            Console.Write("> ")
            userInput = Console.ReadLine()
            emptyMessages = CheckInputsForBoolean(userInput)

            If emptyMessages Then

                UserMessages("", emptyMessages) 'Clearing the messages if the user said yes
                Console.WriteLine("Your messages have been cleared.")
                emptyMessages = False           'Setting this variable back to a default false for use in another loop

            End If

            Console.WriteLine(vbNewLine & "Would you like to keep using the message program? Y / N")
            Console.Write("> ")
            userInput = Console.ReadLine()
            usingMessagesProgram = CheckInputsForBoolean(userInput)

            Console.WriteLine()

        Loop

        Console.WriteLine("-----------------------------------------------------------------------")
        Console.WriteLine("Please press Enter to exit the program.")
        Console.ReadLine()


    End Sub

End Module
