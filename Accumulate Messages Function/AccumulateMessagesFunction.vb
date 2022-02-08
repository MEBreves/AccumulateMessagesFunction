'Miranda Breves
'RCET0265
'Spring 2022
'Accumulate Messages Function
'https://github.com/MEBreves/AccumulateMessagesFunction

Option Strict On
Option Explicit On

Module AccumulateMessagesFunction

    'A function that retains and concatenates messages for the user. The messages can be cleared.
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

    'A function that checks if the user input is a "y" or "n" for yes or no, then returns a true or false value.
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

                'Getting a message from the user and storing it in the message list
                Console.WriteLine("-----------------------------------------------------------------------")
                Console.WriteLine("Please enter a message.")
                Console.Write("> ")
                message = Console.ReadLine()
                messageList = UserMessages(message, emptyMessages)
                Console.WriteLine()

                'Asking the user if they want to add another message and re-running the loop if so
                Console.WriteLine("Would you like to add another message? Y / N")
                Console.Write("> ")
                userInput = Console.ReadLine()
                stillCreatingMessages = CheckInputsForBoolean(userInput)
                Console.WriteLine()

            Loop

            stillCreatingMessages = True    'The variable is reset to true so that if the messages program loops this will run again

            'Displaying the accumulated messages to the user
            Console.WriteLine("Here are your messages:")
            Console.WriteLine(messageList)

            'Asking the user if they want their messages cleared
            Console.WriteLine("Would you like to clear your messages? Y / N")
            Console.Write("> ")
            userInput = Console.ReadLine()
            emptyMessages = CheckInputsForBoolean(userInput)

            'Emptying the user messsages if true; the user will be notified if it happens
            If emptyMessages Then

                UserMessages("", emptyMessages) 'Clearing the messages if the user said yes
                Console.WriteLine("Your messages have been cleared.")
                emptyMessages = False           'Setting this variable back to a default false for use in another loop

            End If

            'Asking if the user wants to re-run the program and then either re-running or ending the loop
            Console.WriteLine(vbNewLine & "Would you like to keep using the message program? Y / N")
            Console.Write("> ")
            userInput = Console.ReadLine()
            usingMessagesProgram = CheckInputsForBoolean(userInput)

            Console.WriteLine()

        Loop

        'Allowing the user to see the end of the program instead of just turning off.
        Console.WriteLine("-----------------------------------------------------------------------")
        Console.WriteLine("Please press Enter to exit the program.")
        Console.ReadLine()


    End Sub

End Module
