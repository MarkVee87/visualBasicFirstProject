'Allows the usage of a library of mathematic tools such as "sqrt" - square root
Imports System.Math
Module Module1034528

    'Global variables for use with the Accuracy Option
    Dim DP As Integer, answer As Double, answerresult As Double


    Sub Main()

        'This title will be shown in the top of the console window
        Console.Title = "Mark's Visual Basic Project"

        'Title and intro for main menu that display in console
        Console.WriteLine("* * * * * * * * MAIN MENU * * * * * * * *")
        Console.WriteLine("")
        Console.WriteLine("")
        Console.WriteLine("Please select an option from the list below:")
        Console.WriteLine("")
        Console.WriteLine("")

        'Main menu options that display in console
        Console.WriteLine("1: Set accuracy (number of decimal points)")
        Console.WriteLine("2: Quadratic equation calculator")
        Console.WriteLine("3: Determine if a number is prime")
        Console.WriteLine("4: Monte-Carlo Integration")
        Console.WriteLine("5: Exit")
        Console.WriteLine("")
        Console.WriteLine("")

        'Main menu functionality from user entry
        Console.Write("Entry: ")
        Dim menuentry As String
        menuentry = Console.ReadLine() 'Whatever is entered is related to in the following "select case"

        'Resulting options to relevant user entry
        Select Case menuentry

            Case 1 '"Decimal points required"/"Accuracy Option"
                Console.Clear()
                Console.WriteLine("* * * * * * * SET DECIMAL PLACES OF ANSWERS * * * * * * *")
                Console.WriteLine("")
                Console.WriteLine("")

                Console.WriteLine("Please enter the decimal value required for all answers")
                Console.WriteLine("")
                Console.WriteLine("(Decimal places can be set to a value between 1 and 5)")
                Console.WriteLine("")
                Console.Write("Entry: ")


                Try 'this is the start of a test of valid user inputs

                    DP = Console.ReadLine() 'User enters a value between 1 and 5, resulting actions are seen below and in the function "FormatAnswers"
                    Console.WriteLine("")
                    Console.WriteLine("")
                    Console.WriteLine("")

                Catch 'if the user enters and invalid input, the following will be ran rather than continuing

                    Console.WriteLine("")
                    Console.WriteLine("")
                    Console.WriteLine("Entry must be an integer value between 1 and 5 (inclusive)")
                    Console.WriteLine("")
                    Console.WriteLine("")
                    Call returntomain()

                End Try 'ends the testing sequence as all entries are deemed as valid


                Select Case DP 'Takes users entry from above and applies to the relevant option in here and in the function "FormatAnswers" as DP is a global variable (see top of page)
                    Case 1
                        Console.WriteLine("Answers will be set to 1 decimal place")
                    Case 2
                        Console.WriteLine("Answers will be set to 2 decimal places")
                    Case 3
                        Console.WriteLine("Answers will be set to 3 decimal places")
                    Case 4
                        Console.WriteLine("Answers will be set to 4 decimal places")
                    Case 5
                        Console.WriteLine("Answers will be set to 5 decimal places")
                    Case Else 'If anything else is entered the user is alerted of it being an invalid entry
                        Console.WriteLine("This program only supports answers up to 5 decimal places")
                        Call returntomain()
                End Select



                'Function called to format the answers in the particular form requested
                Call FormatAnswers(DP)

                'Subroutine designed to display a standard "return to menu" screen when the process is completed
                Call returntomain()


            Case 2 '"Quadratic equation calculator"
                Console.Clear()
                Console.WriteLine("* * * * * * * QUADRATIC EQUATION CALCULATOR * * * * * * *")
                Console.WriteLine("")
                Console.WriteLine("Please enter a quadratic equation of the form: Ax^2 (+/-) Bx (+/-) C")
                Console.WriteLine("")
                Console.WriteLine("")

                'Declares the variables for usage in the Quadratic formula
                Dim A As Double, B As Double, C As Double

                Try 'this is the start of a test of valid user inputs

                    'Takes a value for A to be applied to the formula
                    Console.WriteLine("Please enter a value for A: ")
                    A = Console.ReadLine() 'A = user input
                    Console.WriteLine("")

                    'Takes a value for B to be applied to the formula
                    Console.WriteLine("Please enter a value for B: ")
                    B = Console.ReadLine() 'B = user input
                    Console.WriteLine("")

                    'Takes a value for C to be applied to the formula
                    Console.WriteLine("Please enter a value for C: ")
                    C = Console.ReadLine() 'C = user input
                    Console.WriteLine("")

                Catch 'if the user enters and invalid input, the following will be ran rather than continuing

                    Console.WriteLine("")
                    Console.WriteLine("")
                    Console.WriteLine("Entries must be numerical values only")
                    Console.WriteLine("")
                    Console.WriteLine("")
                    Call returntomain()
                End Try 'ends the testing sequence as all entries are deemed as valid


                'Takes the values of A, B and C (above) to be applied to the subroutine "QuadraticEquationCalculator"
                Call QuadraticEquationCalculator(A, B, C)
                'Calls the main menu when function has completed
                Call Main()


            Case 3 '"Prime number calculator"
                Console.Clear()
                Console.WriteLine("* * * * * * * PRIME NUMBER CALCULATOR * * * * * * *")
                Console.WriteLine("")
                Console.WriteLine("")

                Dim i As Long 'i will be represent the number to be evaluated

                Console.WriteLine("Please enter a number to be evaluated: ")
                i = Console.ReadLine() 'declares the value of i as the users' input

                Console.WriteLine("")
                Console.WriteLine("")

                'takes the i value and applies it to the subroutine "PrimeNumberCalculator"
                Call PrimeNumberCalculator(i)


            Case 4 '"Monte-carlo integration" (not being implemented)
                Console.WriteLine("Monte-Carlo Integration")
                Console.Clear()
                Console.WriteLine("* * * * * * * MONTE-CARLO INTEGRATION * * * * * * *")
                Console.WriteLine("")
                Console.WriteLine("")
                Console.WriteLine("Feature not implemented") 'Displays a simple message saying that this function is not being implemented.
                Call returntomain()


            Case 5 '"Exit"
                Console.WriteLine("")
                'Short message displayed before the console window closes
                Console.WriteLine("Thank you for using Mark's program, press any key to exit")
                Console.WriteLine("")
                Environment.Exit(0) 'closes the console


            Case Else 'If anything other than the applicable options is entered, the following message will be displayed:
                Console.WriteLine("")
                Console.WriteLine("Invalid entry, please enter a number between 1 and 5")
                Console.WriteLine("Press enter to continue...")
                Console.ReadLine()

                'Then the screen will be reset back to the main menu for the user to enter another value, see below
                Console.Clear()
                Call Main()

        End Select

    End Sub

    'This sub takes the values A, B and C from the users' input as shown in the "Sub Main" section
    Sub QuadraticEquationCalculator(ByVal A As Double, ByVal B As Double, ByVal C As Double)

        'These variables define the possible outputs to calculate and display the answers with
        Dim x1answer As Double, x2answer As Double, discriminant As Double

        'working out the discriminant first keeps the coding neat and allows for easy identification of possible bugs
        discriminant = B ^ 2 - 4 * A * C

        'this section immediately defines whether or not it is worthwhile to continue the calculation, if the discriminant is less than 0, there will not be any point in continuing to find the roots.
        If discriminant < 0 Then
            Console.WriteLine("")
            Console.WriteLine("This equation has no real roots") 'if this is the case, a short message is displayed and then the "returntomain" sub is applied

            'if the discriminant is 0 or greater, the quadratic formula is applied for both positive and negative outputs
        Else
            Console.WriteLine("")

            x1answer = (-B + Sqrt(discriminant)) / (2 * A) 'positive is determined
            Call FormatAnswers(x1answer) 'relates to the accuracy option function according to the users' needs
            Console.WriteLine("x intersect 1 = " & answerresult) 'answer is outputted (with set number of dec places if requested)

            x2answer = (-B - Sqrt(discriminant)) / (2 * A) 'negative is determined
            Call FormatAnswers(x2answer) 'relates to the accuracy option function according to the users' needs
            Console.WriteLine("x intesect 2 = " & answerresult) 'answer is outputted (with set number of dec places if requested)
        End If

        'calculations and output are completed so the returntomain subroutine is called again to take the user back to the main menu
        Call returntomain()

    End Sub

    'This sub takes the value "i" from the DP entry in the main sub and applies it to the appropriate format required
    Sub PrimeNumberCalculator(ByVal num As Long) 'num represents the original "i" value

        Dim i As Long, n As Long, totaltimetaken As Double 'defined variables

        'A timer is set to see how long it takes to complete the process
        Dim StartTimer As Double, StopTimer As Double 'timer variables set to double to show an exact length of time
        StartTimer = Timer 'the timer is started


        For i = 2 To num - 1 'i is set to 2 as primes must be 2 or greater
            n = num Mod i
            If n = 0 Then 'rules out the option of trying to process 0 as a prime
                Console.WriteLine("")
                Console.WriteLine(num & " is not a prime number")
                Exit For 'exits as soon as a value other than 1 or itself can be divided into num as is a prime
            End If
        Next i 'loops around to determine each eventuallity of division between 2 and num - 1

        If i = num Then 'if the above for loop cannot create an output, the value is passed to this section to determine it as a prime
            Console.WriteLine("")
            Console.WriteLine(i & " is a prime number")
        End If


        StopTimer = Timer 'timer is stopped

        Console.WriteLine("")
        Console.WriteLine("")
        Console.Write("Time taken to process request = ")
        totaltimetaken = (StopTimer - StartTimer) 'total process time is calculated and displayed
        Console.WriteLine(Format(totaltimetaken, "0.###") & " seconds (3.d.p)")
        Console.WriteLine("")

        Call returntomain()

    End Sub

    'Uses global variables and user input to determine the accuracy of answers in the quadratic calculator
    Function FormatAnswers(ByVal answer As Double) As Double

        'Takes the users' original input from the main menu to determine the required style of output for answers
        Select Case DP

            Case 1 '1 d.p.
                answerresult = Format(answer, "#.0")
            Case 2 '2 d.p.
                answerresult = Format(answer, "#.00")
            Case 3 '3 d.p.
                answerresult = Format(answer, "#.000")
            Case 4 '4 d.p.
                answerresult = Format(answer, "#.0000")
            Case 5 '5 d.p.
                answerresult = Format(answer, "#.00000")
            Case Else
                answerresult = answer

        End Select

    End Function

    'Simple sub to remove the need to repeat coding on each other sub/function when returing to the menu, also maintains consistency in the output
    Sub returntomain()


        'Basic text output to alert the user to press enter to return to the main menu
        Console.WriteLine("")
        Console.WriteLine("*****************************************")
        Console.WriteLine("")
        Console.WriteLine("Press enter to return to the main menu...")

        'Console.readline holds the page and waits for the user to press a key to continue, rather than automatically going to the main menu without time t show this above message
        Console.ReadLine()

        'Clears the screen of all output and recalls the main menu
        Console.Clear()
        Call Main()

    End Sub

End Module

