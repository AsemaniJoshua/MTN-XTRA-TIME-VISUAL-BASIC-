Public Class Form1
    Private Sub Start_Click(sender As Object, e As EventArgs) Handles Start.Click

        Dim UserInput As Integer
        Dim Code As String
        Dim count As Integer

        count = 1


        MsgBox(" Welcome to MTN Xtra Time." & vbNewLine & vbNewLine _
            & vbNewLine & "Click on ""ok"" or Press ""Enter"" on the keyborad to Start")

        Code = InputBox(" Enter the MTN Xtra Time Code " & vbNewLine & vbNewLine & "** Add the * and # **")

        Do While Code <> "*506#"

            count += 1

            If count = 3 Then
                MsgBox("Your Chance to enter the Xtra Time code as reached it's limit." & vbNewLine & vbNewLine &
                       "Please Start the Process Again 🤖")
                Exit Sub
            End If

            MsgBox("🤖 You entered the wrong Code " & vbNewLine & vbNewLine & "Please try again 😊")
            Code = InputBox(" Enter the MTN Xtra Time Code " & vbNewLine & vbNewLine & "** Add the * and # **")

        Loop

        '......................................If the Xtra Time Code Is Correct...................................

        If Code = "*506#" Then

            Try

                UserInput = InputBox("MTN Xtra Time." & vbNewLine &
                                    "Please select an option:" & vbNewLine &
                                    "1. GHS 5.00" & vbNewLine &
                                    "2. GHS 4.00" & vbNewLine &
                                    "3. More Advance Options" & vbNewLine &
                                    "4. Request a Data Advance" & vbNewLine &
                                    "5. Menu" & vbNewLine)

            Catch

                MsgBox("Invalid Input")

                Exit Sub

          
            End Try


        End If

        Try
            '......................................If the User Chose Main Option 1...................................

            If UserInput = 1 Then
                UserInput = InputBox("You have selected GHS 5." & vbNewLine &
                                    "GHS 5.50 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                    "1. Confirm" & vbNewLine &
                                    "2. Cancel" & vbNewLine
                                    )
                '**********************************If the User will confirm or not************************************
                If UserInput = 1 Then
                    MsgBox("The amount of 5.00 GHC has been added to your account." & vbNewLine & vbNewLine &
                       "The amount of 5.50 GHC will be deducted from your next reload.")
                    Exit Sub
                Else
                    MsgBox("Cancelled.")
                    Exit Sub
                End If

                '......................................If the User Chose Main Option 2...................................

            ElseIf UserInput = 2 Then

                UserInput = InputBox("You have selected GHS 4." & vbNewLine &
                                   "GHS 4.40 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                   "1. Confirm" & vbNewLine &
                                   "2. Cancel" & vbNewLine
                                   )

                '**********************************If the User will confirm or not************************************
                If UserInput = 1 Then
                    MsgBox("The amount of 4.00 GHC has been added to your account." & vbNewLine & vbNewLine &
                       "The amount of 4.40 GHC will be deducted from your next reload.")
                    Exit Sub
                Else
                    MsgBox("Cancelled.")
                    Exit Sub
                End If

                '......................................If the User Chose Main Option 3...................................

            ElseIf UserInput = 3 Then

                UserInput = InputBox("Welcome to MTN Xtra Time." & vbNewLine &
                                    "Please select an option:" & vbNewLine &
                                    "1. GHS 5.00" & vbNewLine &
                                    "2. GHS 4.00" & vbNewLine &
                                    "3. GHS 3.00" & vbNewLine &
                                    "4. GHS 2.00" & vbNewLine
                                    )




                '......................................If the User Chose Sub Option 1...................................

                If UserInput = 1 Then
                    UserInput = InputBox("You have selected GHS 5." & vbNewLine &
                                            "GHS 5.50 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                            "1. Confirm" & vbNewLine &
                                            "2. Cancel" & vbNewLine
                                            )
                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("The amount of 5.00 GHC has been added to your account." & vbNewLine & vbNewLine &
                               "The amount of 5.50 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If

                    '......................................If the User Chose Sub Option 2...................................

                ElseIf UserInput = 2 Then

                    UserInput = InputBox("You have selected GHS 4." & vbNewLine &
                                           "GHS 4.40 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                           "1. Confirm" & vbNewLine &
                                           "2. Cancel" & vbNewLine
                                           )

                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("The amount of 4.00 GHC has been added to your account." & vbNewLine & vbNewLine &
                               "The amount of 4.40 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If


                    '......................................If the User Chose Sub Option 3...................................

                ElseIf UserInput = 3 Then

                    UserInput = InputBox("You have selected GHS 3." & vbNewLine &
                                           "GHS 3.30 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                           "1. Confirm" & vbNewLine &
                                           "2. Cancel" & vbNewLine
                                           )

                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("The amount of 3.00 GHC has been added to your account." & vbNewLine & vbNewLine &
                               "The amount of 3.30 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If

                    '......................................If the User Chose Sub Option 4...................................

                ElseIf UserInput = 4 Then

                    UserInput = InputBox("You have selected GHS 2." & vbNewLine &
                                           "GHS 2.20 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                           "1. Confirm" & vbNewLine &
                                           "2. Cancel" & vbNewLine
                                           )

                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("The amount of 2.00 GHC has been added to your account." & vbNewLine & vbNewLine &
                               "The amount of 2.20 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If


                Else

                    MsgBox("Invalid Input")

                    Exit Sub

                End If

                '......................................If the User Chose Main Option 4...................................

            ElseIf UserInput = 4 Then

                UserInput = InputBox("Welcome to MTN Xtra Time." & vbNewLine &
                                   "Please select an option:" & vbNewLine &
                                   "1. 514 MB" & vbNewLine &
                                   "2. 402 MB" & vbNewLine &
                                   "3. 98 MB" & vbNewLine &
                                   "4. 41 MB" & vbNewLine
                                   )

                '......................................If the User Chose Sub Option 1...................................

                If UserInput = 1 Then
                    UserInput = InputBox("You have selected 514 MB." & vbNewLine &
                                            "GHS 5.50 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                            "1. Confirm" & vbNewLine &
                                            "2. Cancel" & vbNewLine
                                            )
                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("514 MB has been added to your data bundle." & vbNewLine & vbNewLine &
                               "The amount of 5.50 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If

                    '......................................If the User Chose Sub Option 2...................................

                ElseIf UserInput = 2 Then

                    UserInput = InputBox("You have selected 402 MB." & vbNewLine &
                                           "GHS 4.40 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                           "1. Confirm" & vbNewLine &
                                           "2. Cancel" & vbNewLine
                                           )

                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("402 MB has been added to your data bundle." & vbNewLine & vbNewLine &
                               "The amount of 4.40 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If


                    '......................................If the User Chose Sub Option 3...................................

                ElseIf UserInput = 3 Then

                    UserInput = InputBox("You have selected 98 MB." & vbNewLine &
                                           "GHS 3.30 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                           "1. Confirm" & vbNewLine &
                                           "2. Cancel" & vbNewLine
                                           )

                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("98 MB has been added to your data bundle." & vbNewLine & vbNewLine &
                               "The amount of 3.30 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If

                    '......................................If the User Chose Sub Option 4...................................

                ElseIf UserInput = 4 Then

                    UserInput = InputBox("You have selected 41 MB." & vbNewLine &
                                           "GHS 2.20 will be deducted from your next airtime reload or MoMo bundle purchase" & vbNewLine &
                                           "1. Confirm" & vbNewLine &
                                           "2. Cancel" & vbNewLine
                                           )

                    '**********************************If the User will confirm or not************************************
                    If UserInput = 1 Then
                        MsgBox("41 MB has been added to your data bundle." & vbNewLine & vbNewLine &
                               "The amount of 2.20 GHC will be deducted from your next reload.")
                        Exit Sub
                    Else
                        MsgBox("Cancelled.")
                        Exit Sub
                    End If


                Else

                    MsgBox("Invalid Input")

                    Exit Sub

                End If


            ElseIf UserInput = 5 Then

                UserInput = InputBox("MTN Xtra Time." & vbNewLine &
                                    "Please select an option:" & vbNewLine &
                                    "1. Status" & vbNewLine &
                                    "2. Info" & vbNewLine &
                                    "3. Outstanding Credit" & vbNewLine &
                                    "4. History" & vbNewLine & vbNewLine
                                    )

                '......................................If the User Chose Sub Option 1...................................

                If UserInput = 1 Then

                    MsgBox("Y'ello. You are eligible for Xtra Time service." & vbNewLine & vbNewLine &
                    "You can dial *506# and request an airtime or data advance anytime.")

                    Exit Sub

                    '......................................If the User Chose Sub Option 2...................................

                ElseIf UserInput = 2 Then

                    MsgBox("Y'ello." & vbNewLine &
                    "If you are eligible for Xtra Time service, you can request an airtime or data advance when your balance runs low." & vbNewLine &
                    "Just dial *506#.")

                    Exit Sub

                    '......................................If the User Chose Sub Option 3...................................

                ElseIf UserInput = 3 Then

                    MsgBox("Y'ello." & vbNewLine &
                    "Your current pending credit for MTN Xtra Time is zero." & vbNewLine &
                    "Dial *124# to know your balance." & vbNewLine &
                    "You can dial *506# and request Xtra Time anytime.")

                    Exit Sub

                    '......................................If the User Chose Sub Option 4...................................

                ElseIf UserInput = 4 Then

                    MsgBox("Y'ello." & vbNewLine &
                    "Your request is being processed." & vbNewLine &
                    "You will receive an SMS notification soon." & vbNewLine
                    )

                    Exit Sub


                Else

                    MsgBox("Invalid Input")

                    Exit Sub


                End If



            Else

                MsgBox("Invalid Input")

                Exit Sub

            End If

        Catch ex As Exception

            MsgBox("Invalid Input")

            Exit Sub

        Finally

            MsgBox("Thank you for using MTN")

        End Try


    End Sub


End Class
