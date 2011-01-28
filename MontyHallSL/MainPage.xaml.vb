Partial Public Class MainPage
    Inherits UserControl

    Public Sub New()
        InitializeComponent()
    End Sub

    Private mintStayWins As Integer = 0
    Private mintStayLoses As Integer = 0
    Private mintSwitchWins As Integer = 0
    Private mintSwitchLoses As Integer = 0
    Private mintLastThread As Integer = 0
    Private mintTotalTrials As Integer = 0
    Private mdatStart As Date
    Private mobjRandom As Random
    ''' <summary>
    ''' Starts the whole process
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        If Not Integer.TryParse(txtTrialNumber.Text, mintTotalTrials) Then
            MessageBox.Show("Please enter a numrice value for trials")
            Exit Sub
        End If
        ' reset variables
        ResetVariables()
        ' setup controls
        Button1.IsEnabled = False
        txtTrialNumber.IsEnabled = False
        lblComplete.Visibility = Windows.Visibility.Collapsed
        ProgressBar1.Maximum = mintTotalTrials
        ' startruns
        DoNextRun()
    End Sub
    Private Sub ResetVariables()
        mintStayWins = 0
        mintStayLoses = 0
        mintSwitchWins = 0
        mintSwitchLoses = 0
        mintLastThread = 0
        mdatStart = Now
        mintTotalTrials = CInt(txtTrialNumber.Text)
        mobjRandom = New Random
    End Sub
    Private Sub DoNextRun()
        mintLastThread += 1
        Dim objStay As New MontyHallCalculator(MontyHallCalculator.RunMode.Stay, mobjRandom, mintLastThread)
        AddHandler objStay.PlayerStatusDetermined, AddressOf PlayerStatusDetermined
        Dim objSwitch As New MontyHallCalculator(MontyHallCalculator.RunMode.Switch, mobjRandom, mintLastThread)
        AddHandler objSwitch.PlayerStatusDetermined, AddressOf PlayerStatusDetermined
        Dim tStayStart As New System.Threading.ThreadStart(AddressOf objStay.LetsMakeADeal)
        Dim tStay As New System.Threading.Thread(tStayStart)
        Dim tSwitchStart As New System.Threading.ThreadStart(AddressOf objSwitch.LetsMakeADeal)
        Dim tSwitch As New System.Threading.Thread(tSwitchStart)
        tStay.Start()
        tSwitch.Start()
    End Sub
    ''' <summary>
    ''' Event handler for MontyHall class
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="PlayerWon"></param>
    ''' <param name="CurrentRunMode"></param>
    ''' <param name="TrialNumber"></param>
    ''' <remarks></remarks>
    Private Sub PlayerStatusDetermined(ByVal sender As MontyHallCalculator, ByVal PlayerWon As Boolean, ByVal CurrentRunMode As MontyHallCalculator.RunMode, ByVal TrialNumber As Integer)
        If PlayerWon Then
            If CurrentRunMode = MontyHallCalculator.RunMode.Stay Then
                mintStayWins += 1
            Else
                mintSwitchWins += 1
            End If
        Else
            If CurrentRunMode = MontyHallCalculator.RunMode.Stay Then
                mintStayLoses += 1
            Else
                mintSwitchLoses += 1
            End If
        End If

        If TrialNumber < mintTotalTrials Then
            If TrialNumber Mod 5 = 0 Then
                ' update the ui
                Me.Dispatcher.BeginInvoke(New UpdateUIDel(AddressOf UpdateUI), TrialNumber)
            End If
            If CurrentRunMode = MontyHallCalculator.RunMode.Switch Then ' do next run
                DoNextRun()
            End If
        Else ' we are done...do final ui update
            ' update the ui
            Me.Dispatcher.BeginInvoke(New UpdateUIDel(AddressOf UpdateUI), TrialNumber)
        End If

        ' remove handler to clean memory
        RemoveHandler sender.PlayerStatusDetermined, AddressOf PlayerStatusDetermined

        sender = Nothing
    End Sub
    ''' <summary>
    ''' Need a delegate since thread callbacks are coming on a different thread.
    ''' </summary>
    ''' <param name="TrialNumber"></param>
    ''' <remarks></remarks>
    Private Delegate Sub UpdateUIDel(ByVal TrialNumber As Integer)
    Private Sub UpdateUI(ByVal TrialNumber As Integer)
        lblStuckCar.Content = mintStayWins
        lblStuckGoat.Content = mintStayLoses

        lblSwitchCar.Content = mintSwitchWins
        lblSwitchGoat.Content = mintSwitchLoses

        lblStuckPercentage.Content = Format(((mintStayWins / TrialNumber) * 100), "##.##") & "%"
        lblSwitchPercentage.Content = Format(((mintSwitchWins / TrialNumber) * 100), "##.##") & "%"

        If TrialNumber < mintTotalTrials Then
            ProgressBar1.Value = TrialNumber
        Else ' complete!
            ProgressBar1.Value = mintTotalTrials
            lblComplete.Content = "Complete!  Total Time to run was: " & Format(Now.Subtract(mdatStart).TotalSeconds, "####.##") & " seconds."
            lblComplete.Visibility = Visibility.Visible
            txtTrialNumber.IsEnabled = True
            Button1.IsEnabled = True
        End If
    End Sub
End Class
''' <summary>
''' This class has the main code with the Let's Make a Deal modeling
''' </summary>
''' <remarks></remarks>
Public Class MontyHallCalculator
    Public Enum RunMode
        Switch = 0
        Stay = 1
    End Enum
    Private mobjRandom As Random ' use the same random object througout to prevent dupe values
    Private menumRunMode As RunMode
    Private mintTrialNumber As Integer
    Public Event PlayerStatusDetermined(ByVal sender As MontyHallCalculator, ByVal PlayerWon As Boolean, ByVal CurrentRunMode As RunMode, ByVal TrialNumber As Integer)
    Public Sub New(ByVal Mode As RunMode, ByRef objRandom As Random, ByVal TrialNumber As Integer)
        menumRunMode = Mode
        mobjRandom = objRandom
        mintTrialNumber = TrialNumber
    End Sub
    ''' <summary>
    ''' This is the routine with the Lets Make a Deal modelling.  It is purposely verbose and sequential to see the logic
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LetsMakeADeal()
        ' randomly choose the users door
        Dim intChoice = mobjRandom.Next
        If intChoice > Integer.MaxValue * 0.66666666666 Then
            intChoice = 3
        ElseIf intChoice > Integer.MaxValue * 0.333333333333 Then
            intChoice = 2
        Else
            intChoice = 1
        End If

        ' randomly choose the car door
        Dim intCar = mobjRandom.Next
        If intCar > Integer.MaxValue * 0.66666666666 Then
            intCar = 3
        ElseIf intCar > Integer.MaxValue * 0.333333333333 Then
            intCar = 2
        Else
            intCar = 1
        End If

        ' Deteremine which door we want to open first based on rules from show.
        ' This is where you see why Monty Hall theory is true...Monty Hall's choice of
        ' what door to open first is based on logic and isn't at random.  Namely it must be
        ' a goat AND also cannot be the players pick.  This means that it effectively eliminates 2 options
        ' for Monty if the player has a goat but only one option if the player has the car.  Since the player
        ' has a goat 66% of the time on initial pick (2/3 are goats) switching turns the odds around to 66% in his favor
        ' since he will win 100% of the time when he initially picked the goat and he switches as Monty has eliminated the other goat!
        Dim intFirstDoorToOpen As Integer
        If intChoice = intCar Then ' customer chose car so we are going to one of the 2 other doors first
            ' determine which of the 2 doors that aren't the car
            Dim bOpenLowestAvailableDoor As Boolean
            bOpenLowestAvailableDoor = mobjRandom.Next > Integer.MaxValue / 2
            If intChoice = 1 Then
                If bOpenLowestAvailableDoor Then
                    intFirstDoorToOpen = 2
                Else
                    intFirstDoorToOpen = 3
                End If
            ElseIf intChoice = 2 Then
                If bOpenLowestAvailableDoor Then
                    intFirstDoorToOpen = 1
                Else
                    intFirstDoorToOpen = 3
                End If
            ElseIf intChoice = 3 Then
                If bOpenLowestAvailableDoor Then
                    intFirstDoorToOpen = 1
                Else
                    intFirstDoorToOpen = 2
                End If
            Else
                'error
            End If
        Else ' goat was chosen so we can only open the door that is NOT the players choice OR the car
            If intChoice = 1 Then
                If intCar = 2 Then
                    intFirstDoorToOpen = 3
                Else
                    intFirstDoorToOpen = 2
                End If
            ElseIf intChoice = 2 Then
                If intCar = 1 Then
                    intFirstDoorToOpen = 3
                Else
                    intFirstDoorToOpen = 1
                End If
            ElseIf intChoice = 3 Then
                If intCar = 1 Then
                    intFirstDoorToOpen = 2
                Else
                    intFirstDoorToOpen = 1
                End If
            Else
                'error
            End If
        End If

        ' if we are in switch mode figure out what to switch to
        If menumRunMode = RunMode.Switch Then
            If intChoice = 1 Then
                If intFirstDoorToOpen = 2 Then
                    intChoice = 3
                Else
                    intChoice = 2
                End If
            ElseIf intChoice = 2 Then
                If intFirstDoorToOpen = 1 Then
                    intChoice = 3
                Else
                    intChoice = 1
                End If
            ElseIf intChoice = 3 Then
                If intFirstDoorToOpen = 1 Then
                    intChoice = 2
                Else
                    intChoice = 1
                End If
            End If
        End If

        If intCar = intChoice Then
            RaiseEvent PlayerStatusDetermined(Me, True, menumRunMode, mintTrialNumber)
        Else
            RaiseEvent PlayerStatusDetermined(Me, False, menumRunMode, mintTrialNumber)
        End If

    End Sub
End Class