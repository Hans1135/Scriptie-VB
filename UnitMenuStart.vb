Imports System.Windows.Controls
Imports System.Windows.Input

Module UnitMenuStart


    Public MenuStart As New ClassMenuStart()

    Class ClassMenuStart
        Inherits MenuItem

        Function Init()

            ' start menu start

            Header = "_Start"

            Items.Add(StartPars.Init)
            Items.Add(StartTo.Init)
            Items.Add(StartStep.Init)
            Items.Add(StartPlay.Init)
            Items.Add(StartStop.Init)

            Return Me
        End Function
    End Class

    Public WithEvents StartPars As New ClassStartPars()

    Class ClassStartPars
        Inherits MenuItem

        Function Init()

            Header = "_Pars"
            InputGestureText = "F5"
            GridMenu.Shortcut(Key.F5, vbNull, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' start parsen

            ParsTo = False
            TextParsInit()
        End Sub
    End Class

    Public WithEvents StartTo As New ClassStartTo()

    Class ClassStartTo
        Inherits MenuItem

        Function Init() As MenuItem


            Header = "_To"
            InputGestureText = "F6"
            GridMenu.Shortcut(Key.F6, vbNull, AddressOf Me_Click)
            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' start parsstap

            ParsTo = True
            TextParsInit()
        End Sub
    End Class

    Public WithEvents StartStep As New ClassStartStep()

    Class ClassStartStep
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Step"
            InputGestureText = "F11"
            GridMenu.Shortcut(Key.F11, vbNull, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' start parsstap

            TextParsStep()
        End Sub
    End Class

    Public WithEvents StartPlay As New ClassStartPlay()

    Class ClassStartPlay
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Play"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' start afspelen

            MediaPlay.Start()
        End Sub
    End Class

    Public WithEvents StartStop As New ClassStartStop()

    Class ClassStartStop
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Stop"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' stopt afspelen

            MediaPlay.Stop()
            TabsCtrl.Done()
        End Sub
    End Class
End Module