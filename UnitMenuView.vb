Imports System.Windows.Controls
Imports System.Windows.Input

Module UnitMenuView

    Public MenuView As New ClassMenuView()

    Class ClassMenuView
        Inherits MenuItem


        Function Init()

            ' start view menu

            Header = "_View"

            Items.Add(ViewMode.Init)
            Items.Add(ViewPrev.Init)
            Items.Add(ViewTools.Init)
            Items.Add(ViewInfo.Init)
            Items.Add(ViewGoogle.Init)

            Return Me
        End Function
    End Class

    Public WithEvents ViewMode As New ClassViewMode()

    Class ClassViewMode
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Mode"
            InputGestureText = "Ctrl+M"
            GridMenu.Shortcut(Key.M, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' verandert grid indeling

            FormGrid.Change()
        End Sub
    End Class

    Public WithEvents ViewPrev As New ClassViewPrev()

    Class ClassViewPrev
        Inherits MenuItem

        Function Init()

            Header = "_Previous"
            InputGestureText = "Ctrl+P"
            GridMenu.Shortcut(Key.P, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' toont vorige bestand

            Dim I As Short

            I = FileUsed.SelectedIndex
            If I > 0 Then
                FileUsed.SelectedIndex -= 1
                TabsFile.Load(FileRoot.SelectedItem & "\" & FileUsed.SelectedItem)
            End If
        End Sub
    End Class

    Public WithEvents ViewTools As New ClassViewTools()

    Class ClassViewTools
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Tools"
            InputGestureText = "Ctrl+T"
            GridMenu.Shortcut(Key.T, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' toont binair bestand

            GridText.Tools()
        End Sub
    End Class

    Public WithEvents ViewInfo As New ClassViewInfo()

    Class ClassViewInfo
        Inherits MenuItem

        Function Init()

            Header = "_Info"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' toont info

            MsgBox(FormGrid.ColumnDefinitions(0).ActualWidth.ToString & " x " & FormGrid.RowDefinitions(1).ActualHeight.ToString & vbCrLf &
               FormGrid.RowDefinitions(2).ActualHeight.ToString)
        End Sub
    End Class

    Public WithEvents ViewGoogle As New ClassViewGoogle()

    Class ClassViewGoogle
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Google"
            InputGestureText = "Ctrl+G"
            GridMenu.Shortcut(Key.G, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' opent Google window

            Dim P As String

            P = FileText.Text
            Interaction.Shell("C:\Program Files (x86)\Internet Explorer\iexplore.exe " & """" & "https://www.google.nl/search?q=" & P & """", vbNormalFocus)
        End Sub
    End Class
End Module