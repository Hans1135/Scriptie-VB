Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Module UnitGridMenu

    Public GridMenu As New ClassGridMenu

    Class ClassGridMenu
        Inherits Menu

        Function Init() As UIElement

            Items.Add(MenuFile.Init)
            Items.Add(MenuEdit.Init)
            Items.Add(MenuView.Init)
            Items.Add(MenuStart.Init)
            Return Me
        End Function

        Sub Shortcut(K As Key, M As ModifierKeys, E As ExecutedRoutedEventHandler)

            Dim C As New RoutedCommand()

            If M = vbNull Then
                C.InputGestures.Add(New KeyGesture(K))
            Else
                C.InputGestures.Add(New KeyGesture(K, M))
            End If
            MainForm.CommandBindings.Add(New CommandBinding(C, E))
        End Sub
    End Class
End Module