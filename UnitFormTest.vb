Imports System.Windows
Imports System.Windows.Controls

Module UnitFormTest

    Public TabsTest As New ClassTest

    Class ClassTest
        Inherits TabItem


        Function Init() As UIElement

            TestLabel.Content = "GridLines"
            TestStack.Children.Add(TestLabel)
            TestStack.Children.Add(TestCheck)

            Content = TestStack
            Header = "Test"

            Return Me
        End Function

        Sub Test()

            FormGrid.ShowGridLines = TestCheck.IsChecked
        End Sub

        Public TestLabel As New Label
        Public Shared TestCheck As New CheckBox
        Public TestStack As New StackPanel
    End Class
End Module
