Imports System.Windows
Imports System.Windows.Controls

Module UnitTabsTest

    Public TabsTest As New ClassTabsTest

    Class ClassTabsTest
        Inherits TabItem

        Function Init() As UIElement

            Content = TestStack.Init
            Header = "Test"
            Return Me
        End Function

        Sub Test()

            ' geen functie

        End Sub
    End Class

    Public TestStack As New ClassTestStack

    Class ClassTestStack
        Inherits StackPanel

        Function Init() As StackPanel

            TestStack.Children.Add(TestText)
            TestForm.Content = "FormSize"
            TestStack.Children.Add(TestForm)
            TestGrid.Content = "GridLines"
            TestStack.Children.Add(TestGrid)
            Return Me
        End Function

        Sub TestForm_Checked() Handles TestForm.Checked

            TestText.Text = MainForm.ActualWidth \ 1 & "x" & MainForm.ActualHeight \ 1
        End Sub

        Sub TestGrid_Click() Handles TestGrid.Click

            FormGrid.ShowGridLines = TestGrid.IsChecked
        End Sub

        Public TestText As New TextBox
        Public WithEvents TestGrid As New CheckBox
        Public WithEvents TestForm As New CheckBox
    End Class
End Module
