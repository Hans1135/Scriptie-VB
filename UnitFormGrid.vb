Imports System.Windows
Imports System.Windows.Controls

Module UnitFormGrid

    Public FormGrid As New ClassFormGrid

    Class ClassFormGrid
        Inherits Grid

        Function Init() As UIElement

            For I = 0 To 4  ' maakt 5 kolommen
                ColumnDefinitions.Add(New ColumnDefinition)
            Next
            For I = 0 To 3 ' maakt eerste 4 kolommen 450 breed
                ColumnDefinitions(I).Width = New GridLength(450)
            Next
            For I = 1 To 3  ' maakt 3 rijen
                RowDefinitions.Add(New RowDefinition)
            Next
            RowDefinitions(0).Height = GridLength.Auto ' menu rij
            RowDefinitions(1).Height = New GridLength(500) ' maakt rij 1 500 hoog
            Children.Add(GridMenu.Init())
            Children.Add(GridTabs.Init())
            Children.Add(GridText.Init())
            Children.Add(GridMedia.Init())
            Modes(1)
            Return Me
        End Function

        Sub Done()

            GridText.Done()
            PlayMidiDone()
        End Sub

        Sub Modes(I As Short)

            If I = 1 Then
                Move(GridMenu, 0, 0, 6, 1)
                GridTabs.Visibility = Visibility.Visible
                Move(GridTabs, 4, 1, 2, 2)
                GridText.Visibility = Visibility.Visible
                Move(GridText, 2, 1, 2, 2)
                Move(GridMedia, 0, 1, 2, 1)
            ElseIf I = 2 Then
                Move(GridMedia, 0, 1, 2, 2)
            ElseIf I = 3 Then ' half beeld
                Move(GridText, 4, 1, 1, 2)
                Move(GridMedia, 0, 1, 4, 2)
            ElseIf I = 4 Then ' vol beeld
                GridText.Visibility = Visibility.Hidden
                Move(GridMedia, 0, 1, 6, 2)
            End If
            Mode = I
        End Sub

        Sub Move(E As UIElement, X As Short, Y As Short, W As Short, H As Short)

            SetColumn(E, X)
            SetRow(E, Y)
            SetColumnSpan(E, W)
            SetRowSpan(E, H)
        End Sub

        Sub Change()

            If Mode < 4 Then Mode += 1 Else Mode = 1
            Modes(Mode)
        End Sub

        Sub Pars()

            ' W(0) = "mode"
            ' V(1) = mode

            Modes(V(1))
        End Sub

        Dim Mode As Short ' onthoudt veld indeling
    End Class
End Module