Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading

Module UnitMediaPlay

    Public WithEvents MediaPlay As New ClassPlay

    Class ClassPlay
        Inherits MediaElement

        Function Init() As UIElement

            ' start de speler

            LoadedBehavior = MediaState.Manual
            PlayVolume.Value = 0.5
            HorizontalAlignment = HorizontalAlignment.Left
            VerticalAlignment = VerticalAlignment.Top
            Name = "Player"

            Return Me
        End Function

        Sub Done()

            ' stopt de speler

            MediaPlay.Source = Nothing ' wis bron
            GridMedia.Children.Remove(MediaPlay)
        End Sub

        Sub Load(S As String)

            ' laadt een muziek of video bestand

            If File.Exists(S) Then
                If GridMedia.Children.Contains(Me) = False Then GridMedia.Children.Add(Me)
                Stops()
                LoadedBehavior = MediaState.Manual
                Volume = 0.5
                PlayFile = S
                Source = New Uri(S)
                Play()
                PlayButton.Content = "pause"
            Else
                MsgBox(S & " niet gevonden")
            End If
        End Sub
        '
        Sub Start()

            ' speelt een bestand

            If PlayButton.Content = "start" Then
                PlayButton.Content = "pause"
                Play()
            Else
                PlayButton.Content = "start"
                Pause()
            End If
        End Sub
        '
        Sub Stops()

            ' stopt de speler

            MediaPlay.Stop()
            PlayTimer.Stop()
            PlayState = 0
            MediaPlay.Source = Nothing
        End Sub

        Sub Me_MediaOpened(sender As Object, e As RoutedEventArgs) Handles Me.MediaOpened

            'MsgBox("Open")
            PlayText.Text = GetFileName(PlayFile)
            PlayPosition.Maximum = MediaPlay.NaturalDuration.TimeSpan.TotalSeconds
            Position = TimeSpan.FromSeconds(PlayFrom)
            If PlayTill = 0 Then PlayTill = PlayPosition.Maximum
            PlayTimer.Interval = TimeSpan.FromSeconds(1)
            PlayTimer.Start()
            PlayFileTags(PlayFile)
            TabsPlay.IsSelected = True
        End Sub
    End Class

    Sub PlayTimerTick() Handles PlayTimer.Tick

        Dim D As Double

        PlayState = 3
        D = MediaPlay.Position.TotalSeconds
        PlayPosition.Value = D
        PositionLabel.Content = "position " & Format(D \ 60, "00") & ":" & Format(D Mod 60, "00")
        PlayState = 1
        'If PlayPosition.Value >= PlayTill Then
        'MediaPlayStop()
        'End If
    End Sub

    Sub PlayPosition_ValueChanged() Handles PlayPosition.ValueChanged

        If PlayState < 3 Then
            MediaPlay.Position = TimeSpan.FromSeconds(PlayPosition.Value)
            PositionLabel.Content = "position " & Format(PlayPosition.Value, "0")
        End If
    End Sub

    Sub PlayVolume_ValueChanged() Handles PlayVolume.ValueChanged

        MediaPlay.Volume = PlayVolume.Value
        VolumeLabel.Content = "volume " & Format(PlayVolume.Value * 100, "0")
    End Sub

    Sub PlayButton_Click() Handles PlayButton.Click

        MediaPlay.Start()
    End Sub

    Public PlayState As Short
    Public PlayFile As String
    Public PlayFrom As Single
    Public PlayTill As Single
    Public WithEvents PlayTimer As New DispatcherTimer()
    Public PlayMusician As New TextBox
    Public PlayYear As New TextBox
    Public PlayAlbum As New TextBox
    Public PlayNumber As New TextBox
    Public PlayTrack As New TextBox
    Public PlayCoworker As New TextBox
    Public PlayTime As New TextBox
    Public PlayText As New TextBox
    Public PositionLabel As New Label
    Public WithEvents PlayPosition As New Slider
    Public VolumeLabel As New Label
    Public WithEvents PlayVolume As New Slider
    Public WithEvents PlayButton As New Button
End Module
