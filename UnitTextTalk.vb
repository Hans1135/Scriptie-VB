Imports System.Speech.Synthesis
Imports System.Windows
Imports System.Windows.Controls
'
Module UnitTextTalk

    Public TextTalk As New ClassTextTalk

    Class ClassTextTalk

        Public TalkSynth As New SpeechSynthesizer() ' kan niet vererft worden

        Sub Init()

            ' start een stem

            ' W(0) = "talk"
            ' W(1) = stem
            ' V(2) = volume

            If W(1) = "" Then ' als er geen stem is opgegeven

                Dim L As New ListBox
                Dim S As String = ""
                Dim F As New Window

                For Each Voice In TalkSynth.GetInstalledVoices
                    L.Items.Add(Voice.VoiceInfo.Name)
                Next
                F.Height = 300
                F.Width = 400
                F.Content = L
                F.Title = "Kies een stem"
                F.Show()
                While S = "" ' wacht op een keuze
                    DoEvents() ' is nodig
                    If L.SelectedIndex > -1 Then S = L.SelectedItem ' als een stem gekozen is
                End While
                F.Close()
                TalkSynth.SelectVoice(S)
            Else ' als er wel een stem is opgegeven
                W(1) = LCase(W(1))
                If W(1) = "david" Then
                    TalkSynth.SelectVoice("Microsoft David Desktop")
                ElseIf W(1) = "zira" Then
                    TalkSynth.SelectVoice("Microsoft Zira Desktop")
                Else
                    TalkSynth.SelectVoice("Microsoft Server Speech Text to Speech Voice (nl-NL, Hanna)")
                End If
            End If
            If W(2) = "" Then V(2) = 50
            TalkSynth.Volume = V(2)
        End Sub

        Sub Say()

            ' spreekt een opgegeven tekst

            ' W(0) = "say"
            ' W(1) = text

            Dim I As Short

            I = 2 ' als er komma's in de regel stonden worden de woorden weer aan elkaar gemaakt
            While W(I) <> "" ' zolang de woorden niet leeg zijn
                W(1) &= ", " & W(I) ' voeg dat woord toe aan woord 1
                I += 1
            End While

            TalkSynth.Speak(W(1)) ' returned pas na tekst gesproken is, speakcompleted event is daarom niet nodig
        End Sub
    End Class
End Module
