Imports System.IO
Imports System.Windows

Module UnitPlayMusic

    Sub PlayFindTrack()

        ' W(0) = "track"
        ' W(1) = bestand

        Dim A As String ' album
        Dim E As Collections.Generic.IEnumerable(Of String)
        Dim F As String ' bestand naam
        Dim J As String ' jaar
        Dim M As String ' muzikant
        Dim N As Short ' aantal
        Dim P As String ' pad
        Dim R As String
        Dim S As String() ' bestand regels
        Dim T As String ' track

        J = Mid(W(1), 12, 4)
        S = File.ReadAllLines(W(1)) ' lees lijst bestand
        N = UBound(S) ' aantal regels in lijst bestand
        For I = 0 To N ' voor alle regels
            TextParsLine(S(I), ",") ' haal de woorden uit de regele
            If W(2) <> "" Then ' als woord 2 een muzikant naam heeft
                M = W(2) ' onthoudt muzikant
                FileText.Text = M ' zet muzikant naam in FileText, kan met enter geopend worden of met ctrl+g gegoogled
                If W(3) = "" Then A = "#" Else A = W(3) ' als woord 3 een album naam heeft
                T = W(4) ' single naam
                P = "D:\OneDrive\Muziek\" & GetArtist(M)
                If Not Directory.Exists(P) Then ' als muzikant nog geen map heeft
                    If MsgBox("muzikant " & M & " niet gevonden, maken?", vbYesNo) = vbYes Then ' maak map
                        W(1) = M
                        PlayMakeArtist()
                    End If
                End If
                E = Directory.EnumerateDirectories(P, "*" & A) ' zoek album
                F = ""
                For Each D In E
                    F = D
                Next
                If F = "" Then ' als album nog niet bestaat bestaat
                    MsgBox("album " & A & " van " & M & " niet gevonden")
                    FileRoot.SelectedIndex = 2 ' 
                    R = "D:\OneDrive\Muziek\" & GetArtist(M) & "albums.txt"
                    GridText.Load(R)
                    If GridText.Text = "" Then
                        GridText.AppendText("pars, album" & vbCrLf)
                    End If
                    If A = "#" Then A = "#-" & J & "-01 - #" Else A = "A-" & J & "-01 - " & A
                    If T = "" Then T = "\#" Else T = "\S-" & J & "-01 - " & T
                    Clipboard.SetText(A & T & vbCrLf)
                    ParsSkip = True
                    Exit Sub
                Else
                    E = Directory.EnumerateDirectories(F, "*" & T & "*")
                    F = ""
                    For Each D In E
                        F = D
                    Next
                    If F <> "" Then
                        E = Directory.EnumerateFiles(F, "*.mp3") ' zoek mp3 bestanden
                        F = ""
                        For Each G In E
                            F = Path.GetFileName(G)
                            P = Path.GetDirectoryName(G)
                            If Mid(F, 4, 3) <> " - " Then
                                PlayFileTags(G)
                            End If
                        Next
                        If F = "" Then
                            If MsgBox("track " & T & " van " & M & " niet gevonden, zoeken?", vbYesNo) = vbYes Then
                                Interaction.Shell("C:\Windows\explorer.exe /select," & """" & P & """", vbMaximizedFocus)
                                Wait(1)
                                P = "F:\muziek-oud\" & GetArtist(M)
                                Interaction.Shell("C:\Windows\explorer.exe /select," & """" & P & """", vbMaximizedFocus)
                                ParsSkip = True
                                Exit Sub
                            End If
                        End If
                        E = Directory.EnumerateFiles(P, "*.jpg") ' zoek folder.jpg
                        F = ""
                        For Each G In E
                            F = Path.GetFileName(G)
                            If F <> "folder.jpg" Then
                                P = Path.GetDirectoryName(G)
                                File.Move(P & "\" & F, P & "\folder.jpg")
                            End If
                        Next
                        If F = "" Then
                            MsgBox("geen hoes gevonden")
                            Clipboard.SetText(M & " - " & T)
                            Interaction.Shell("C:\Windows\explorer.exe /select," & """" & P & """", vbMaximizedFocus)
                            ParsSkip = True
                            Exit Sub
                        End If
                    Else
                        MsgBox("single" & vbCrLf & T & vbCrLf & "van" & vbCrLf & M & vbCrLf & "niet gevonden")
                        FileRoot.SelectedIndex = 2
                        R = "D:\OneDrive\Muziek\" & GetArtist(M)
                        GridText.Load(R & "albums.txt")
                        If A = "#" Then
                            A = Mid(Directory.GetDirectories(R, "#-*")(0), Len(R) + 1)
                        Else
                            A = "A-" & J & "-01 - " & A
                        End If
                        T = "\S-" & J & "-01 - " & T
                        Clipboard.SetText(A & T & vbCrLf)
                        ParsSkip = True
                        Exit Sub
                    End If
                End If
            End If
        Next
    End Sub

    Sub PlayMakeAlbum()

        ' W(0) = "album"
        ' W(1) = pad

        Dim I As Short

        If W(1) <> "" Then
            I = 2
            While W(I) <> "" ' in het geval dat er komma's in de album naam staan
                W(1) &= ", " & W(I)
                I += 1
            End While
            W(1) = GetFilePath(ParsFile(FI)) & W(1)
            FileMakeDirs()
        End If
    End Sub

    Sub PlayMakeArtist()

        ' W(0) = "artist"
        ' W(1) = naam

        If W(1) <> "" Then

            W(1) = "D:\OneDrive\Muziek\" & GetArtist(W(1)) & "albums.txt"
            If File.Exists(W(1)) = False Then
                W(0) = "makeFile"
                FileMakeFile()
            Else

            End If
        End If
    End Sub

    Function GetArtist(A As String) As String

        ' bepaalt het pad voor een muzikant

        Dim L1 As String
        Dim L2 As String

        L1 = UCase(Left(A, 1))
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", L1) = 0 Then
            L2 = L1
            L1 = "#"
        Else
            L2 = UCase(Mid(A, 2, 1))
            If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", L2) = 0 Then L2 = "#"
        End If

        A = L1 & "\" & L1 & L2 & "\" & A & "\"

        Return A
    End Function


    Sub PlayFileTags(S As String)

        ' leest het nummer, de titel, meewerkend artiesten en de lengte van een mp3 file

        Dim F As TagLib.File
        Dim B As String ' bron
        Dim T As String ' target

        If GetFileType(S) = ".mp3" Then
            F = TagLib.File.Create(S)
            PlayAlbum.Text = F.Tag.Album
            PlayTime.Text = Left(F.Properties.Duration.ToString, 8)
            PlayCoworker.Text = F.Tag.JoinedPerformers
            PlayMusician.Text = F.Tag.FirstAlbumArtist
            PlayNumber.Text = F.Tag.Track
            PlayTrack.Text = F.Tag.Title
            PlayYear.Text = F.Tag.Year
            B = GetFileName(S)
            T = PlayNumber.Text & " - " & PlayTrack.Text
            If B <> T Then
                If MsgBox("verander " & vbCrLf & B & vbCrLf & " in " & vbCrLf & T, vbYesNo) = vbYes Then
                    B = FileRoot.SelectedItem & "\" & GetFilePath(S)
                    T = B & T & ".mp3"
                    Directory.Move(S, T)
                    PlayFile = T
                    PlayText.Text = GetFileName(T)
                    FilePath.Update(B)
                    FileFile.Update(B)
                End If
            End If
        End If
    End Sub
End Module
