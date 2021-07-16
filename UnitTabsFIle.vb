Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Module UnitTabsFIle

    Public TabsFile As New ClassTabsFile

    Class ClassTabsFile
        Inherits TabItem

        Function Init() As UIElement

            ' start bestanden veld

            FileStack.Children.Add(FileText)
            FileStack.Children.Add(RootsLabel)
            FileStack.Children.Add(FileRoot.Init)
            FileStack.Children.Add(PathsLabel)
            FileStack.Children.Add(FilePath.Init)
            FileStack.Children.Add(FilesLabel)
            FileStack.Children.Add(FileFile.Init)
            FileStack.Children.Add(UsedsLabel)
            FileStack.Children.Add(FileUsed.Init)

            Content = FileStack
            Header = "Files"

            Return Me
        End Function

        Sub Load(S As String)

            ' laadt een bestand

            Select Case GetFileType(S)
                Case ".txt", ".srt", ".dat", ".ldr", ".mtl", ".vb", ".bvh"
                    GridText.Load(S)
                Case "gif", ".jpg", ".png", ".dds", ".tga", ".tif"
                    MediaPict.Load(S)
                Case ".mp3", ".avi", ".mp4", ".mid", ".flv", ".mpg"
                    MediaPlay.Load(S)
                Case ".obj", ".dae", ".fbx", ".stl"
                    MediaScene.Load(S)
                Case ".url", ".lnk"
                    Interaction.Shell("c:\windows\explorer.exe " & S, AppWinStyle.NormalFocus)
            End Select
        End Sub
    End Class

    Public FileText As New ClassFileText

    Class ClassFileText
        Inherits TextBox


        Sub Me_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp

            ' zoekt muzikant

            Dim S As String

            If e.Key = Key.Enter Then
                S = FileText.Text
                FileRoot.SelectedIndex = 2
                TabsFile.Load("d:\onedrive\muziek\" & GetArtist(S) & "albums.txt")
            End If
        End Sub
    End Class

    Public FileRoot As New ClassFileRoot

    Class ClassFileRoot
        Inherits ListBox

        Function Init() As ListBox

            ' vult de lijst met bron mappen

            RootsLabel.Content = "Roots"
            MaxHeight = 200
            Update()

            Return Me
        End Function

        Sub Update()

            Dim I As Short
            Dim S = File.ReadAllLines("D:\OneDrive\Index\b\bron\data.txt")
            Dim N = UBound(S)

            For I = 0 To N
                Items.Add(S(I))
            Next
            SelectedIndex = 0
        End Sub

        Sub Me_MouseLeftButtonUp() Handles Me.MouseLeftButtonUp

            ' selecteert een bron pad

            GridText.Save()
            GridText.Clear()
            TextFile = ""

            FormGrid.Modes(1)
            FileUsed.Items.Clear()
            FilePath.Update(FileRoot.SelectedItem)
            FileFile.Update(FileRoot.SelectedItem)
        End Sub
    End Class

    Public FilePath As New ClassFilePath

    Class ClassFilePath
        Inherits ListBox

        Function Init() As ListBox

            PathsLabel.Content = "Paths"
            MaxHeight = 200

            Return Me
        End Function

        Sub Update(S As String)

            ' vult de padlijst opnieuw

            Dim L As Short
            Dim T As String

            L = Len(FileRoot.SelectedItem) + 2 ' lengte van het bronpad
            If FilePath.SelectedItem = Nothing Then ' als er geen selectie in de padlijst is gemaakt
                T = Directory.GetParent(Directory.GetParent(FileRoot.SelectedItem & "\" & GetFilePath(TextFile)).ToString).ToString
                T = Mid(T, L)
            Else
                T = FilePath.SelectedItem
                T = GetPrevPath(T)
            End If
            FilePath.Items.Clear()
            FilePath.Items.Add(T) ' eerste pad om terug te gaan naar ouderpad
            For Each D In Directory.GetDirectories(S)
                FilePath.Items.Add(Mid(D, L))
            Next
        End Sub

        Sub Me_MouseUp() Handles Me.MouseUp

            ' als een pad geselecteerd wordt

            Dim P As String

            P = FileRoot.SelectedItem & "\" & SelectedItem
            FilePath.Update(P)
            FileFile.Update(P)
        End Sub
    End Class

    Public FileFile As New ClassFileFile

    Class ClassFileFile
        Inherits ListBox

        Function Init() As ListBox

            FilesLabel.Content = "Files"
            MaxHeight = 200

            Return Me
        End Function

        Sub Update(S As String)

            ' vult de bestandenlijst opnieuw

            Dim L As Short

            L = Len(FileRoot.SelectedItem) + 2 ' lengte van het geselecteerde bron pad
            Me.Items.Clear() ' maak de lijst leeg
            For Each F In Directory.GetFiles(S) ' zoek bestanden in het pad
                Me.Items.Add(Mid(F, L))
            Next
            If File.Exists(S & "\info.jpg") Then
                MediaPict.Load(S & "\info.jpg")
            Else
                MediaPict.Load(S & "\folder.jpg")
            End If
        End Sub

        Sub Me_KeyUp() Handles Me.KeyUp

            ' selecteert een bestand

            Me_MouseLeftButtonUp()
        End Sub

        Sub Me_MouseLeftButtonUp() Handles Me.MouseLeftButtonUp

            ' selecteert een bestand

            TabsFile.Load(FileRoot.SelectedItem & "\" & SelectedItem)
        End Sub

        Sub Me_MouseRightButtonUp() Handles Me.MouseRightButtonUp

            ' verwijdert bestand

            Dim S As String = FileRoot.SelectedItem & "\" & SelectedItem
            Dim P As String = FileRoot.SelectedItem & "\" & GetFilePath(TextFile)

            If MsgBox(S & " verwijderen?", vbYesNo) = vbYes Then
                FileSystem.DeleteFile(S, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin, UICancelOption.ThrowException)
                FilePath.Update(P)
                FileFile.Update(P)
            End If
        End Sub
    End Class

    Public FileUsed As New ClassFileUsed

    Class ClassFileUsed
        Inherits ListBox

        Function Init() As ListBox

            UsedsLabel.Content = "Useds"
            MaxHeight = 200

            Return Me
        End Function

        Sub Update()

            ' vernieuwt de gebruikte bestanden lijst

            Dim J As Short
            Dim L As Short
            Dim N As Short
            Dim T As String

            L = Len(FileRoot.SelectedItem) + 2
            N = Me.Items.Count - 1
            J = Me.SelectedIndex + 1
            T = LCase(Mid(GetFileAddress(TextFile), L))
            For I = 0 To N ' kijk of het bestand al eerder is gebruikt
                If LCase(FileUsed.Items(I)) = T Then
                    FileUsed.SelectedIndex = I
                    GoTo verder
                End If
            Next
            FileUsed.Items.Insert(J, T)
            FileUsed.SelectedIndex = J
verder:
        End Sub

        Sub Me_MouseUp() Handles Me.MouseUp

            ' selecteert een gebruikt bestand

            TabsFile.Load(FileRoot.SelectedItem & "\" & SelectedItem)
        End Sub
    End Class

    Sub FileMakeBackup()

        ' maakt reserve copie

        ' W(0) = "backup"
        ' W(1) = bron pad
        ' W(2) = sub pad
        ' W(3) = masker "*.*" e.d.
        ' W(4) = doel pad

        'Dim S As String

        'If Directory.Exists(W(1) & W(2)) Then
        'For Each D In Directory.GetDirectories(W(1) & W(2))
        'S = GetFileName(D) & "\"
        'MakeBackupFiles(S)
        'Next
        'Else
        'MsgBox("map " & W(1) & W(2) & " niet gevonden")
        'End If
    End Sub

    'Sub MakeBackupFiles(T As String)

    'Dim S As String

    'For Each F In Directory.GetFiles(W(1) & W(2) & T, W(3))
    'S = W(4) & W(2) & T & GetFileName(F) & GetFileType(F)
    'Dim B As New FileInfo(F)
    'Dim D As New FileInfo(S)
    'If Directory.Exists(W(4) & W(2) & T) = False Then Directory.CreateDirectory(W(4) & W(2) & T)
    'If B.LastWriteTime > D.LastWriteTime Then File.Copy(F, S)
    'Next
    'End Sub

    Sub FileMakeDirs()

        ' maakt een map
        ' W(0) = "makedirs"
        ' W(1) = mapnaam

        'Dim P As String

        'If InStr(W(1), ":") = 0 Then
        'P = FileRoots.SelectedItem & W(1)
        'Else
        'P = W(1)
        'End If

        'If Directory.Exists(P) = False Then
        'Directory.CreateDirectory(P)
        'Else
        '' MsgBox(S & " bestaat al")
        'End If
    End Sub

    Sub FileMakeFile()

        ' maakt een bestand
        ' W(0) = "makefile"
        ' W(1) = filenaam

        'Dim P As String
        'Dim S As String

        'If InStr(W(1), ":") = 0 Then
        'P = FileRoots.SelectedItem & GetFilePath(W(1))
        'S = FileRoots.SelectedItem & W(1)
        'Else
        'P = GetFilePath(W(1))
        'S = W(1)
        'End If
        'If File.Exists(S) = False Then ' het bestand wordt alleen gemaakt als het bestand nog niet bestaat
        'If Directory.Exists(P) = False Then Directory.CreateDirectory(P)
        'File.WriteAllText(S, "")
        'Else
        'MsgBox(S & " bestaat al")
        'End If
    End Sub

    Function GetFileAddress(S As String) As String

        ' bepaalt fileroot\filepath\filenaam.filetype

        Dim R As String = FileRoot.SelectedItem & "\" ' bronpad
        Dim T As String

        If GetFileType(S) = "" Then S &= ".txt"

        If InStr(S, ":") > 0 Then ' als in het adres een schijf voorkomt
            T = S
        ElseIf Left(S, 1) = "\" Then ' als adres van een vervolgpad is
            T = R & GetFilePath(ParsFile(FI)) & Mid(S, 2)
        Else ' als adres vanaf het bronpad is
            T = R & Left(S, 1) & "\" & S
        End If

        Return T
    End Function

    Function GetFileName(S As String) As String

        ' bepaalt de filenaam van een FileAddress als in fileroot\filepath\filenaam.filetype

        Dim I As Short

        I = InStrRev(S, "\") + 1 ' verwijder het filepad
        S = Mid(S, I)
        I = InStr(S, ".") - 1 ' verwijder het filetype
        If I > 0 Then S = Strings.Left(S, I)

        GetFileName = S
    End Function

    Function GetFilePath(S As String) As String

        ' bepaalt het pad

        Dim I As Short

        If Left(S, 1) = "@" Then
            If InStr(S, "\") > 0 Then
                S = Mid(S, 2, 1) & "\" & Mid(S, 2)
            Else
                S = ""
            End If
        End If
        If InStr(S, ":") > 0 Then ' staat er een schijf in het pad
            S = Mid(S, Len(FileRoot.SelectedItem) + 2)
        End If
        I = InStrRev(S, "\") ' verwijder de filenaam
        S = Left(S, I)

        Return S
    End Function

    Function GetFileType(S As String) As String

        ' bepaalt het filetype

        Dim I As Short

        I = InStrRev(S, ".")

        If I > 0 And Len(S) - I > 1 Then
            GetFileType = LCase(Mid(S, I))
        Else
            GetFileType = ""
        End If
    End Function

    Function GetPrevPath(S As String) As String

        ' bepaalt de vorige map

        S = Left(S, InStrRev(Left(S, Len(S) - 1), "\"))

        Return S
    End Function

    Public RootsLabel As New Label
    Public PathsLabel As New Label
    Public FilesLabel As New Label
    Public UsedsLabel As New Label
    Public FileStack As New StackPanel
End Module
