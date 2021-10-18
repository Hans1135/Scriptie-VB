Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Microsoft.Win32

Module UnitMenuFile


    Public MenuFile As New ClassMenuFile()

    Class ClassMenuFile
        Inherits MenuItem

        Function Init() As MenuItem

            ' start file menu

            Header = "_File"

            Items.Add(FileNew.Init)
            Items.Add(FileOpens.Init)
            Items.Add(FileSave.Init)
            Items.Add(FileSaveAs.Init)
            Items.Add(FileFavorits.Init)
            Items.Add(FileLast.Init)
            Items.Add(FileExplorer.Init)
            Items.Add(FileExit.Init)

            Return Me
        End Function
    End Class

    Public WithEvents FileNew As New ClassFileNew()

    Class ClassFileNew
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_New"
            InputGestureText = "Ctrl+N"
            GridMenu.Shortcut(Key.N, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' maakt nieuw onderwerp

            Dim S As String

            S = InputBox("New Path")
            If S <> "" Then
                If Left(S, 1) = "\" Then
                    S = GetFilePath(TextFile) & Mid(S, 2)
                End If
                S = FileRoot.SelectedItem & "\" & S
                Directory.CreateDirectory(S)
                TabsFile.Load(S & "\info.txt")
                FilePath.Update(S)
                FileFile.Update(S)
            End If
        End Sub
    End Class

    Public WithEvents FileOpens As New ClassFileOpens

    Class ClassFileOpens
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Open"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' opent bestand

            Dim OpenDialog As New OpenFileDialog

            If OpenDialog.ShowDialog() = True Then
                TabsFile.Load(OpenDialog.FileName)
            End If
        End Sub
    End Class

    Public WithEvents FileSave As New ClassFileSave()

    Class ClassFileSave
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Save"
            InputGestureText = "Ctrl+S"
            GridMenu.Shortcut(Key.S, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' bewaart bestand en onthoudt laatste bestand

            GridText.Save()
            If TextFile <> "" Then
                File.WriteAllText(FileRoot.Items(0) & "\f\favoriet\laatst.txt", FileRoot.SelectedIndex & vbCrLf & TextFile)
            Else
                File.WriteAllText(FileRoot.Items(0) & "\f\favoriet\laatst.txt", FileRoot.SelectedIndex & vbCrLf & PlayFile)
            End If
        End Sub
    End Class

    Public WithEvents FileSaveAs As New ClassFileSaveAs()

    Class ClassFileSaveAs
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_SaveAs"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' hernoemt een bestand

            Dim SaveAsDialog As New SaveFileDialog With {.FileName = TextFile}

            If SaveAsDialog.ShowDialog = True Then
                File.Move(TextFile, SaveAsDialog.FileName)
                TextFile = SaveAsDialog.FileName
                FilePath.Update(TextFile)
                FileFile.Update(TextFile)
            End If
        End Sub
    End Class

    Public WithEvents FileFavorits As New ClassFileFavorits()

    Class ClassFileFavorits
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Favorits"
            InputGestureText = "Ctrl+R"
            GridMenu.Shortcut(Key.R, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' opent favorieten.txt

            TabsFile.Load(FileRoot.Items(0) & "\f\favoriet\favorieten.txt")
        End Sub
    End Class

    Public WithEvents FileLast As New ClassFileLast()

    Class ClassFileLast
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Last"
            InputGestureText = "Ctrl+L"
            GridMenu.Shortcut(Key.L, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' opent laatst gebruikte bestand

            Dim S As String()

            S = File.ReadAllLines(FileRoot.Items(0) & "\f\favoriet\laatst.txt")
            FileRoot.SelectedIndex = S(0)
            TabsFile.Load(S(1))
        End Sub
    End Class

    Public WithEvents FileExplorer As New ClassFileExplorer()

    Class ClassFileExplorer
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Explorer"
            InputGestureText = "Ctrl+E"
            GridMenu.Shortcut(Key.E, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' start explorer

            Dim F As String

            F = TextFile
            If Not File.Exists(F) Then ' als het bestand niet gevonden wordt
                F = GetFilePath(F) ' open alleen de map
            End If
            Interaction.Shell("C:\Windows\explorer.exe /select," & """" & F & """", vbMaximizedFocus)
        End Sub
    End Class

    Public WithEvents FileExit As New ClassFileExit()

    Class ClassFileExit
        Inherits MenuItem

        Function Init()

            Header = "_Exit"
            InputGestureText = "Alt+F4"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' stopt het programma

            MainForm.Close()
        End Sub
    End Class
End Module