Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Input

Module UnitMenuEdit

    Public MenuEdit As New ClassMenuEdit()

    Class ClassMenuEdit
        Inherits MenuItem

        Function Init() As MenuItem

            ' start edit menu

            Header = "_Edit"

            Items.Add(EditExport.Init)
            Items.Add(EditInfo.Init)
            Items.Add(EditMaps.Init)
            Items.Add(EditPhoto.Init)
            Items.Add(EditReplace.Init)
            Items.Add(EditNotePad.Init)

            Return Me
        End Function
    End Class

    Public WithEvents EditExport As New ClassEditExport()

    Class ClassEditExport
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Export"
            InputGestureText = "Ctrl+W"
            GridMenu.Shortcut(Key.W, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' exporteert lego objeect

            ModelExportBin()
        End Sub
    End Class

    Public WithEvents EditInfo As New ClassEditInfo()

    Class ClassEditInfo
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Info"
            InputGestureText = "Ctrl+I"
            GridMenu.Shortcut(Key.I, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' maakt de standaard .txt bestanden in de huidige map

            Dim P As String = FileRoot.SelectedItem & "\" & GetFilePath(TextFile)

            If Not File.Exists(P & "boek.txt") Then File.WriteAllText(P & "boek.txt", "")
            If Not File.Exists(P & "data.txt") Then File.WriteAllText(P & "data.txt", "")
            If Not File.Exists(P & "doel.txt") Then File.WriteAllText(P & "doel.txt", "")
            If Not File.Exists(P & "gids.txt") Then File.WriteAllText(P & "gids.txt", "")
            If Not File.Exists(P & "info.txt") Then File.WriteAllText(P & "info.txt", "")
            If Not File.Exists(P & "plan.txt") Then File.WriteAllText(P & "plan.txt", "")
            If Not File.Exists(P & "test.txt") Then File.WriteAllText(P & "test.txt", "")
            If Not File.Exists(P & "werk.txt") Then File.WriteAllText(P & "werk.txt", "")
            FilePath.Update(P)
            FileFile.Update(P)
        End Sub
    End Class

    Public WithEvents EditMaps As New ClassEditMaps()

    Class ClassEditMaps
        Inherits MenuItem
        Function Init() As MenuItem

            Header = "_Maps"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            ' bewaart een materiaal map

            GridMapsSave()
        End Sub
    End Class

    Public WithEvents EditPhoto As New ClassEditPhoto()

    Class ClassEditPhoto
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Photo"

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            '  maakt een foto van het media grid

            GridMedia.Photo()
        End Sub
    End Class

    Public WithEvents EditReplace As New ClassEditReplace()

    Class ClassEditReplace
        Inherits MenuItem

        Function Init() As MenuItem

            Header = "_Replace"
            InputGestureText = "Ctrl+H"
            GridMenu.Shortcut(Key.H, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            '  vervangt tekst

            GridText.Replace()
        End Sub
    End Class

    Public WithEvents EditNotePad As New ClassEditNotepad()

    Class ClassEditNotepad
        Inherits MenuItem

        Function Init()

            Header = "_NotePad"
            InputGestureText = "Ctrl+K"
            GridMenu.Shortcut(Key.K, ModifierKeys.Control, AddressOf Me_Click)

            Return Me
        End Function

        Sub Me_Click() Handles Me.Click

            '  opent kladblok

            Interaction.Shell("C:\windows\system32\notepad.exe " & """" & TextFile & """", vbNormalFocus)
        End Sub
    End Class
End Module