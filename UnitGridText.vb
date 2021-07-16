Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media

Module UnitGridText

    Public GridText As New ClassGridText

    Class ClassGridText
        Inherits TextBox

        Function Init() As UIElement

            ' start het tekst veld

            FontFamily = New FontFamily("Consolas")
            FontSize = 14
            AcceptsReturn = True
            AcceptsTab = True
            TextWrapping = TextWrapping.Wrap
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto

            Return Me
        End Function

        Sub Done()

            ' stopt het tekst veld

            Save()
        End Sub

        Sub Load(S As String)

            ' laadt een tekst bestand

            Dim T As String

            Save() ' bewaar vorige bestand

            If File.Exists(S) Then ' als het bestand bestaat
                TextFile = S
                Text = File.ReadAllText(S)
                TextSaved = True
                Show()
                If GetFileType(S) = ".txt" Then
                    S = FileRoot.SelectedItem & "\" & GetFilePath(S)
                    FilePath.Update(S)
                    FileFile.Update(S)
                    FileUsed.Update()
                End If
            Else ' als bestand niet bestaat
                If MsgBox("Bestand" & vbCrLf & S & vbCrLf & "niet gevonden. Maken?", vbYesNo) = vbYes Then
                    T = GetFilePath(S)
                    Directory.CreateDirectory(FileRoot.SelectedItem & "\" & T)
                    File.WriteAllText(S, "")
                    TabsFile.Load(S)
                Else
                    ParsSkip = True
                End If
            End If
        End Sub

        Sub Replace()

            ' vervangt tekst

            Dim I As Short
            'Dim J As Short
            Dim K As Short
            Dim N As Short
            'Dim S() As String
            Dim T() As String
            'Dim B As String
            'Dim D As String

            Save()
            'S = File.ReadAllLines(FileRoots.SelectedItem & "f\favorieten\vervang.txt")
            'K = UBound(S)
            T = File.ReadAllLines(TextFile)
            N = UBound(T)
            For I = 0 To K
                '    If S(I) <> "" Then
                '   J = InStr(S(I), ",")
                '  B = Left(S(I), J - 1)
                ' D = Mid(S(I), J + 1)
                'For J = 0 To N
                'T(J) = Strings.Replace(T(J), B, D)
                'Next
                '   End If
            Next
            Clear()
            For I = 0 To N
                AppendText(T(I) & vbCrLf)
            Next
        End Sub

        Sub Save()

            ' bewaart een tekst bestand

            If Not TextSaved And TextFile <> "" Then
                File.WriteAllText(TextFile, Text)
                TextSaved = True
                Show()
            End If
        End Sub

        Sub Show()

            ' toont de naam van het tekst bestand, eventueel met een sterretje als het bestand opgeslagen moet worden

            If TextSaved Then
                MainForm.Title = TextFile
            Else
                MainForm.Title = TextFile & "*"
            End If
        End Sub

        Sub Tools()

            ' toont de binaire inhoud van een file

            'Dim B() As Byte
            Dim L As Long
            Dim S As String

            Save()

            'B = File.ReadAllBytes(PictFile)
            'L = B.Length \ 16 - 1
            If L > 4095 Then L = 4095
            GridText.Clear()
            For I = 0 To L
                S = HW(I * 16) & " "
                For J = 0 To 15
                    'S &= HB(B(I * 16 + J)) & " "
                Next
                S &= " | "
                For J = 0 To 15
                    'If B(I * 16 + J) > 19 Then
                    'S &= Chr(B(I * 16 + J))
                    'Else
                    'S &= " "
                    'End If
                Next
                AppendText(S & vbCrLf)
            Next
        End Sub

        Sub Me_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseRightButtonUp

            ' selecteert een woord in een tekst

            Dim F As String = ""
            Dim I As Long
            Dim J As Long
            Dim K As Long
            Dim L As String = ""
            Dim N As Long
            Dim P As String
            Dim Q As String = ""
            Dim R As String = ""
            Dim T As String

            T = GridText.Text ' alle tekst
            N = Len(T) - 1 ' aantal karakters in de tekst
            I = GridText.SelectionStart ' plaats van cursor
            If I < N Then ' als cursor niet aan het eind van de tekst staaat
                J = I ' zoek het begin van het woord vanaf de cursor
                While J > 0 And T(J) > " " ' zolang niet het begin van de tekst of een spatie gevonden wordt
                    J -= 1
                End While
                K = I ' zoek het eind van het woord vanaf de cursor
                While K < N And T(K) > " " ' zolang niet het eind van de tekst of een spatie gevonden wordt
                    K += 1
                End While
                For I = J To K ' doorzoek de tekens van het woord  
                    If Asc(T(I)) > 31 Then  ' als het een normaal teken betreft
                        F &= T(I) ' voeg het teken toe aan het woord
                    End If
                Next
                F = Trim(F) ' wis eventuele spaties aan het begin en eind van het woord
                If Right(F, 1) = "." Or Right(F, 1) = "," Then F = Left(F, Len(F) - 1) ' als laatste letter een punt of komma is
                If F = "" Then Exit Sub ' als het woord leeg is stop procedure
                If InStr(F, ":") > 0 Then ' als een schijf bekent is
                    If Left(F, 4) = "http" Then
                        Microsoft.VisualBasic.Shell("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & F, AppWinStyle.NormalFocus)
                    Else
                        TabsFile.Load(F)
                    End If
                Else
                    R = FileRoot.SelectedItem & "\" ' bronpad
                    T = GetFileType(F)
                    If T = "" Then T = ".txt" ' is het filetype een tekstfile
                    If Left(F, 1) = "\" Then ' dwing vanaf het huidige pad te zoeken
                        L = ""
                        P = GetFilePath(TextFile)
                        F = Mid(F, 2)
                        P &= GetFilePath(F)
                        F = GetFileName(F)
                        If F = "" Then F = "info"
                    Else ' zoek vanaf bron pad
                        P = GetFilePath(F)
                        If P = "" Then P = F & "\"
                        F = GetFileName(F)
                        L = Left(P, 1) & "\"
                        If File.Exists(R & L & P & F & T) Then

                        ElseIf File.Exists(R & L & P & "boek.txt") Then
                            F = "boek"
                        Else
                            F = "info"
                        End If
                    End If
                    TabsFile.Load(R & L & P & F & T)
                End If
            End If
            e.Handled = True
        End Sub

        Sub Me_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown

            ' als op een toets gedrukt wordt

            Dim I As Long
            Dim N As Short = 4
            Dim T As New String(" ", N)

            If WaitKey Then
                WaitKey = False
                e.Handled = True
                Exit Sub
            End If
            If e.Key = Key.Tab Then
                e.Handled = True
                I = GridText.CaretIndex
                GridText.Text = GridText.Text.Insert(I, T)
                GridText.CaretIndex = I + N
            End If
        End Sub

        Sub Me_TextChanged() Handles Me.TextChanged

            ' als tekst veranderd

            TextSaved = False ' tekst moet opgeslagen worden
            Show() ' toon filenaam met sterretje
        End Sub
    End Class

    Function SF(F As Double) As String

        ' string format aanpassen

        Dim S As String

        S = Format(F, "0.0000")
        While Len(S) < 8
            S = " " & S
        End While
        SF = "," & S
    End Function

    Function LB(I As Short) As String

        ' maakt een string van 4 karakters lang die begint met een komma

        Dim S As String

        S = I
        While Len(S) < 4
            S = " " & S
        End While
        LB = "," & S
    End Function

    Function LE(I As Short) As String

        ' maakt een string van 4 karakters lang die eindigt met een komma

        Dim S As String

        S = I
        While Len(S) < 4
            S = " " & S
        End While
        LE = S & ", "
    End Function

    Function HB(B As Byte) As String

        ' maakt een byte hexadecimaal

        Dim S As String

        S = Hex(B)
        If Len(S) = 1 Then S = "0" & S

        HB = S
    End Function

    Function HW(N As UShort) As String

        ' maakt een word hexadecimaal

        Dim I As Byte
        Dim J As Byte

        I = N \ 256
        J = N Mod 256

        HW = HB(I) & HB(J)
    End Function

    Public TextFile As String = "" ' adres van tekstbestand
    Public TextSaved As Boolean = True
End Module