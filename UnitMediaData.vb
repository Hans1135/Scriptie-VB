Imports System.Data
Imports System.IO
Imports System.Windows.Controls

Imports Scriptie.ClassMedia

Module UnitMediaData



    Public MediaData As New ClassMediaData

    Class ClassMediaData
        Inherits DataGrid

        Sub Init()

            ' start een database

            TI = -1
            CI = 0
            RI = 0

            TableSet = New DataSet(W(1))

            TableList.Items.Clear()

            If GridMedia.Children.Contains(MediaData) = False Then
                GridMedia.Children.Add(MediaData)
            End If
        End Sub

        Sub Done()

            GridMedia.Children.Remove(MediaData)
        End Sub
    End Class

    Sub DataTableAdd()

        ' maakt een tabel

        ' W(0) = "table"
        ' W(1) = tabel naam

        If TableSet IsNot Nothing Then
            TI += 1
            CI = 0
            RI = -1
            TableSet.Tables.Add(W(1))
            TableSet.Tables(W(1)).Columns.Add(New DataColumn("index"))
            TableSet.Tables(W(1)).Columns(0).DataType = Type.GetType("System.Int16")

            DataTable(TI) = New DataView
            DataTable(TI).Table = TableSet.Tables(TI)

            TableList.Items.Add(W(1))
            TableList.SelectedIndex = TableList.Items.Count - 1
        Else
            MsgBox("Geen database")
            ParsSkip = True
        End If
    End Sub

    Sub DataTableColumn()

        ' voegt een kolom toe aan een tabel

        ' W(0) = "column"
        ' W(1) = naam
        ' W(2) = [type]
        ' W(3) = [lengte]
        ' W(4) = ["unique"]

        CI += 1

        TableSet.Tables(TI).Columns.Add(New DataColumn(W(1)))
        If W(2) <> "" Then
            Select Case LCase(W(2))
                Case "short"
                    TableSet.Tables(TI).Columns(CI).DataType = Type.GetType("System.Int16")
                Case "single"
                    TableSet.Tables(TI).Columns(CI).DataType = Type.GetType("System.Single")
                Case Else
                    TableSet.Tables(TI).Columns(CI).DataType = Type.GetType("System.String")
            End Select
        End If
        If W(3) <> "" Then
            TableSet.Tables(TI).Columns(CI).MaxLength = V(3)
        End If
        If W(4) = "uniek" Then
            TableSet.Tables(TI).Columns(CI).Unique = True
        End If
    End Sub

    Sub DataTableData()

        ' voegt gegevens uit data.txt aan huidige tabel toe

        ' W(0) = "data"
        ' W(1) = pad naam
        ' W(2) = tabel naam met vervolg pad
        ' W(3) = kolom naam met vervolg pad
        ' W(4) = bestand naam

        Dim P As String = GetFileAddress(W(1))
        Dim T As String = W(2)
        Dim C As String = W(3)
        Dim D As String = W(4)
        Dim F As String
        Dim S As String()
        Dim I As Short
        Dim N As Short
        Dim R As DataRow

        For Each Row In TableSet.Tables(T).Rows ' tabel met paden
            F = P & Row(C) & D ' adres van data bestand
            If File.Exists(F) Then ' als het data bestand bestaat
                S = System.IO.File.ReadAllLines(F) ' lees alle regels
                N = UBound(S) ' bepaal aantal regels
                For I = 0 To N ' voor alle regels
                    TextParsLine(S(I), ",") ' bepaal de woorden van een regel
                    If W(0) <> "" Then
                        If W(0) = V(0).ToString Then ' als het 1e woord een waarde is
                            RI = 1
                            For Each Row2 In TableSet.Tables(TI).Rows
                                If Row2(2) = W(1) Then
                                    GoTo gevonden
                                End If
                                RI += 1
                            Next
                            R = TableSet.Tables(TI).NewRow()
                            R(0) = RI
                            R(1) = 0
                            R(2) = W(1)
                            TableSet.Tables(TI).Rows.Add(R)
                            RI = TableSet.Tables(TI).Rows.Count
gevonden:
                            TableSet.Tables(TI).Rows(RI - 1)(1) += V(0)
                        Else
                            RI += 1 ' rij index van nieuwe tabel
                            R = TableSet.Tables(TI).NewRow ' nieuwe rij in nieuwe tabel
                            R(0) = RI  ' index
                            R(1) = Row(1)
                            R(2) = Row(2)
                            For J = 3 To CI
                                R(J) = W(J - 2)
                            Next
                            TableSet.Tables(TI).Rows.Add(R)
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    Sub DataTableDirs()

        ' voegt paden aan de laatst gemaakte tabel toe

        ' W(0) = "dirs"
        ' W(1) = pad
        ' W(2) = [tabel waarin vervolg paden staan]
        ' W(3) = [kolom waarin vervolg pad staat]

        Dim F As String = W(1)
        Dim T As String = W(2)
        Dim C As String = W(3)
        Dim I As Short
        Dim N As DataRow
        Dim Z As String

        RI = 0
        If T <> "" Then ' als een tabel met vervolg paden is opgegeven
            For Each Row In TableSet.Tables(T).Rows ' rij uit bron tabel
                Z = Row(C) ' zoek kolom in bron tabel
                For Each D In Directory.GetDirectories(F & Z) ' zoek mappen
                    'If Not File.Exists(D & "\folder.jpg") Then MsgBox(D & " heeft geen folder")
                    RI += 1
                    N = TableSet.Tables(TI).NewRow
                    N(0) = RI ' index
                    N(1) = Mid(D, Len(F) + 1) & "\" ' gevonden pad
                    If W(4) <> "" Then N(W(4)) = Mid(D, InStrRev(D, "\") + 1) ' gevonden map
                    I = 5
                    While W(I) <> ""
                        N(W(I)) = Row(W(I))
                        I += 1
                    End While
                    TableSet.Tables(TI).Rows.Add(N)
                Next
            Next
        Else ' zonder vervolgpad
            For Each D In Directory.GetDirectories(F)
                W(1) = Mid(D, Len(F) + 1) & "\"
                W(2) = Mid(D, InStrRev(D, "\") + 1)
                For I = 3 To CI
                    W(I) = ""
                Next
                DataTableRows()
            Next
        End If
    End Sub

    Sub DataTableFiles()

        ' voegt bestand namen aan een tabel toe

        ' W(0) = "files"
        ' W(1) = filepath
        ' W(2) = fileType
        ' W(3) = [table]
        ' W(4) = [column]

        Dim R As String = FileRoot.SelectedItem
        Dim P As String = R & GetFilePath(W(1))
        Dim E As String = W(2)
        Dim T As String = W(3)
        Dim C As String = W(4)
        Dim N As String

        RI = 0
        If T <> "" Then
            For Each Row In TableSet.Tables(T).Rows
                N = Row(C)
                For Each FoundFile In Directory.GetFiles(P & N, "*" & E)
                    W(1) = Mid(FoundFile, Len(P) + 1)
                    W(1) = Left(W(1), Len(W(1)) - 4)
                    For I = 2 To CI - 1
                        W(I) = Row(I)
                    Next
                    DataTableRows()
                Next
            Next
        Else
            For Each FoundFile In Directory.GetDirectories(P, "*." & E)
                W(1) = Mid(FoundFile, InStrRev(FoundFile, "\") + 1)
                V(1) = Val(W(1))
                DataTableRows()
            Next
        End If
    End Sub

    Sub DataTableIf()

        ' W(0) = "if"
        ' W(1) = bron tabel 
        ' W(2) = zoek kolom
        ' W(3) = zoek waarde
        ' W(4) = gevonden 

        Dim I As Short
        Dim N As Short
        Dim R As DataRow

        For Each Row In TableSet.Tables(W(1)).Rows
            If InStr(Row(W(2)), W(3)) > 0 Then
                I += 1
                R = TableSet.Tables(TI).NewRow
                R("Index") = I
                N = 4
                While W(N) <> ""
                    R(W(N)) = Row(W(N))
                    N += 1
                End While
                TableSet.Tables(TI).Rows.Add(R)
            End If
        Next
    End Sub

    Sub DataTableInfo()

        ' leest gegevens uit een info bestand

        ' W(0) = "info"
        ' W(1) = pad
        ' W(2) = tabel
        ' W(3) = kolom met het pad naar de info bestanden
        ' W(4)..W(20) = namen van de te lezen kolommen

        Dim P As String = FileRoot.SelectedItem & Left(W(1), 1) & "\" & W(1) ' pad
        Dim T As String = W(2) ' tabel
        Dim C As String = W(3) ' kolom
        Dim F As String
        Dim S As String()
        Dim N As Short
        Dim M As Short
        Dim K As Short

        M = 3
        Do
            M += 1
        Loop Until W(M + 1) = ""

        For Each Row In TableSet.Tables(T).Rows
            F = P & Row(C) & "info.txt"
            If File.Exists(F) Then
                S = File.ReadAllLines(F)
                N = UBound(S)
                For I = 0 To N
                    For J = 3 To M
                        K = InStr(S(I), W(J) & ",")
                        If K > 0 Then
                            Row(W(J)) = Trim(Mid(S(I), K + Len(W(J)) + 1))
                        End If
                    Next
                Next
            End If
        Next
    End Sub

    Sub DataTableRows()

        ' W(0) = "row"
        ' W(n) = data

        Dim I As Short
        Dim Rows As DataRow

        RI += 1
        Try
            CI = TableSet.Tables(TI).Columns.Count - 1
            Rows = TableSet.Tables(TI).NewRow
            Rows(0) = RI ' index
            For I = 1 To CI
                If TableSet.Tables(TI).Columns(I).DataType = Type.GetType("System.String") Then
                    Rows(I) = W(I)
                Else
                    Rows(I) = V(I)
                End If
            Next
            TableSet.Tables(TI).Rows.Add(Rows)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub DataTableSort()

        ' W(0) = "sort"
        ' W(1) = tabel
        ' W(2) = kolom
        ' W(3) = richting

        Dim I As Short

        I = TableIndex(W(1))

        ST(I) = W(2) & " " & W(3)
        DataTable(I).Sort = ST(I)
    End Sub

    Function TableIndex(S As String) As Short

        Dim I As Short
        Dim M As Short

        TableIndex = -1
        M = TableSet.Tables.Count - 1
        For I = 0 To M
            If TableSet.Tables(I).TableName = S Then
                TableIndex = I
            End If
        Next
    End Function
    Sub TableList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TableList.SelectionChanged

        ' als een tabel wordt geselecteerd

        Dim N As Short

        N = TableIndex(TableList.SelectedItem)
        If N <> -1 Then
            MediaData.ItemsSource = DataTable(N)
        End If
    End Sub

    Public TableSet As DataSet
    Public DataTable(9) As DataView
    Public ST(9) As String ' sorteer tabel
    Public TI As Short ' tabel index
    Public CI As Short ' kolom index
    Public RI As Long  ' rij index
    Public WithEvents TableList As New ListBox
End Module