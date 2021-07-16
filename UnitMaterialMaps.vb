Imports HelixToolkit.Wpf.SharpDX
Imports HelixToolkit.Wpf.SharpDX.Model
Imports HelixToolkit.Wpf.SharpDX.Model.Scene
Imports SharpDX
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Module UnitMaterialMaps

    ' van UnitGrid
    ' van UnitMaterial

    Public GridMaps As New ScrollViewer

    Function GridMapsInit() As UIElement

        ' start het tonen van een textuur bestand

        MapsGrid.Children.Clear()
        MapsGrid.Background = New SolidColorBrush(Colors.Transparent)
        MapsGrid.HorizontalAlignment = Windows.HorizontalAlignment.Left
        MapsGrid.VerticalAlignment = Windows.VerticalAlignment.Top

        GridMaps.Content = MapsGrid
        GridMaps.HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden
        GridMaps.VerticalScrollBarVisibility = ScrollBarVisibility.Hidden

        GridMapsInit = GridMaps
    End Function

    Sub GridMapsDone()

        MapsGrid.Children.Clear()
        MapsGrid.Background = New SolidColorBrush(Colors.Transparent)
        MapsGrid.Height = 380
        MapsGrid.Width = 380
    End Sub

    Sub GridMapsShow()

        ' bewerkt uvmap

        ' W(0) = "uvmap"
        ' W(1) = [bestandnaam om te openen]
        ' W(2) = [paint] 
        ' W(3) = [kleur]
        ' W(4) = [toon masker]
        ' W(5) = [bestandnaam om te bewaren]

        Dim B As New BitmapImage
        Dim F As String = ""
        Dim M As PhongMaterialCore
        Dim N As MeshNode
        Dim G As HelixToolkit.Wpf.SharpDX.MeshGeometry3D
        Dim I As Long
        Dim J As Long
        Dim S As String
        Dim T As String

        MapsLog = False
        If W(1) <> "" Then
            S = GetFileAddress(W(1))
            If File.Exists(S) Then
                MapsFile = S
                B.BeginInit()
                B.CacheOption = BitmapCacheOption.OnLoad
                B.UriSource = New Uri(S)
                MediaPict.Source = B
                B.EndInit()
                MapsGrid.Children.Clear()
                MapsGrid.Background = New ImageBrush(MapsPict.Source)
            Else
                MapsGrid.Children.Clear()
                MapsGrid.Background = New SolidColorBrush(Colors.Transparent)
                MsgBox(S & " niet gevonden")
            End If
        End If
        For Each Node In SceneModel.SceneNode.Traverse(True)
            T = Node.GetType.Name
            If InStr(T, "MeshNode") > 0 Then ' als de node een meshnode is
                N = Node
                If W(2) <> "" Then
                    Paint = True
                    If W(3) <> "" Then
                        MapsColor = PartMaterial.Color(W(3)).ToColor
                    Else
                        MapsColor = Colors.Red
                    End If
                End If
                M = N.Material

                MapsGrid.Children.Clear()
                MapsGrid.Background = New SolidColorBrush(Colors.Transparent)

                If W(4) <> "" Then
                    G = N.Geometry
                    J = G.TriangleIndices.Count
                    For I = 0 To J - 1 Step 3
                        MapsGrid.Children.Add(DrawTriangle(
                                (G.TextureCoordinates(G.TriangleIndices(I)) * MapsHeight).ToPoint,
                                (G.TextureCoordinates(G.TriangleIndices(I + 1)) * MapsHeight).ToPoint,
                                (G.TextureCoordinates(G.TriangleIndices(I + 2)) * MapsHeight).ToPoint,
                            Colors.Black, False))
                    Next
                End If
                If W(5) = "" Then
                    MapsFile = F
                Else
                    MapsFile = GetFileAddress(W(5))
                End If

                If W(6) = "" Then
                    MapsHeight = 1024 ' MapsPict.ActualHeight
                Else
                    MapsHeight = V(6)
                End If
                MapsGrid.Height = MapsHeight
                MapsGrid.Width = MapsHeight
                If M.DiffuseMap Is Nothing Then
                    MapsGrid.Background = New SolidColorBrush(M.DiffuseColor.ToColor)
                Else
                    MapsPict.Source = Nothing

                    Dim Decoder As New PngBitmapDecoder(M.DiffuseMap.CompressedStream, BitmapCreateOptions.None, BitmapCacheOption.Default)
                    MapsPict.Source = Decoder.Frames(0)
                    MapsGrid.Background = New ImageBrush(MapsPict.Source)
                End If

                If W(7) <> "" Then MapsLog = True
                Exit Sub
            End If
        Next
    End Sub

    Sub GridMapsPaint(C As Vector2)

        ' kleurt een driehoek van een model als op het muiswiel geklikt wordt

        If Paint Then
            Paint = False

            Dim H As Integer = MapsHeight
            Dim O As HelixToolkit.Wpf.SharpDX.MeshGeometryModel3D
            Dim Q As MeshNode
            Dim M As PhongMaterialCore

            Dim U(3) As Vector2
            Dim P(3) As Vector3
            Dim N(3) As Vector3
            Dim S As String
            Dim T As Tuple(Of Integer, Integer, Integer)

            T = MediaScene.FindHits(C)(0).TriangleIndices ' geeft de 3 indexen van de geselecteerde driehoek!
            PartGeo = MediaScene.FindHits(C)(0).Geometry ' copieerd de gehele geometry van het model
            P(1) = PartGeo.Positions(T.Item1)
            P(2) = PartGeo.Positions(T.Item2)
            P(3) = PartGeo.Positions(T.Item3)
            U(1) = PartGeo.TextureCoordinates(T.Item1)
            U(2) = PartGeo.TextureCoordinates(T.Item2)
            U(3) = PartGeo.TextureCoordinates(T.Item3)
            N(1) = PartGeo.Normals(T.Item1)
            N(2) = PartGeo.Normals(T.Item2)
            N(3) = PartGeo.Normals(T.Item3)
            If MapsLog Then ' bewaart driehoek
                GridText.AppendText("point, 1 " & SF(PartGeo.Positions(T.Item1).X) & SF(PartGeo.Positions(T.Item1).Y) & SF(PartGeo.Positions(T.Item1).Z) & vbCrLf)
                GridText.AppendText("point, 2 " & SF(PartGeo.Positions(T.Item2).X) & SF(PartGeo.Positions(T.Item2).Y) & SF(PartGeo.Positions(T.Item2).Z) & vbCrLf)
                GridText.AppendText("point, 3 " & SF(PartGeo.Positions(T.Item3).X) & SF(PartGeo.Positions(T.Item3).Y) & SF(PartGeo.Positions(T.Item3).Z) & vbCrLf)
                GridText.AppendText("vt, 1 " & SF(U(1).X) & SF(U(1).Y) & vbCrLf)
                GridText.AppendText("vt, 2 " & SF(U(2).X) & SF(U(2).Y) & vbCrLf)
                GridText.AppendText("vt, 3 " & SF(U(3).X) & SF(U(3).Y) & vbCrLf)
                GridText.AppendText("vn, 1 " & SF(N(1).X) & SF(N(1).Y) & SF(N(1).Z) & vbCrLf)
                GridText.AppendText("vn, 2 " & SF(N(2).X) & SF(N(2).Y) & SF(N(2).Z) & vbCrLf)
                GridText.AppendText("vn, 3 " & SF(N(3).X) & SF(N(3).Y) & SF(N(3).Z) & vbCrLf)
                GridText.AppendText("face, 1/1/1, 2/2/2, 3/3/3" & vbCrLf)
            End If
            U(1) *= H
            U(2) *= H
            U(3) *= H
            MapsGrid.Children.Add(DrawTriangle(U(1).ToPoint, U(2).ToPoint, U(3).ToPoint, MapsColor, True))

            S = MediaScene.FindHits(C)(0).ModelHit.GetType.ToString
            If InStr(S, "MeshNode") > 0 Then
                Q = MediaScene.FindHits(C)(0).ModelHit
                M = Q.Material
            Else
                O = MediaScene.FindHits(C)(0).ModelHit
                M = O.Material
            End If
            Wait(0.2)

            Dim Encoder As New PngBitmapEncoder()
            Dim MapsBitmap As New RenderTargetBitmap(MapsHeight, MapsHeight, 96, 96, PixelFormats.Default)

            MapsBitmap.Render(MapsGrid)
            Encoder.Frames.Add(BitmapFrame.Create(MapsBitmap))
            Encoder.Save(MapsStream)
            M.DiffuseMap = MapsStream
            M.DiffuseColor = New Color4(1, 1, 1, 1)
            Paint = True
        End If
    End Sub

    Sub GridMapsSave()

        ' bewaart een materiaal map

        GridMaps.ScrollToHorizontalOffset(0)
        GridMaps.ScrollToVerticalOffset(0)

        Dim W As Integer = MapsHeight
        Dim H As Integer = MapsHeight

        Dim Bitmap As New RenderTargetBitmap(W, H, 96, 96, PixelFormats.Default)
        Dim Encoder As New PngBitmapEncoder()
        Dim Stream As New FileStream(MapsFile, FileMode.Create)

        Bitmap.Render(MapsGrid)
        Encoder.Frames.Add(BitmapFrame.Create(Bitmap))
        Using Stream
            Encoder.Save(Stream)
        End Using
    End Sub

    Function LoadMapsToMemory(F As String) As MemoryStream

        If File.Exists(F) Then
            Using File = New FileStream(F, FileMode.Open)

                Dim M = New MemoryStream()

                File.CopyTo(M)
                Return M
            End Using
        Else
            MsgBox(F & " niet gevonden")
            ParsSkip = True

            Return Nothing
        End If
    End Function

    Public MapsGrid As New Grid
    Public MapsPict As New Image
    Public MapsStream As New MemoryStream
    Public MapsFile As String
    Public MapsColor As Windows.Media.Color
    Public MapsHeight As Short
    Public MapsLog As Boolean
End Module
