Imports HelixToolkit.Wpf.SharpDX
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes

Module UnitMediaDraw

    Public MediaDraw As New ClassMediaDraw

    Class ClassMediaDraw
        Inherits Grid

        Sub Init()

            ' start het tekenveld

            MediaDraw.Background = New SolidColorBrush(Colors.Transparent) ' anders blijven foto's van het grid zwart
        End Sub

        Sub Runs()

            ' start een tekening

            ' W(0) = "draw"

            MediaPict.Done()
            MediaScene.Done()
            MediaDraw.Children.Clear()
            If GridMedia.Children.Contains(MediaDraw) = False Then GridMedia.Children.Add(MediaDraw)
        End Sub

        Sub Done()

            ' wist tekening

            MediaDraw.Children.Clear()
        End Sub

        Sub Angle()

            ' hoeken

            ' W(0) = "angle"
            ' V(1) = X
            ' V(2) = Y
            ' V(3) = Z

            AX(0) = V(1)
            AY(0) = V(2)
            AZ(0) = V(3)
        End Sub

        Sub Back()

            ' achtergrond kleur

            ' W(0) = "back"
            ' W(1) = color

            If InStr(W(1), ".") = 0 Then

                Dim C As Color

                C = PartMaterial.Color(W(1)).ToColor
                Background = New SolidColorBrush(C)
                MediaScene.BackgroundColor = C
            Else

                Dim M = New MemoryStream(LoadMapsToMemory(GetFileAddress(W(1))).ToArray)
                Dim B As New BitmapImage

                B.BeginInit()
                B.StreamSource = M
                B.CacheOption = BitmapCacheOption.OnLoad
                B.EndInit()

                Background = New ImageBrush(B)
                MediaScene.Background = New ImageBrush(B)
            End If
        End Sub

        Sub Circle()

            ' tekent een cirkel

            ' W(0) = "circle"
            ' V(1) = straal
            ' V(2) = aantal segmenten
            ' V(3) = [X middelpunt]
            ' V(4) = [Y middelpunt]
            ' V(5) = [beginhoek]
            ' V(6) = [eindhoek]
            ' W(7) = [kleur]
            ' W(8) = [vullen]

            Dim P As Point

            V(1) *= SC * GS
            If W(5) = "" Then
                V(5) = 0
                V(6) = 360
            End If

            For I = FI To 0 Step -1
                V(5) += AZ(I)
                V(6) += AZ(I)
            Next

            FI += 1
            OX(FI) = V(3) * SC * GS
            OY(FI) = V(4) * SC * GS

            For I = FI To 0 Step -1
                P.X += OX(I)
                P.Y += OY(I)
                Dim E As New RotateTransform(AZ(I), OX(I), OY(I))
                ' verdraai middelpunt om draaipunt
                P = E.Transform(P)
            Next

            If W(7) = "" Then W(7) = "black"

            Children.Add(DrawSegment(P, V(1), V(5), V(6), V(2), PartMaterial.Color(W(7)).ToColor, W(8) <> ""))
            FI -= 1
        End Sub
    End Class

    Sub MediaDrawClear()

        ' wist een tekening

        ' W(0) = "clear"
        ' V(1) = [wachttijd]

        If W(1) <> "" Then Wait(V(1))

        MediaDraw.Children.Clear()
    End Sub

    Sub MediaDrawGrid()

        ' plaatst een raster

        ' W(0) = "grid"
        ' V(1) = raster afstand
        ' W(2) = ["show"]
        ' W(3) = ["use"]

        GS = V(1) * SC

        If W(2) <> "" Then
            For I = OY(0) - GS To 720 Step GS ' horizontaal
                MediaDraw.Children.Add(DrawLine(0, I, 1280, I, Colors.LightGray, 1))
            Next
            For I = OX(0) - GS To 1280 Step GS ' verticaal
                MediaDraw.Children.Add(DrawLine(I, 0, I, 720, Colors.LightGray, 1))
            Next
        End If
        If W(3) = "" Then GS = 1
    End Sub

    Sub MediaDrawLine()

        ' tekent een lijn

        ' W(0) = "lijn"
        ' V(1) = lengte langs x-as
        ' V(2) = x begin punt
        ' V(3) = y begin punt
        ' V(4) = [Z hoek]
        ' W(5) = [kleur]
        ' w(6) = [dikte]

        For I = 1 To 3 ' schaal de waardes
            V(I) *= SC * GS
        Next

        Dim P(2) As Point ' lijn bestaat uit 2 punten

        P(1).X = V(2) ' beginpunt
        P(1).Y = V(3)

        P(2).X = P(1).X + V(1) ' eindpunt
        P(2).Y = P(1).Y

        Dim D As New RotateTransform(V(4), P(1).X, P(1).Y) ' verdraai eindpunt met opgegeven hoek om beginpunt

        P(2) = D.Transform(P(2))

        For I = FI To 0 Step -1
            P(1).X += OX(I)
            P(1).Y += OY(I)
            P(2).X += OX(I)
            P(2).Y += OY(I)

            Dim E As New RotateTransform(AZ(I), OX(I), OY(I)) ' verdraai begin- en eindpunt om draaipunt

            P(1) = E.Transform(P(1))
            P(2) = E.Transform(P(2))
        Next
        If W(5) = "" Then W(5) = "black"
        If W(6) = "" Then V(6) = 1 ' dikte
        MediaDraw.Children.Add(DrawLine(P(1).X, P(1).Y, P(2).X, P(2).Y, PartMaterial.Color(W(5)).ToColor, V(6))) ' voeg lijn toe aan tekening
    End Sub

    Sub MediaDrawOffset()

        ' plaatst een onderdeel of tekening op een bepaalde afstand 

        ' W(0) = "offset"
        ' V(1) = X
        ' V(2) = Y
        ' V(3) = Z
        ' W(4) = [show]

        OX(0) = V(1) * SC * GS
        OY(0) = V(2) * SC * GS
        OZ(0) = V(3) * SC * GS

        If W(4) <> "" Then
            MediaDraw.Children.Add(DrawLine(0, OY(FI), 1280, OY(FI), Colors.Gray, 1)) ' X axis
            MediaDraw.Children.Add(DrawLine(OX(FI), 0, OX(FI), 720, Colors.Gray, 1)) ' Y axis
        End If
    End Sub

    Sub MediaDrawPicture()

        ' plaats een foto

        ' W(0) = "pict"
        ' W(1) = bestand
        ' V(2) = [wacht]

        W(1) = GetFileAddress(W(1))
        MediaPict.Load(W(1))
        Wait(V(2))
    End Sub

    Sub MediaDrawPlay()

        ' plaatst een film of speelt muziek

        ' W(0) = "play"
        ' W(1) = file
        ' V(2) = [start]
        ' V(3) = [stop]

        'If Not MediaDraw.Children.Contains(MediaPlay) Then MediaDraw.Children.Add(MediaPlay)
        W(1) = GetFileAddress(W(1))
        MediaPlay.Load(W(1))
        PlayFrom = V(2)
        PlayTill = V(3)
    End Sub

    Sub MediaDrawRectangle()

        ' tekent een rechthoek

        ' W(0) = "rectangle"
        ' V(1) = width
        ' V(2) = height
        ' V(3) = X left
        ' V(4) = Y top
        ' V(5) = hoek
        ' W(6) = [kleur]
        ' W(7) = [vullen]

        Dim P(4) As Point

        For I = 1 To 4
            V(I) *= SC * GS
        Next

        FI += 1
        OX(FI) = V(3)
        OY(FI) = V(4)
        AZ(FI) = V(5)
        P(1) = New Point(0, 0) ' links boven
        P(2) = New Point(V(1), 0) ' rechts boven
        P(3) = New Point(V(1), V(2)) ' rechts onder
        P(4) = New Point(0, V(2)) ' links onder

        For I = FI To 0 Step -1
            P(1).X += OX(I)
            P(1).Y += OY(I)
            P(2).X += OX(I)
            P(2).Y += OY(I)
            P(3).X += OX(I)
            P(3).Y += OY(I)
            P(4).X += OX(I)
            P(4).Y += OY(I)

            Dim E As New RotateTransform(AZ(I), OX(I), OY(I)) ' verdraai begin- en eindpunt om draaipunt

            P(1) = E.Transform(P(1))
            P(2) = E.Transform(P(2))
            P(3) = E.Transform(P(3))
            P(4) = E.Transform(P(4))
        Next

        FI -= 1

        If W(6) = "" Then W(6) = "black"
        MediaDraw.Children.Add(DrawRect(P(1), P(2), P(3), P(4), PartMaterial.Color(W(6)).ToColor, W(7) <> ""))
    End Sub

    Sub MediaDrawScale()

        ' schaal

        ' W(0) = "scale"
        ' V(1) = percentage

        If W(1) <> "" Then
            SC = V(1) / 100
        Else
            SC = 1
        End If
        SX(0) = SC
        SY(0) = SC
        SZ(0) = SC
    End Sub

    Sub MediaDrawText()

        ' plaatst een tekst

        ' W(0) = "text"
        ' W(1) = text
        ' V(2) = X
        ' V(3) = Y
        ' V(4) = [fontsize)
        ' W(5) = [fontname]
        ' V(6) = [wait ms]

        V(2) *= SC * GS
        V(3) *= SC * GS
        If W(4) = "" Then ' als geen lettergroote is opgegeven
            V(4) = 56
        End If
        V(4) *= SC
        V(3) -= V(4) / 2

        Dim P As New Point(V(2), V(3))

        For I = FI To 0 Step -1
            P.X += OX(I)
            P.Y += OY(I)

            Dim E As New RotateTransform(AZ(I), OX(I), OY(I)) ' verdraai begin- en eindpunt om draaipunt

            P = E.Transform(P)
        Next

        MediaDraw.Children.Add(DrawText(W(1), P.X, P.Y, V(4), W(5), Colors.Black))
        If W(6) <> "" Then Wait(V(6))
    End Sub

    Sub MediaDrawTriangle()

        ' tekent een driehoek
        ' W(0) = "triangle"
        ' V(1) = breedte
        ' V(2) = hoogte
        ' V(3) = top links
        ' V(4) = [x]
        ' V(5) = [y]
        ' V(6) = [hoek]
        ' W(7) = [kleur]
        ' W(8) = [vullen]

        Dim P(3) As Point

        For I = 1 To 5
            V(I) *= SC * GS
        Next

        FI += 1

        OX(FI) = V(4)
        OY(FI) = V(5)
        AZ(FI) = V(6)
        P(1) = New Point(-V(1) / 2, 0) ' basis links
        P(2) = New Point(V(1) / 2, 0) ' basis rechts
        P(3) = New Point(-V(1) / 2 + V(3), -V(2)) ' top

        For I = FI To 0 Step -1
            P(1).X += OX(I)
            P(1).Y += OY(I)
            P(2).X += OX(I)
            P(2).Y += OY(I)
            P(3).X += OX(I)
            P(3).Y += OY(I)

            Dim E As New RotateTransform(AZ(I), OX(I), OY(I)) ' verdraai begin- en eindpunt om draaipunt

            P(1) = E.Transform(P(1))
            P(2) = E.Transform(P(2))
            P(3) = E.Transform(P(3))
        Next

        FI -= 1

        If W(7) = "" Then W(7) = "black"
        MediaDraw.Children.Add(DrawTriangle(P(1), P(2), P(3), PartMaterial.Color(W(7)).ToColor, W(8) <> ""))
    End Sub


    Function DrawLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, C As Color, T As Short) As Line

        ' tekent een lijn

        Dim L As New Windows.Shapes.Line With {
            .X1 = X1,
            .Y1 = Y1,
            .X2 = X2,
            .Y2 = Y2,
            .Stroke = New Windows.Media.SolidColorBrush(C),
            .StrokeThickness = T
        }
        Return L
    End Function

    Function DrawRect(P1 As Point, P2 As Point, P3 As Point, P4 As Point, C As Windows.Media.Color, F As Boolean) As Windows.Shapes.Polygon

        ' tekent rechthoek

        Dim P As New Windows.Shapes.Polygon

        P.Points.Add(P1) ' links boven
        P.Points.Add(P2) ' links onder
        P.Points.Add(P3) ' rechts onder
        P.Points.Add(P4) ' rechts boven
        If F Then
            P.Fill = New SolidColorBrush(C)
        Else
            P.StrokeThickness = 1
            P.Stroke = New SolidColorBrush(C)
        End If
        Return P
    End Function

    Function DrawSegment(M As Point, R As Single, HB As Single, HE As Single, N As Short, C As Windows.Media.Color, F As Boolean) As Polyline

        ' tekent segment

        Dim I As Short
        Dim J As Short
        Dim L As New Polyline

        J = Math.Abs(HE - HB) * N / 360 + 1 ' aantal punten
        If HE < HB Then N = -N
        For I = 1 To J
            L.Points.Add(New Point(M.X + DXC(R, HB + 360 / N * (I - 1)), M.Y + DYS(R, HB + 360 / N * (I - 1))))
        Next
        L.Stroke = New SolidColorBrush(C)
        L.StrokeThickness = 1
        If F Then L.Fill = New SolidColorBrush(C)
        Return L
    End Function

    Function DrawText(T As String, X As Single, Y As Single, S As Short, N As String, C As Windows.Media.Color) As TextBlock

        ' plaatst een tekst

        Dim B As New TextBlock

        Y -= S / 14
        If N <> "" Then ' font naam
            B.FontFamily = New FontFamily(N)
        Else
            B.FontFamily = New FontFamily("Consolas")
        End If
        B.FontSize = S
        B.Text = T
        B.Foreground = New SolidColorBrush(C)
        B.Margin = New Thickness(X, Y, 0, 0)
        Return B
    End Function

    Function DrawTriangle(P1 As Point, P2 As Point, P3 As Point, C As Windows.Media.Color, F As Boolean) As Shapes.Polygon

        ' tekent een driehoek

        Dim P As New Shapes.Polygon

        P.Points.Add(P1)
        P.Points.Add(P2)
        P.Points.Add(P3)
        If F Then
            P.Fill = New SolidColorBrush(C)
        Else
            P.Stroke = New SolidColorBrush(C)
            P.StrokeThickness = 1
        End If
        Return P
    End Function

    Function Deg(R As Single) As Single

        ' maakt graden van radialen

        Deg = R * 180 / Math.PI
    End Function

    Function DXC(Hypothenusa As Double, Angle As Double) As Double

        ' berekent liggende rechthoekzijde of x van cosinus

        DXC = Hypothenusa * Math.Cos(Rad(Angle))
    End Function

    Function DYS(Hypothenusa As Double, Angle As Double) As Double

        ' berekent staande rechthoekzijde of y van sinus

        DYS = Hypothenusa * Math.Sin(Rad(Angle))
    End Function

    Function Rad(D As Double) As Double

        ' maakt radialen van graden

        Rad = D * Math.PI / 180
    End Function
End Module