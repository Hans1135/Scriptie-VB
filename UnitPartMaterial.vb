Imports HelixToolkit.Wpf.SharpDX
Imports HelixToolkit.Wpf.SharpDX.Model
Imports HelixToolkit.Wpf.SharpDX.Model.Scene
Imports SharpDX
Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Media

Module UnitPartMaterial

    Public PartMaterial As New ClassPartMaterial

    Class ClassPartMaterial
        Inherits PhongMaterialCore

        Sub Init()

            ' start materiaal


            MC = 0
            DiffuseColor = New Color4(0, 0, 1, 1) ' blauw
        End Sub

        Function Color(S As String) As Color4

            ' zoekt een kleur

            Dim C As Color4

            Select Case LCase(S)
                Case "black"
                    C = New Color4(&H1B / 255, &H2A / 255, &H34 / 255, 1) ' niet helemaal zwart #1B2A34
                Case "blue"
                    C = New Color4(&H1E / 255, &H5A / 255, &HA8 / 255, 1)
                Case "blue-l" ' licht blauw
                    C = Colors.LightBlue.ToColor4
                Case "blue-m" ' medium blauw
                    C = New Color4(&H73 / 255, &H96 / 255, &HC8 / 255, 1) ' 73 96 C8
                Case "blue-mt" ' medium transparant blauw
                    C = New Color4(&H55 / 255, &H9A / 255, &HB7 / 255, 0.5) ' 55 9A B7
                Case "brown"
                    C = New Color4(&H7C / 255, &H5C / 255, &H45 / 255, 1)
                Case "brown-l" ' licht bruin
                    C = New Color4(&H7C / 255, &H5C / 255, &H45 / 255, 1)
                Case "brown-r" ' rood bruin
                    C = New Color4(&H5F / 255, &H31 / 255, &H9 / 255, 1)
                Case "brown-t" ' transparant bruin
                    C = New Color4(&H63 / 255, &H5F / 255, &H52 / 255, 0.5)
                Case "clear" ' transparant wit = helder
                    C = New Color4(&HFC / 255, &HFC / 255, &HFC / 255, 0.5)
                Case "grey" ' grijs, oud-grijs
                    C = New Color4(&H8A / 255, &H92 / 255, &H8D / 255, 1)
                Case "grey-b" ' blauw - grijs steen grijs
                    C = New Color4(&H96 / 255, &H96 / 255, &H96 / 255, 1)
                Case "grey-d" ' donker grijs
                    C = New Color4(&H54 / 255, &H59 / 255, &H55 / 255, 1)
                Case "grey-db" ' donker blauw grijs donker steen grijs
                    C = New Color4(&H64 / 255, &H64 / 255, &H64 / 255, 1)
                Case "grey-l" ' licht grijs
                    C = Colors.LightGray.ToColor4
                Case "green"
                    C = New Color4(&H28 / 255, &H7F / 255, &H46 / 255, 1)
                Case "green-l" ' licht groen
                    C = Colors.LightGreen.ToColor4
                Case "green-d" ' donker groen
                    C = New Color4(&H0 / 255, &H45 / 255, &H1A / 255, 1)
                Case "green-t" ' transparant groen
                    C = New Color4(&H23 / 255, &H78 / 255, &H41 / 255, 0.5) '# 23 78 41
                Case "gold"
                    C = PhongMaterials.Gold.DiffuseColor
                Case "orange-d"
                    C = New Color4(&H91 / 255, &H50 / 255, &H1C / 255, 1) '# 91 50 1C
                Case "pink"
                    C = New Color4(1.0F, 0.75F, 0.8F, 1)
                Case "red"
                    C = New Color4(&HB4 / 255, 0, 0, 1)
                Case "red-l" ' licht rood
                    C = Colors.LightPink.ToColor4
                Case "red-t" ' transparant rood
                    C = New Color4(&HC9 / 255, &H1A \ 255, &H9 \ 255, 0.5) ' # C9 1A 09
                Case "silver"
                    C = New Color4(&HCE / 255, &HCE / 255, &HCE / 255, 1)
                Case "tan"
                    C = New Color4(0.82F, 0.7F, 0.55F, 1)
                Case "tan-d" ' donker beige
                    C = New Color4(&H89 / 255, &H7D / 255, &H62 / 255, 1) ' # 89 7D 62
                Case "yellow"
                    C = New Color4(&HFA / 255, &HC8 / 255, &HA / 255, 1)
                Case "yellow-t"
                    C = New Color4(&HF5 / 255, &HCD / 255, &H2F / 255, 0.5)
                Case "white"
                    C = New Color4(&HF4 / 255, &HF4 / 255, &HF4 / 255, 1)
                Case "white-milk"
                    C = New Color4(&HEE / 255, &HEE / 255, &HEE / 255, 1)
                Case Else ' black
                    MsgBox("kleur " & S & " in " & ParsFile(FI) & " niet gevonden")
                    ParsSkip = True
                    C = New Color4(0, 0, 0, 1)
            End Select

            Return C
        End Function

        Sub Export()

            ' exporteer materialen

            ' W(0) = "materials"
            ' W(1) = bestandnaam
            ' W(2) = map type
            ' W(3) = [exporteer onderdeelnummers]

            Dim F As String
            Dim I As Short
            Dim J As Short
            Dim K As Short
            Dim M As MeshNode
            Dim N As String
            Dim O As String
            Dim P As PhongMaterialCore
            Dim D As DiffuseMaterialCore
            Dim S As String
            Dim T As String

            If W(2) = "" Then
                MsgBox("geen type opgegeven")
                ParsSkip = True
                Exit Sub
            End If
            W(1) = GetFileAddress(W(1))
            FileOpen(1, W(1), OpenMode.Output)
            I = 0
            For Each Node In SceneModel.SceneNode.Traverse(True)
                If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                    I += 1
                    M = Node
                    S = M.Material.GetType.ToString
                    If InStr(S, "Diffuse") > 0 Then
                        D = M.Material
                        If D.DiffuseMapFilePath Is Nothing Then
                            S = ""
                        Else
                            S = D.DiffuseMapFilePath.ToString
                            If W(4) = "" Then
                                J = InStrRev(S, "\")
                                K = InStrRev(S, "/")
                                If J < K Then J = K
                                F = Mid(S, J + 1)
                                T = Right(GetFileType(F), 3)
                                S = "d:\data\" & T & "\" & W(2) & "\" & F
                            End If
                        End If
                        PrintLine(1, "material, " & "d:\data\mtl\" & W(2) & "\" & M.Name & ".mtl, " & I & ", " & M.Name) ' 
                        T = ""
                        T &= "# " & D.Name & vbCrLf
                        T &= "Kd " & D.DiffuseColor.Red & " " & D.DiffuseColor.Green & " " & D.DiffuseColor.Blue & vbCrLf
                        If S <> "" Then T &= "map_Kd " & S & vbCrLf
                        S = FileRoot.SelectedItem & GetFilePath(W(1))
                        File.WriteAllText("d:\data\mtl\new\" & M.Name & ".mtl", T)
                    ElseIf InStr(S, "Phong") > 0 Then
                        P = M.Material
                        If P.NormalMapFilePath Is Nothing Then
                            N = ""
                        Else
                            N = P.NormalMapFilePath
                            If W(4) = "" Then
                                J = InStrRev(N, "\")
                                K = InStrRev(N, "/")
                                If J < K Then J = K
                                F = Mid(N, J + 1) ' filenaam
                                T = Right(GetFileType(F), 3) ' filetype
                                N = "d:\data\" & T & "\" & W(2) & "\" & F
                            End If
                        End If
                        If P.DiffuseMapFilePath Is Nothing Then
                            S = ""
                        Else
                            S = P.DiffuseMapFilePath.ToString
                            If W(4) = "" Then
                                J = InStrRev(S, "\")
                                K = InStrRev(S, "/")
                                If J < K Then J = K
                                F = Mid(S, J + 1) ' filenaam
                                T = Right(GetFileType(F), 3) ' filetype
                                S = "d:\data\" & T & "\" & W(2) & "\" & F
                            End If
                        End If
                        If W(3) = "" Then
                            O = ""
                        Else
                            O = ", " & I ' onderdeel nummer
                        End If
                        PrintLine(1, "material, " & "d:\data\mtl\" & W(2) & "\" & P.Name & ".mtl" & ", " & P.Name & O)
                        T = ""
                        T &= "# " & P.Name & vbCrLf
                        'T &= "Ka " & P.AmbientColor.Red & " " & P.AmbientColor.Green & " " & P.AmbientColor.Blue & vbCrLf
                        T &= "Kd " & P.DiffuseColor.Red & " " & P.DiffuseColor.Green & " " & P.DiffuseColor.Blue & vbCrLf
                        'T &= "Ke " & P.EmissiveColor.Red & " " & P.EmissiveColor.Green & " " & P.EmissiveColor.Blue & vbCrLf
                        'T &= "Kr " & P.ReflectiveColor.Red & " " & P.ReflectiveColor.Green & " " & P.ReflectiveColor.Blue & vbCrLf
                        'T &= "Ks " & P.SpecularColor.Red & " " & P.SpecularColor.Green & " " & P.SpecularColor.Blue & vbCrLf
                        'T &= "Ns " & "20" & vbCrLf ' & P.SpecularShininess & vbCrLf
                        If S <> "" Then T &= "map_kd " & S & vbCrLf
                        'If N <> "" Then T &= "map_kn " & N & vbCrLf
                        S = FileRoot.SelectedItem & GetFilePath(W(1))
                        File.WriteAllText("d:\data\mtl\new\" & P.Name & ".mtl", T)
                    End If
                End If
            Next
            FileClose(1)
        End Sub

        Sub Import_old()

            ' importeer een materiaal

            ' W(0) = "material"
            ' W(1) = kleur of file
            ' V(2) = [doorzichtigheid of materiaalnaam]
            ' V(3) = [spiegeling]

            Dim I As Short
            Dim K As Short
            Dim N As MeshNode
            Dim O As Short
            Dim S As String = ""
            Dim T As String

            MC = 0

            PartMaterial = New PhongMaterialCore
            PartMaterial.DiffuseColor = New Color4(0, 0, 1, 1) ' blauw
            PartMaterial.AmbientColor = New SharpDX.Color4(1, 1, 1, 1) ' wit
            PartMaterial.EmissiveColor = New SharpDX.Color4(0, 0, 0, 1) ' zwart
            PartMaterial.ReflectiveColor = New SharpDX.Color4(0, 0, 0, 1) ' zwart
            PartMaterial.SpecularColor = New SharpDX.Color4(ShininessSlider.Value, ShininessSlider.Value, ShininessSlider.Value, 1) ' wit
            PartMaterial.SpecularShininess = (1 - ShininessSlider.Value) * 255 ' 0..10000? hoe hoger hoe minder schittering
            PartMaterial.EnableAutoTangent = True
            PartMaterial.EnableFlatShading = False
            PartMaterial.EnableTessellation = True
            PartMaterial.RenderEnvironmentMap = True
            PartMaterial.RenderShadowMap = True
            PartMaterial.RenderDiffuseMap = True

            If InStr(W(1), ".") = 0 Then ' als het materiaal niet uit een bestand komt maar een kleur is
                PartMaterial.Name = W(1) ' naam van de kleur
                PartMaterial.DiffuseColor = PartMaterial.Color(W(1)) ' zoek de kleur
                If W(2) <> "" Then ' doorzichtig
                    PartMaterial.DiffuseColor = New Color4(PartMaterial.DiffuseColor.Red, PartMaterial.DiffuseColor.Green, PartMaterial.DiffuseColor.Blue, 1 - V(2))
                End If
                PartMaterial.AmbientColor = PartMaterial.DiffuseColor
                If W(3) <> "" Then ' glans
                    PartMaterial.SpecularColor = New SharpDX.Color4(V(3), V(3), V(3), 1)
                    PartMaterial.SpecularShininess = 255 - V(3) * 255
                End If
            Else ' als materiaal uit een bestand komt
                If GetFileType(W(1)) = ".txt" Then ' pbr materiaal
                    'L.AlbedoColor = New Color4(1, 1, 1, 1) ' weerkaatsingsvermogen
                    'L.AlbedoMapFilePath = "d:\data\png\mh\curly2.png"
                    'L.AlbedoMap = LoadMapsToMemory(L.AlbedoMapFilePath)
                    'L.AmbientOcculsionMapFilePath = "d:\data\png\mh\curly2.png" ' mate van verlichting
                    'L.AmbientOcculsionMap = LoadMapsToMemory(L.AmbientOcculsionMapFilePath)
                    'PBR = L
                    'FileAddress(FI) = GetFileAddress(W(1))
                    'TextParsRuns()
                    'PartMaterial(MC) = PBR
                ElseIf GetFileType(W(1)) = ".mtl" Then ' phong materiaal
                    S = W(2) ' naam
                    O = V(3)
                    MaterialTrans = 1
                    T = GetFileAddress(W(1))
                    FI += 1
                    ParsFile(FI) = T
                    TextParsRuns(ParsFile(FI))
                    FI -= 1
                    PartMaterial.Name = S
                Else ' als het een afbeelding bestand is
                    W(1) = GetFileAddress(W(1)) ' bestand
                    If File.Exists(W(1)) Then
                        S = W(2)
                        O = V(3)
                        PartMaterial.DiffuseMap = LoadMapsToMemory(W(1))
                        PartMaterial.DiffuseMapFilePath = W(1)
                        PartMaterial.DiffuseColor = New Color4(1, 1, 1, 1)
                        PartMaterial.AmbientColor = New Color4(AmbientSlider.Value, AmbientSlider.Value, AmbientSlider.Value, 1)
                        PartMaterial.Name = GetFileName(W(1))
                        UVCorr(0) = New Vector2(V(4), V(5))
                    Else
                        MsgBox("materiaal" & vbCrLf & W(1) & vbCrLf & "niet gevonden")
                        ParsSkip = True
                        Exit Sub
                    End If
                End If
                If S <> "" Then ' als materiaal een naam heeft
                    I = 0
                    K = 0
                    For Each Node In SceneModel.SceneNode.Traverse(False)
                        K += 1
                        If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                            I += 1
                            N = Node
                            If O > 0 Then ' als het materiaal een nummer heeft
                                If I = O Then
                                    N.Material = PartMaterial
                                    GoTo Verder
                                End If
                            ElseIf N.Material.Name = S Then
                                N.Material = PartMaterial
                                GoTo Verder
                            End If
                        End If
                    Next
                    MsgBox("onderdeel " & S & " voor " & O & " niet gevonden")
                    ParsSkip = True
                End If
            End If
Verder:
        End Sub

        Sub Import()

            ' importeer een materiaal

            ' W(0) = "material"
            ' W(1) = kleur of bestand
            ' V(2) = [doorzichtigheid of materiaalnaam]
            ' V(3) = [spiegeling of onderdeel]

            Dim I As Short
            Dim K As Short
            Dim N As MeshNode
            Dim O As Short
            Dim S As String
            Dim T As String

            MC = 0

            AmbientColor = New SharpDX.Color4(1, 1, 1, 1) ' wit
            EmissiveColor = New SharpDX.Color4(0, 0, 0, 1) ' zwart
            ReflectiveColor = New SharpDX.Color4(0, 0, 0, 1) ' zwart
            SpecularColor = New SharpDX.Color4(ShininessSlider.Value, ShininessSlider.Value, ShininessSlider.Value, 1) ' wit
            SpecularShininess = (1 - ShininessSlider.Value) * 255 ' 0..10000? hoe hoger hoe minder schittering
            EnableAutoTangent = False
            EnableFlatShading = False
            EnableTessellation = False
            RenderEnvironmentMap = False
            RenderShadowMap = False
            RenderDiffuseMap = False
            DiffuseMap = Nothing
            DisplacementMap = Nothing
            NormalMap = Nothing
            SpecularColorMap = Nothing

            If InStr(W(1), ".") > 0 Then ' het materiaal is opgegeven als een bestand
                If GetFileType(W(1)) = ".mtl" Then ' als het materiaal uit een .mtl bestand komt
                    S = W(2) ' naam
                    O = V(3) ' onderdeel
                    MaterialTrans = 1
                    T = GetFileAddress(W(1))
                    FI += 1
                    ParsFile(FI) = T
                    TextParsRuns(ParsFile(FI))
                    FI -= 1
                    Name = S
                Else ' het bestand is een afbeelding
                    W(1) = GetFileAddress(W(1)) ' bestand
                    If File.Exists(W(1)) Then
                        S = W(2) ' naam
                        O = V(3) ' onderdeel
                        DiffuseMap = LoadMapsToMemory(W(1))
                        DiffuseMapFilePath = W(1)
                        DiffuseColor = New Color4(1, 1, 1, 1)
                        AmbientColor = New Color4(AmbientSlider.Value, AmbientSlider.Value, AmbientSlider.Value, 1)
                        Name = GetFileName(W(1))
                        If W(4) <> "" Then UVCorr(0) = New Vector2(V(4), V(5)) Else UVCorr(0) = New Vector2(0, 0)
                    Else
                        MsgBox("materiaal" & vbCrLf & W(1) & vbCrLf & "niet gevonden")
                        ParsSkip = True
                        Exit Sub
                    End If
                End If
                If S <> "" Then ' als materiaal een naam heeft
                    I = 0
                    K = 0
                    For Each Node In SceneModel.SceneNode.Traverse(False)
                        K += 1
                        If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                            I += 1
                            N = Node
                            If O > 0 Then ' als het onderdeel een nummer is
                                If I = O Then
                                    N.Material = Me
                                    GoTo verder
                                End If
                            ElseIf N.Material.Name = S Then
                                N.Material = Me
                                GoTo verder
                            End If
                        End If
                    Next
                    MsgBox("onderdeel " & S & " voor " & O & " niet gevonden")
                    ParsSkip = True
verder:
                    PartMaterial = New ClassPartMaterial
                End If
            Else ' het materiaal is opgegeven als een kleur
                DiffuseColor = Color(W(1)) ' zoek de kleur
                If W(2) <> "" Then ' doorzichtig
                    DiffuseColor = New Color4(DiffuseColor.Red, DiffuseColor.Green, DiffuseColor.Blue, 1 - V(2))
                End If
                If W(3) <> "" Then ' glans
                    SpecularColor = New SharpDX.Color4(V(3), V(3), V(3), 1)
                    SpecularShininess = 255 - V(3) * 255
                End If
                Name = W(1) ' naam van de kleur
            End If
        End Sub
    End Class

    Public MC As Short ' material count
    Public MaterialTrans As Single ' transparantie
    Public AmbientColor As New TextBox
    Public DiffuseColor As New TextBox
    Public DiffusePath As New TextBox
    Public SpecularColor As New TextBox
    Public Tessalation As New TextBox
    Public UVCorr(2) As Vector2
    Public PBR As PBRMaterialCore
    Public AmbientSlider As New Slider
    Public ShininessSlider As New Slider
End Module
