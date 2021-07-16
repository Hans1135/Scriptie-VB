Imports System.Windows.Media
Imports HelixToolkit.Wpf.SharpDX.Animations
Imports SharpDX

Module UnitFormFunc

    Function EulerToQuaternion(H As SharpDX.Vector3) As SharpDX.Quaternion

        Dim Q As SharpDX.Quaternion

        Dim c3 As Double = Math.Cos(Rad(H.X) / 2)
        Dim s3 As Double = Math.Sin(Rad(H.X) / 2)
        Dim c1 As Double = Math.Cos(Rad(H.Y) / 2)
        Dim s1 As Double = Math.Sin(Rad(H.Y) / 2)
        Dim c2 As Double = Math.Cos(Rad(H.Z) / 2)
        Dim s2 As Double = Math.Sin(Rad(H.Z) / 2)
        Dim c1c2 As Double = c1 * c2
        Dim s1s2 As Double = s1 * s2

        Q.W = c1c2 * c3 - s1s2 * s3
        Q.X = c1c2 * s3 + s1s2 * c3
        Q.Y = s1 * c2 * c3 + c1 * s2 * s3
        Q.Z = c1 * s2 * c3 - s1 * c2 * s3
        EulerToQuaternion = Q
    End Function

    Function MatrixToKeyframe(M As SharpDX.Matrix) As Keyframe
        '
        '
        '
        Dim K As Keyframe
        '
        K.Translation = M.TranslationVector
        K.Rotation = MatrixToQuaternion(M)
        K.Scale = New SharpDX.Vector3(1, 1, 1)
        MatrixToKeyframe = K
    End Function

    Function MatrixToQuaternion(X As SharpDX.Matrix) As SharpDX.Quaternion

        Dim R As SharpDX.Quaternion
        Dim TR As Single
        Dim S As Single

        TR = X.M11 + X.M22 + X.M33

        If (TR > 0) Then
            S = Math.Sqrt(TR + 1) * 2 ' S=4*qw 
            R.W = 0.25 * S
            R.X = (X.M32 - X.M23) / S
            R.Y = (X.M13 - X.M31) / S
            R.Z = (X.M21 - X.M12) / S
        ElseIf ((X.M11 > X.M22) And (X.M11 > X.M33)) Then
            S = Math.Sqrt(1 + X.M11 - X.M22 - X.M33) * 2 ' S=4*qx 
            R.W = (X.M32 - X.M23) / S
            R.X = 0.25 * S
            R.Y = (X.M12 + X.M21) / S
            R.Z = (X.M13 + X.M31) / S
        ElseIf (X.M22 > X.M33) Then
            S = Math.Sqrt(1 + X.M22 - X.M11 - X.M33) * 2 ' S=4*qy
            R.W = (X.M13 - X.M31) / S
            R.X = (X.M12 + X.M21) / S
            R.Y = 0.25 * S
            R.Z = (X.M23 + X.M32) / S
        Else
            S = Math.Sqrt(1 + X.M33 - X.M11 - X.M22) * 2 ' S=4*qz
            R.W = (X.M21 - X.M12) / S
            R.X = (X.M13 + X.M31) / S
            R.Y = (X.M23 + X.M32) / S
            R.Z = 0.25 * S
        End If
        MatrixToQuaternion = R
    End Function

    Function QuaternionToEuler(Q As SharpDX.Quaternion) As SharpDX.Vector3

        Dim V As SharpDX.Vector3

        Dim Sqw As Single = Q.W * Q.W
        Dim Sqx As Single = Q.X * Q.X
        Dim Sqy As Single = Q.Y * Q.Y
        Dim Sqz As Single = Q.Z * Q.Z
        Dim Unit As Single = Sqx + Sqy + Sqz + Sqw ' If normalised Is one, otherwise Is correction factor
        Dim Test As Single = Q.X * Q.Y + Q.Z * Q.W

        If (Test > 0.499 * Unit) Then ' singularity at north pole
            V.Y = 2 * Math.Atan2(Q.X, Q.W) 'heading
            V.Z = Math.PI / 2 ' attitude
            V.X = 0 ' bank

        ElseIf (Test < -0.499 * Unit) Then ' singularity at south pole
            V.Y = -2 * Math.Atan2(Q.X, Q.W)
            V.Z = -Math.PI / 2
            V.X = 0
        Else
            V.Y = Math.Atan2(2 * Q.Y * Q.W - 2 * Q.X * Q.Z, Sqx - Sqy - Sqz + Sqw)
            V.Z = Math.Asin(2 * Test / Unit)
            V.X = Math.Atan2(2 * Q.X * Q.W - 2 * Q.Y * Q.Z, -Sqx + Sqy - Sqz + Sqw)
        End If
        V.X = Deg(V.X)
        V.Y = Deg(V.Y)
        V.Z = Deg(V.Z)
        QuaternionToEuler = V
    End Function

    Function QuaternionToMatrix(Q As SharpDX.Quaternion) As SharpDX.Matrix

        Dim O As SharpDX.Matrix

        Dim SQW As Single = Q.W * Q.W
        Dim SQX As Single = Q.X * Q.X
        Dim SQY As Single = Q.Y * Q.Y
        Dim SQZ As Single = Q.Z * Q.Z

        Dim Invs As Single = 1 / (SQX + SQY + SQZ + SQW)
        O.M11 = (SQX - SQY - SQZ + SQW) * Invs
        O.M22 = (-SQX + SQY - SQZ + SQW) * Invs
        O.M33 = (-SQX - SQY + SQZ + SQW) * Invs

        Dim Tmp1 As Single = Q.X * Q.Y
        Dim Tmp2 As Single = Q.Z * Q.W
        O.M21 = 2.0 * (Tmp1 + Tmp2) * Invs
        O.M12 = 2.0 * (Tmp1 - Tmp2) * Invs

        Tmp1 = Q.X * Q.Z
        Tmp2 = Q.Y * Q.W
        O.M31 = 2.0 * (Tmp1 - Tmp2) * Invs
        O.M13 = 2.0 * (Tmp1 + Tmp2) * Invs
        Tmp1 = Q.Y * Q.Z
        Tmp2 = Q.X * Q.W
        O.M32 = 2.0 * (Tmp1 + Tmp2) * Invs
        O.M23 = 2.0 * (Tmp1 - Tmp2) * Invs
        O.M44 = 1
        QuaternionToMatrix = O
    End Function

    Function QM(Q As SharpDX.Quaternion) As Media3D.Quaternion
        '
        ' omzetten van een sharpdx quaternion naar windows quaternion
        '
        QM.W = Q.W
        QM.X = Q.X
        QM.Y = Q.Y
        QM.Z = Q.Z
    End Function

    Function GetColorString(C As Color4) As String

        Return C.Red & ", " & C.Green & ", " & C.Blue & ", " & C.Alpha
    End Function
End Module
