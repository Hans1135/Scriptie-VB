Imports SharpDX

Module UnitTextGame

    Function FindNextCard(N As String) As Vector3

        Dim A As String = ""
        Dim B As String = ""
        Dim D As Vector3
        Dim I As Short = 1
        Dim Found As Boolean = False

        Select Case N
            Case "H2", "R2"
                A = "S3"
                B = "K3"
            Case "H3", "R3"
                A = "S4"
                B = "K4"
            Case "H4", "R4"
                A = "S5"
                B = "K5"
            Case "H5", "R5"
                A = "S6"
                B = "K6"
            Case "H6", "R6"
                A = "S7"
                B = "K7"
            Case "H7", "R7"
                A = "S8"
                B = "K8"
            Case "H8", "R8"
                A = "S9"
                B = "K9"
            Case "H9", "R9"
                A = "S10"
                B = "K10"
            Case "H10", "R10"
                A = "SB"
                B = "KB"
            Case "HB", "RB"
                A = "SV"
                B = "KV"
            Case "HV", "RV"
                A = "SH"
                B = "KH"

            Case "S2", "K2"
                A = "H3"
                B = "R3"
            Case "S3", "K3"
                A = "H4"
                B = "R4"
            Case "S4", "K4"
                A = "H5"
                B = "R5"
            Case "S5", "K5"
                A = "H6"
                B = "R6"
            Case "S6", "K6"
                A = "H7"
                B = "R7"
            Case "S7", "K7"
                A = "H8"
                B = "R8"
            Case "S8", "K8"
                A = "H9"
                B = "R9"
            Case "S9", "K9"
                A = "H10"
                B = "R10"
            Case "S10", "K10"
                A = "HB"
                B = "RB"
            Case "SB", "KB"
                A = "HV"
                B = "RV"
            Case "SV", "KV"
                A = "HH"
                B = "RH"
        End Select

        While Not Found And I < 53
            If (GameField(I).N = A Or GameField(I).N = B) And GameField(I).T = 0 Then
                Found = True
            Else
                I += 1
            End If
        End While

        If I < 53 Then
            D.X = GameField(I).P.X * 0.04 - 0.16
            D.Y = 0.06 - (GameField(I).P.Y + 1) * 0.01
            D.Z = GameField(I).P.Z * 0.001
        Else
            D.X = 100
        End If

        Return D
    End Function
End Module
