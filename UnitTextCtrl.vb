Imports System.IO
Imports System.IO.Ports
Imports System.Windows.Controls
Imports System.Windows.Threading

Imports Scriptie.UnitTextCtrl

Module UnitTextCtrl

    Public TabsCtrl As New ClassCtrl

    Class ClassCtrl
        Inherits TabItem

        Function Init()

            ' start een regelaar

            CtrlLabel(1) = New Label With {
                .Content = "Port"
            }
            CtrlStack.Children.Add(CtrlLabel(1))
            CtrlText(1) = New TextBox
            CtrlStack.Children.Add(CtrlText(1))

            CtrlLabel(2) = New Label With {
                .Content = "Inputs"
            }
            CtrlStack.Children.Add(CtrlLabel(2))
            CtrlText(2) = New TextBox
            CtrlStack.Children.Add(CtrlText(2))
            CtrlText(3) = New TextBox
            CtrlStack.Children.Add(CtrlText(3))
            CtrlText(4) = New TextBox
            CtrlStack.Children.Add(CtrlText(4))

            CtrlLabel(3) = New Label With {
                .Content = "Outputs"
            }
            CtrlStack.Children.Add(CtrlLabel(3))
            CtrlText(5) = New TextBox
            CtrlStack.Children.Add(CtrlText(5))
            CtrlText(6) = New TextBox
            CtrlStack.Children.Add(CtrlText(6))
            CtrlText(7) = New TextBox
            CtrlStack.Children.Add(CtrlText(7))
            CtrlText(8) = New TextBox
            CtrlStack.Children.Add(CtrlText(8))

            Content = CtrlStack
            Header = "Ctrl"

            Return Me
        End Function

        Sub Runs()

            ' start een compoort

            ' W(0) = "control"
            ' W(1) = type
            ' W(2) = "com" & "n", nummer van compoort
            ' V(3) = baudrate
            ' V(4) = stopbits
            ' V(5) = interval in millisecondes
            ' V(6) = melders

            If V(4) = 0 Then V(4) = 1
            CtrlType = UCase(W(1))
            CtrlTime = Now.Ticks / 10000
            If CtrlType = "K8055" Then
                K8055.Init()
            Else
                With CtrlPort
                    If .IsOpen Then .Close()
                    .PortName = W(2)
                    .BaudRate = V(3)
                    .DataBits = 8
                    .Parity = IO.Ports.Parity.None
                    .StopBits = V(4)
                    .Handshake = IO.Ports.Handshake.None
                    .ReadTimeout = 500
                    .ReceivedBytesThreshold = 1
                End With
                Try
                    CtrlPort.Open()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
            If V(5) = 0 Then V(5) = 1000
            PortClock.Interval = TimeSpan.FromMilliseconds(V(5))
            PortClock.Start()
            TabsCtrl.IsSelected = True ' selecteer ctrl` tab
            InputUnits = V(6)
        End Sub

        Sub Done()

            '

            PortClock.Stop()
        End Sub

        Sub Load()

            ' laadt een .bin of .hex bestand in een ram

            ' W(0) = "load"
            ' W(1) = bestandnaam
            ' W(2) = [controleer]

            Dim B(2) As Byte
            Dim I As Short
            'Dim J As Short
            'Dim L As Short
            'Dim N As Short
            Dim T As Byte

            'Dim S() As String

            'W(1) = GetFileAddress(W(1))
            'S = File.ReadAllLines(W(1))
            'N = UBound(S)
            While CtrlPort.BytesToRead > 0
                B(0) = Read()
            End While
            T = 0
            For I = 0 To 255
                B(0) = T
                Send(B, 1)
                Wait(0.01)
                B(1) = Read()
                If B(0) <> B(1) Then MsgBox(B(0) & " " & B(1))
                T += 1
            Next
            'For I = 0 To N
            'L = Len(S(I)) - 1
            'For J = 0 To L Step 2
            'B(0) = Val(S(I)(J)) * 16 + Val(S(I)(J + 1))
            'Send(B, 1)
            'Next
            'Next
            'B(0) = 2
            'Send(B, 1)
            'Wait(1)
            'While CtrlPort.BytesToRead > 0
            'GridText.AppendText(Read() & " ")
            'End While
            'GridText.AppendText(vbCrLf)
            'B(0) = 0
            'Send(B, 1)
            'TextSaved = True
            'GridText.Show()
        End Sub

        Function Read() As Byte

            ' leest een byte uit de compoort

            Dim T As Long = Now.Ticks / 10000

            Read = 0
            If CtrlPort.IsOpen Then
                While CtrlPort.BytesToRead = 0
                    If Now.Ticks \ 10000 - T > 100 Then
                        MsgBox("Time out")
                        Exit Function
                    End If
                End While
                Read = CtrlPort.ReadByte
            End If
        End Function

        Sub Send(B As Byte(), N As Short)

            ' stuurt 1 of meerdere bytes naar de compoort

            If CtrlPort.IsOpen Then
                CtrlPort.Write(B, 0, N)
            End If
        End Sub

        Sub PortClock_Tick() Handles PortClock.Tick

            ' start telkens als het PortClock interval bereikt is

            InputS() ' leest ingangen
            Proces() ' verwerkt ingangen
            Outputs() ' schrijft uitgangen
        End Sub

        Sub InputS()

            Select Case CtrlType
                Case "CU"
                    CtrlCU.CU.Inputs()
                Case "IB"
                    CtrlIB.IB.Inputs()
                Case "GN"
                    CtrlGN.GN.Inputs()
                Case "K8055"
                    K8055.Inputs()
            End Select
        End Sub

        Sub Proces()

            Select Case CtrlType
                Case "CU"
                    CtrlCU.CU.Proces()
                Case "IB"
                    CtrlIB.IB.Proces()
                Case "GN"
                    CtrlGN.GN.Proces()
                Case "K8055"
                    K8055.Proces()
            End Select
        End Sub

        Sub Outputs()

            Select Case CtrlType
                Case "CU"
                    CtrlCU.CU.Outputs()
                Case "IB"
                    CtrlIB.IB.Outputs()
                Case "GN"
                    CtrlGN.GN.Outputs()
                Case "K8055"
                    K8055.Outputs()
            End Select
        End Sub

        Sub InputCalc(N As Byte, B As Byte)

            ' berekent de 8 melders van een meldunit byte

            Dim I As Short

            For I = 1 To 8
                Input((N - 1) * 8 + I).N = (B \ (2 ^ (8 - I)))
                B = B Mod (2 ^ (8 - I))
            Next I
        End Sub

        Sub InputShow()

            CtrlText(2).Text = "Inp  " & InputUnits & " | "
            For J = 1 To InputUnits
                For I = 1 To 8
                    CtrlText(2).Text &= Input((J - 1) * 16 + I).N & " "
                Next I
                CtrlText(2).Text &= "| "
                For I = 9 To 16
                    CtrlText(2).Text &= Input((J - 1) * 16 + I).N & " "
                Next I
                CtrlText(2).Text &= "| "
            Next J
        End Sub

        Sub OutputShow()

            CtrlText(3).Text = "Out  " & InputUnits & " | "
            For I = 1 To 16
                CtrlText(3).Text &= Output(I).N & " "
            Next I
        End Sub

        Sub ParsButton()

            ' maakt nieuwe knop

            ' W(0) = "button"
            ' V(1) = index
            ' W(2) = opschrift

        End Sub

        Sub ParsBlock()

            ' voegt een blok toe

            ' W(0) = "block"
            ' V(1) = index
            ' V(2) = vooraansluiting
            ' V(3) = achteraansluiting
            ' V(4) = melder
            ' V(5) = lengte tot melder
            ' V(6) = lengte na melder

            Blocks(V(1)).V = V(2)
            Blocks(V(1)).A = V(3)
            Blocks(V(1)).M = V(4)
            Blocks(V(1)).L = V(5)
            Blocks(V(1)).K = V(6)
        End Sub

        Sub ParsInput()

            ' definieer ingangen

            ' W(0) = "input"

        End Sub

        Sub ParsOutput()

            ' definieert een uitgang

            ' W(0) = "output"
            ' V(1) = index

            ' Output(V(1))
        End Sub

        Sub ParsRung()

            ' bepaalt een sport van een ladder bv 1,  = 1 

            ' W(0) = "rung"
            ' V(1) = output
            ' W(2)..W(20) = voorwaarden

            Output(V(1)).N = Input(V(2)).N
        End Sub

        Sub ParsTimer()
            '
            ' definieert een timer
            '
            ' W(0) = "timer"
            ' V(1) = index
        End Sub

        Sub ParsTrain()

            ' voegt een trein toe

            ' W(0) = "train"
            ' W(1) = naam
            ' W(2) = type, DI = digitaal, AN = analoog
            ' V(3) = adres
            ' V(4) = melder
            ' V(5) = top snelheid
            ' V(6) = blok
            ' V(7) = richting
            ' V(8) = dienst

            TrainI += 1
            Trains(TrainI).N = W(1)
            Trains(TrainI).Y = W(2)
            Trains(TrainI).A = V(3)
            Trains(TrainI).M = V(4)
            Trains(TrainI).T = V(5)
            Trains(TrainI).B = V(6)
            Trains(TrainI).R = W(7)
            Trains(TrainI).D = V(8)
        End Sub

        Public Structure TBlock
            Dim V As Short
            Dim A As Short
            Dim M As Short
            Dim L As Single
            Dim K As Single
        End Structure

        Public Structure TTrain
            Dim N As String
            Dim Y As String
            Dim A As Short
            Dim M As Short
            Dim T As Short
            Dim B As Short
            Dim R As Char
            Dim D As Short
        End Structure

        Public Blocks(16) As TBlock
        Public Trains(8) As TTrain
        Public TrainI As Short


        Public Sub CtrlPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles CtrlPort.DataReceived

            ' als de compoort data ontvangt

            CtrlBytes += 1
        End Sub

        Public WithEvents PortClock As New DispatcherTimer
        Public WithEvents CtrlPort As New System.IO.Ports.SerialPort
    End Class


    Class CtrlCU

        Public Shared CU As New CtrlCU

        Sub Inputs()

            Dim B(2) As Byte

            B(0) = 192 + 1 ' lees terugmelder 1
            TabsCtrl.Send(B, 1)
            TabsCtrl.InputCalc((0 * 2 + 1), TabsCtrl.Read)
            TabsCtrl.InputCalc((0 * 2 + 2), TabsCtrl.Read)
            TabsCtrl.InputShow()
        End Sub

        Sub Proces()

        End Sub

        Sub Outputs()

            Dim B(2) As Byte

            If Clear Then
                B(0) = 32
                TabsCtrl.Send(B, 2)
                Clear = False
            Else
                B(0) = 34
                B(1) = 1
                TabsCtrl.Send(B, 2)
                Clear = True
            End If
        End Sub

        Public Clear As Boolean
    End Class

    Class CtrlIB

        ' intellibox

        Public Shared IB As New CtrlIB

        Sub Inputs()

            Dim T As Long
            Dim E As Byte
            Dim B(5) As Byte

            T = Now.Ticks / 10000

            CtrlText(1).Text = T - CtrlTime
            CtrlTime = T

            For I = 1 To InputUnits ' voor alle meldunits
                B(0) = &H98         ' opdracht
                B(1) = I            ' muldunit nummer
                TabsCtrl.Send(B, 2)
                E = TabsCtrl.Read() '
                If E = 0 Then
                    TabsCtrl.InputCalc((I - 1) * 2 + 1, TabsCtrl.Read)
                    TabsCtrl.InputCalc((I - 1) * 2 + 2, TabsCtrl.Read)
                End If
            Next I
            TabsCtrl.InputShow()
        End Sub

        Sub Proces()

            If Input(1).S <> Input(1).N Then
                If Input(1).N = 1 Then
                    Output(2).N = 0
                    Output(1).N = 1
                End If
                Input(1).S = Input(1).N
            End If
            If Input(2).S <> Input(2).N Then
                If Input(2).N = 1 Then
                    Output(1).N = 0
                    Output(2).N = 1
                End If
                Input(2).S = Input(2).N
            End If
            TabsCtrl.OutputShow()
        End Sub

        Sub Outputs()

            Dim B(5) As Byte
            Dim E As Byte

            For I = 1 To 100
                If Output(I).S <> Output(I).N Then
                    B(0) = &H90 ' opdracht voor wissel
                    B(1) = 1
                    B(2) = Output(I).N * 128 + 64
                    TabsCtrl.Send(B, 3)
                    E = TabsCtrl.CtrlPort.ReadByte
                    Output(I).S = Output(I).N
                End If
            Next
        End Sub
    End Class

    Class CtrlGN

        ' generator

        Public Shared GN As New CtrlGN

        Sub Inputs()

            ' leest ingangen

            Dim B() As Byte = {255}
            Dim I As Short ' ingang index

            I = TabsCtrl.CtrlPort.BytesToRead
            CtrlText(1).Text = I
            If I = 24 Then
                For I = 1 To 24
                    Input(I).N = TabsCtrl.Read() ' lees ingangen
                Next
            End If
            While TabsCtrl.CtrlPort.BytesToRead > 0
                TabsCtrl.Read()
            End While
            CtrlText(2).Text = ""
            For I = 1 To 8
                CtrlText(2).Text &= LE(Input(I).N)
            Next
            CtrlText(3).Text = ""
            For I = 9 To 16
                CtrlText(3).Text &= LE(Input(I).N)
            Next
            CtrlText(4).Text = ""
            For I = 17 To 24
                CtrlText(4).Text &= LE(Input(I).N)
            Next
            TabsCtrl.Send(B, 1)
        End Sub

        Sub Proces()

            ' CtrlText(3).Text = CtrlText(2).Text \ 8
            'I = Val(CtrlText(2).Text) \ 4

            If Input(1).S = 255 Then Input(1).S = 0 Else Input(1).S += 1
            Output(1).N = Input(1).S
            ' TabsCtrl.OutputShow()
        End Sub

        Sub Outputs()

            If Output(1).N <> Output(1).S Then
                CtrlByte(0) = Output(1).N
                ' TabsCtrl.Send(CtrlByte, 1)
                Output(1).S = Output(1).N
            End If
        End Sub
    End Class

    Public K8055 As New ClassK8055

    Class ClassK8055

        '

        Private Declare Auto Function OpenDevice Lib "k8055d.dll" (ByVal CardAddress As Integer) As Integer
        Private Declare Sub CloseDevice Lib "k8055d.dll" ()
        Private Declare Function Version Lib "k8055d.dll" () As Integer
        Private Declare Function SearchDevices Lib "k8055d.dll" () As Integer
        Private Declare Function SetCurrentDevice Lib "k8055d.dll" (ByVal CardAddress As Integer) As Integer
        Private Declare Function ReadAnalogChannel Lib "k8055d.dll" (ByVal Channel As Integer) As Integer
        Private Declare Sub ReadAllAnalog Lib "k8055d.dll" (ByRef Data1 As Integer, ByRef Data2 As Integer)
        Private Declare Sub OutputAnalogChannel Lib "k8055d.dll" (ByVal Channel As Integer, ByVal Data As Integer)
        Private Declare Sub OutputAllAnalog Lib "k8055d.dll" (ByVal Data1 As Integer, ByVal Data2 As Integer)
        Private Declare Sub ClearAnalogChannel Lib "k8055d.dll" (ByVal Channel As Integer)
        Private Declare Sub SetAllAnalog Lib "k8055d.dll" ()
        Private Declare Sub ClearAllAnalog Lib "k8055d.dll" ()
        Private Declare Sub SetAnalogChannel Lib "k8055d.dll" (ByVal Channel As Integer)
        Public Declare Sub WriteAllDigital Lib "k8055d.dll" (ByVal Data As Integer)
        Private Declare Sub ClearDigitalChannel Lib "k8055d.dll" (ByVal Channel As Integer)
        Private Declare Sub ClearAllDigital Lib "k8055d.dll" ()
        Public Declare Sub SetDigitalChannel Lib "k8055d.dll" (ByVal Channel As Integer)
        Public Declare Sub SetAllDigital Lib "k8055d.dll" ()
        Private Declare Function ReadDigitalChannel Lib "k8055d.dll" (ByVal Channel As Integer) As Boolean
        Public Declare Function ReadAllDigital Lib "k8055d.dll" () As Integer
        Private Declare Function ReadCounter Lib "k8055d.dll" (ByVal CounterNr As Integer) As Integer
        Private Declare Sub ResetCounter Lib "k8055d.dll" (ByVal CounterNr As Integer)
        Private Declare Sub SetCounterDebounceTime Lib "k8055d.dll" (ByVal CounterNr As Integer, ByVal DebounceTime As Integer)

        Sub Init()

            '

            OpenDevice(0)

            TabsCtrl.PortClock.Interval = TimeSpan.FromMilliseconds(100)
            TabsCtrl.PortClock.Start()
        End Sub

        Sub Inputs()

            K8055Data = ReadAllDigital()
        End Sub

        Sub Proces()

            CtrlText(2).Text = K8055Data
        End Sub

        Sub Outputs()

            WriteAllDigital(K8055Data)
        End Sub
    End Class

    Public Structure TOutput

        Public N As Byte ' nieuwe stand
        Public S As Byte ' stand
    End Structure

    Public Structure TInput ' ingang type

        Public N As Byte ' nieuwe waarde
        Public S As Byte ' stand
        Public K As Double ' klok
    End Structure

    Public CtrlType As String
    Public CtrlTime As Long
    Public CtrlText(8) As TextBox
    Public CtrlLabel(8) As Label
    Public CtrlByte(8) As Byte
    Public Output(100) As TOutput
    Public InputUnits As Short
    Public CtrlBytes As Long
    Public Input(24) As TInput ' ingangen; 8 analoge + 16 x 8 digitale
    Public CtrlStack As New StackPanel
    Public K8055Data As Integer
End Module
