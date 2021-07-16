Imports System.Runtime.InteropServices
Imports System.Windows.Controls
Imports System.Windows.Threading

Module UnitMediaMidi

    Public TabsMidi As New ClassTabsMidi

    Class ClassTabsMidi
        Inherits TabItem

        Function Init() As TabItem

            MidiInLabel.Content = "Input"
            MidiStack.Children.Add(MidiInLabel)
            MidiStack.Children.Add(MidiInList)
            MidiOutLabel.Content = "Output"
            MidiStack.Children.Add(MidiOutLabel)
            MidiStack.Children.Add(MidiOutList)
            MidiTextLabel.Content = "Data"
            MidiStack.Children.Add(MidiTextLabel)
            MidiStack.Children.Add(MidiText)

            Content = MidiStack
            Header = "Midi"

            Return Me
        End Function
    End Class

    Dim OutCaps As New MidiOutCaps

    Sub PlayMidiStart()

        ' start midi in en uit

        ' W(0) = "midi"

        TabsMidi.IsSelected = True

        Dim InCaps As New MIDIINCAPS

        MidiInList.Items.Clear()
        For DevCnt = 0 To (midiInGetNumDevs - 1)
            midiInGetDevCaps(DevCnt, InCaps, Len(InCaps))
            MidiInList.Items.Add(InCaps.szPname.ToString)
        Next DevCnt
        MidiInList.SelectedIndex = 0

        MidiOutList.Items.Clear()
        For DevCnt = 0 To (midiOutGetNumDevs - 1)
            midiOutGetDevCaps(DevCnt, OutCaps, Len(OutCaps))
            MidiOutList.Items.Add(CapsName)
        Next DevCnt
        MidiOutList.SelectedIndex = 0

        midiInOpen(hMidiIn, 0, ptrCallback, 0, &H30000 Or &H20)
        midiInStart(hMidiIn)

        midiOutClose(OutHandle)
        midiOutOpen(OutHandle, 0, Nothing, 0, &H30000)

        For I = 0 To 1000
            MidiTask(I).D = True
        Next

        MidiTime = 1
        MidiTempo = 120
        MidiStart = Now.Ticks / 10000 ' milisecondes
        MidiClock.Interval = TimeSpan.FromMilliseconds(100) ' bij 120 tellen per minuut = 500 ms per tel = 62.5 ms per 32ste noot 
        MidiClock.Start()
    End Sub

    Public Structure MIDIINCAPS

        Dim wMid As Int16 ' Manufacturer ID
        Dim wPid As Int16 ' Product ID
        Dim vDriverVersion As Integer ' Driver version
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Dim szPname As String ' Product Name
        Dim dwSupport As Integer ' Reserved
    End Structure

    Function CapsName() As String

        Dim S As String

        S = Chr(OutCaps.szPname01) & Chr(OutCaps.szPname02) & Chr(OutCaps.szPname03) & Chr(OutCaps.szPname04) & Chr(OutCaps.szPname05) & Chr(OutCaps.szPname06) & Chr(OutCaps.szPname07) & Chr(OutCaps.szPname08) &
            Chr(OutCaps.szPname09) & Chr(OutCaps.szPname10) & Chr(OutCaps.szPname11) & Chr(OutCaps.szPname12) & Chr(OutCaps.szPname13) & Chr(OutCaps.szPname14) & Chr(OutCaps.szPname15) & Chr(OutCaps.szPname16) &
            Chr(OutCaps.szPname17) & Chr(OutCaps.szPname18) & Chr(OutCaps.szPname19) & Chr(OutCaps.szPname20) & Chr(OutCaps.szPname21) & Chr(OutCaps.szPname22) & Chr(OutCaps.szPname23) & Chr(OutCaps.szPname24) &
            Chr(OutCaps.szPname25) & Chr(OutCaps.szPname26) & Chr(OutCaps.szPname27) & Chr(OutCaps.szPname28) & Chr(OutCaps.szPname29) & Chr(OutCaps.szPname30) & Chr(OutCaps.szPname31) & Chr(OutCaps.szPname32)
        For I = 0 To 31
            If Asc(S(I)) = 0 Then
                For J = I To 32
                    S = Left(S, I)
                    GoTo verder
                Next
            End If
        Next
verder:
        Return S
    End Function

    Public Structure MidiOutCaps

        Dim wMid As Int16 ' Manufacturer ID
        Dim wPid As Int16 ' Product ID
        Dim vDriverVersion As Integer ' Driver version
        Dim szPname01 As Byte
        Dim szPname02 As Byte
        Dim szPname03 As Byte
        Dim szPname04 As Byte
        Dim szPname05 As Byte
        Dim szPname06 As Byte
        Dim szPname07 As Byte
        Dim szPname08 As Byte
        Dim szPname09 As Byte
        Dim szPname10 As Byte
        Dim szPname11 As Byte
        Dim szPname12 As Byte
        Dim szPname13 As Byte
        Dim szPname14 As Byte
        Dim szPname15 As Byte
        Dim szPname16 As Byte
        Dim szPname17 As Byte
        Dim szPname18 As Byte
        Dim szPname19 As Byte
        Dim szPname20 As Byte
        Dim szPname21 As Byte
        Dim szPname22 As Byte
        Dim szPname23 As Byte
        Dim szPname24 As Byte
        Dim szPname25 As Byte
        Dim szPname26 As Byte
        Dim szPname27 As Byte
        Dim szPname28 As Byte
        Dim szPname29 As Byte
        Dim szPname30 As Byte
        Dim szPname31 As Byte
        Dim szPname32 As Byte
        '<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Dim szPname As String ' Product Name
        Dim dwSupport As Integer ' Reserved
    End Structure

    Sub PlayMidiDone()

        midiInStop(hMidiIn)
        midiInClose(hMidiIn)
    End Sub

    Sub PlayMidiBar()

        ' stelt een maat in

        ' W(0) = "bar"
        ' V(1) = nummer

        MidiBar = V(1)
    End Sub

    Sub PlayMidiVoice()

        ' stelt een instrument in

        ' W(0) = "voice"
        ' V(1) = instrument nummer
        '
        For I = 0 To 15
            If W(I + 1) <> "" Then
                MidiSendMessage(New MidiShortMessage(&HC0, I, V(I + 1), 0))
            End If
        Next
    End Sub

    Sub PlayMidiNote()
        '
        ' W(0) = "note"
        ' V(1) = tel
        ' W(n) = noot kanaal n - 2
        '
        Dim A As String ' akkoord
        Dim I As Short
        Dim J As Short
        Dim K As Short
        Dim T1 As Long ' begin tijd
        Dim T2 As Long ' eind tijd
        Dim N(3) As Short ' noot
        Dim O As String = "A#BC#D#EF#G#"
        Dim S As Short ' aanslag
        Dim D As Single = 60000 / MidiTempo ' milisecondes per kwart noot 

        T1 = MidiBar * D * MidiTime + (V(1) - 1) * D ' een tempo van 120 tellen per minuut geeft 1 tel per 500 ms 

        For I = 0 To 15
            If W(I + 2) <> "" Then ' lees noot b.v. 4 C# 1 127 = oktaaf 4 noot C# = 61 duurt 1 tel aanslag 127
                J = InStr(W(I + 2), " ") ' zoek 1e spatie
                If Left(W(I + 2), 1) = "[" Then ' akkoord
                    K = InStr(W(I + 2), "]")
                    A = Mid(W(I + 2), 2, K - 2)
                    Select Case A
                        Case "A"
                            N(1) = 45 'a
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Am"
                            N(1) = 45 'a
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                        Case "B"
                            N(1) = 47
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Bm"
                            N(1) = 47
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                        Case "C"
                            N(1) = 48
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Cm"
                            N(1) = 48
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                        Case "D"
                            N(1) = 50
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Dm"
                            N(1) = 50
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                        Case "E"
                            N(1) = 52
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Em"
                            N(1) = 52
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                        Case "F"
                            N(1) = 53
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Fm"
                            N(1) = 53
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                        Case "G"
                            N(1) = 55
                            N(2) = N(1) + 4
                            N(3) = N(1) + 7
                        Case "Gm"
                            N(1) = 55
                            N(2) = N(1) + 3
                            N(3) = N(1) + 7
                    End Select
                    W(I + 2) = Mid(W(I + 2), J + 1)
                    T2 = T1 + Val(W(I + 2)) * D
                    For K = 1 To 3
                        J = 0
                        While MidiTask(J).D = False
                            J += 1
                        End While
                        MidiTask(J).T = T1
                        MidiTask(J).C = &H90
                        MidiTask(J).K = I ' kanaal
                        MidiTask(J).P1 = N(K)
                        MidiTask(J).P2 = 127
                        MidiTask(J).D = False

                        J = 0
                        While MidiTask(J).D = False
                            J += 1
                        End While
                        MidiTask(J).T = T2
                        MidiTask(J).C = &H90
                        MidiTask(J).K = I
                        MidiTask(J).P1 = N(K)
                        MidiTask(J).P2 = 0
                        MidiTask(J).D = False

                    Next
                Else
                    N(0) = Val(Left(W(I + 2), J - 1)) ' woord voor 1e spatie = oktaaf nummer of noot nummer
                    If N(0) < 8 Then ' het is een oktaaf nummer
                        N(0) *= 12 ' oktaaf 4 = 48
                        N(0) += 8 ' = 56
                        W(I + 2) = Trim(Mid(W(I + 2), J + 1))
                        N(0) += InStr(O, UCase(Left(W(I + 2), 1))) ' C = 4; = 60
                        If Mid(W(I + 2), 2, 1) = "#" Then N(0) += 1 ' C# = 61
                        J = InStr(W(I + 2), " ")
                        W(I + 2) = Trim(Mid(W(I + 2), J + 1))
                    End If
                    J = InStr(W(I + 2), " ") ' zoek volgende spatie
                    If J = 0 Then
                        T2 = T1 + Val(W(I + 2)) * D ' noot duur
                        S = 127
                    Else
                        T2 = T1 + Val(Left(W(I + 2), J - 1)) * D ' noot duur
                        S = Val(Trim(Mid(W(I + 2), J + 1))) ' rest van woord is aanslag snelheid
                    End If

                    J = 0
                    While MidiTask(J).D = False
                        J += 1
                    End While
                    MidiTask(J).T = T1
                    MidiTask(J).C = &H90
                    MidiTask(J).K = I ' kanaal
                    MidiTask(J).P1 = N(0)
                    MidiTask(J).P2 = S
                    MidiTask(J).D = False

                    J = 0
                    While MidiTask(J).D = False
                        J += 1
                    End While
                    MidiTask(J).T = T2
                    MidiTask(J).C = &H90
                    MidiTask(J).K = I
                    MidiTask(J).P1 = N(0)
                    MidiTask(J).P2 = 0
                    MidiTask(J).D = False
                End If
            End If
        Next
    End Sub

    Sub PlayMidiTempo()
        '
        ' W(0) = "tempo"
        ' V(1) = maat teller
        ' V(2) = maat noemer
        ' V(3) = tempo
        '
        MidiTime = V(1) / V(2) * 4
        If W(3) <> "" Then MidiTempo = V(3)
    End Sub

    Sub DisplayData(dwParam1)

        ' toont een midi in boodschap

        If ((HideMidiSysMessages = True) And ((dwParam1 And &HF0) = &HF0)) Then
            Exit Sub
        Else
            StatusByte = (dwParam1 And &HFF)
            DataByte1 = (dwParam1 And &HFF00) >> 8
            DataByte2 = (dwParam1 And &HFF0000) >> 16
            'CtrlText(1).Text = (String.Format("{0:X2} {1:X2} {2:X2}", StatusByte, DataByte1, DataByte2))
        End If
    End Sub

    Function MidiInProc(ByVal hMidiIn As IntPtr, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer

        ' dit is een soort interrupt die start als een noot wordt aangeslagen

        MidiData = dwParam1

        DisplayData(MidiData)
        If DataByte2 > 0 Then DataByte2 = 127
        MidiSendMessage(New MidiShortMessage(StatusByte, 0, DataByte1, DataByte2))

        Return 1
    End Function

    Sub MidiSendMessage(message As MidiShortMessage)

        midiOutShortMsg(OutHandle, message)
    End Sub

    Sub MidiClock_Tick() Handles MidiClock.Tick

        ' als midiklok tikt

        'Dim T As Long

        'DisplayData(MidiData)
        MidiText.Text = (String.Format("{0:X2} {1:X2} {2:X2}", StatusByte, DataByte1, DataByte2))
        'If MidiData0 <> MidiData Then
        '    MidiSendMessage(New MidiShortMessage(StatusByte, 1, DataByte1, DataByte2))
        'MidiData0 = MidiData
        'End If
        'T = Now.Ticks / 10000
        'MenuText3.Text = (T - MidiStart) Mod 250
        'For I = 0 To 1000
        'If MidiTask(I).D = False Then
        'If MidiTask(I).T < T - MidiStart Then
        'MidiSendMessage(New MidiShortMessage(MidiTask(I).C, MidiTask(I).K, MidiTask(I).P1, MidiTask(I).P2))
        'MidiTask(I).D = True
        'End If
        'End If
        'Next
    End Sub

    Public Structure TMidiTask
        Dim D As Boolean ' done
        Dim T As Long ' tijd
        Dim C As Single ' commando
        Dim K As Short ' kanaal
        Dim P1 As Short ' parameter 1
        Dim P2 As Short ' parameter 2
    End Structure

    Public Structure MidiShortMessage

        Private Const CHANNEL_MASK As Byte = &HF
        Private Const STATUS_MASK As Byte = &HF0

        Private data0 As Byte
        Private data1 As Byte
        Private data2 As Byte
        Private ReadOnly data3 As Byte

        Public Sub New(command As MessageCommandMask, midiChannel As Byte, value1 As Byte, value2 As Byte)

            StatusCommand = command
            Channel = midiChannel
            Parameter1 = value1
            Parameter2 = value2
        End Sub

        Public Property StatusCommand As MessageCommandMask
            Get
                Return CType(data0 >> 4, MessageCommandMask)
            End Get
            Set(value As MessageCommandMask)
                data0 = value Or (data0 And CHANNEL_MASK)
            End Set
        End Property

        Public Property Channel As Byte
            Get
                Return (data0 And CHANNEL_MASK)
            End Get
            Set(value As Byte)
                data0 = (data0 And STATUS_MASK) Or (value And CHANNEL_MASK)
            End Set
        End Property

        Public Property Parameter1 As Byte
            Get
                Return data1
            End Get
            Set(value As Byte)
                data1 = value
            End Set
        End Property

        Public Property Parameter2 As Byte
            Get
                Return data2
            End Get
            Set(value As Byte)
                data2 = value
            End Set
        End Property

        Public Shared Widening Operator CType(target As MidiShortMessage) As Integer

            Return BitConverter.ToInt32({target.data0, target.data1, target.data2, target.data3}, 0)
        End Operator

        Public Enum MessageCommandMask As Byte
            NoteOff = &H80
            NoteOn = &H90
            PolyKeyPressure = &HA0
            ControllerChange = &HB0
            ProgramChange = &HC0
            ChannelPressure = &HD0
            PitchBend = &HD0
        End Enum
    End Structure

    Sub MidiOutList_MousLeftButtonUp() Handles MidiOutList.MouseLeftButtonUp

        midiOutClose(OutHandle)
        midiOutOpen(OutHandle, MidiOutList.SelectedIndex, Nothing, 0, &H30000)
    End Sub

    Structure TBA

        Dim BA() As Byte
    End Structure

    Public MidiStack As New StackPanel
    Public MidiInLabel As New Label
    Public MidiOutLabel As New Label
    Public MidiTextLabel As New Label

    Public MidiInList As New ListBox
    Public WithEvents MidiOutList As New ListBox
    Public MidiText As New TextBox
    Public MidiBar As Short ' maat nummer
    Public MidiTempo As Short
    Public MidiTime As Single
    Public MidiTask(1000) As TMidiTask
    Public WithEvents MidiClock As New DispatcherTimer
    Public MidiStart As Long
    Public MidiData As Integer
    Public MidiData0 As Integer

    Dim StatusByte As Byte
    Dim DataByte1 As Byte
    Dim DataByte2 As Byte
    Public HideMidiSysMessages As Boolean = False

    Dim hMidiIn As IntPtr ' midiinhandle

    Public OutHandle As IntPtr ' midiouthandle
    Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)

    Public Delegate Sub MidiOutProc(HandleOut As IntPtr, msg As Integer, instance As Integer, param1 As Integer, param2 As Integer)
    Public Delegate Function MidiInCallback(ByVal hMidiIn As IntPtr, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer

    Public Declare Function midiInOpen Lib "winmm.dll" (ByRef hMidiIn As IntPtr, ByVal uDeviceID As Integer, ByVal dwCallback As MidiInCallback, ByVal dwInstance As Integer, ByVal dwFlags As Integer) As Integer
    Public Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As IntPtr) As Integer
    Public Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As IntPtr) As Integer
    Public Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As IntPtr) As Integer
    Public Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As IntPtr) As Integer

    Declare Auto Function midiOutOpen Lib "winmm.dll" (ByRef OutHandle As IntPtr, deviceId As Integer, proc As MidiOutProc, instance As Integer, flags As Integer) As Integer
    Declare Auto Function midiOutClose Lib "winmm.dll" (handle As IntPtr) As Integer
    Declare Auto Function midiOutShortMsg Lib "winmm.dll" (OutHandle As IntPtr, message As Integer) As Integer

    Public Declare Function midiInGetNumDevs Lib "winmm.dll" () As Integer
    Public Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As MIDIINCAPS, ByVal uSize As Integer) As Integer

    Public Declare Function midiOutGetNumDevs Lib "winmm.dll" () As Integer
    Public Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As MidiOutCaps, ByVal uSize As Integer) As Integer
End Module