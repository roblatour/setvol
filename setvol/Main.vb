' SetVol 3.1, Copyright © 2021, Rob Latour  
'             https://www.raltour.com/setvol
' License MIT https://opensource.org/licenses/MIT
' Source      https://github.com/roblatour/setvol
'
' SetVol makes use of NAudio by Mark Heath
' NAudio is licensed under the Microsoft Public License (MS-PL)
' https://opensource.org/licenses/MS-PL
'
' SetVol makes use of Fody and Costura.Fody
' Both are licensed under the MIT License
' https://opensource.org/licenses/MIT

' Install-Package Costura.Fody which embeds the Naudio.dll
' https://stackoverflow.com/questions/189549/embedding-dlls-in-a-compiled-executable

' Additional thanks to IronRazerz for the significant portion of the ChangeDefaultDevice functionality as excerpted from here:
' https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/3c437708-0e90-483d-b906-63282ddd2d7b/change-audio-input

Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Module Main

    ' Global variables

    Private gEenumer As MMDeviceEnumerator = New MMDeviceEnumerator()
    Private gDev As MMDevice
    Private gRecorder As WaveInEvent
    Private Enum PlayOrRecordEnum
        Play = 0
        Record = 1
    End Enum
    Private Structure DeviceTableStructure

        Dim DeviceName As String
        Dim EndPoint As MMDevice
        Dim PayOrRecord As PlayOrRecordEnum
        Dim DefaultDevice As Boolean

    End Structure

    Private gDeviceTable() As DeviceTableStructure

    Private gAlignFlag As Boolean = False
    Private gBeepFlag As Boolean = False
    Private gErrorExplained As Boolean = False
    Private gReportFlag As Boolean = False
    Private gUnMuteFlag As Boolean = False
    Private gMuteFlag As Boolean = False

    Private gStartingColour As ConsoleColor
    Private gCurrentColour As ConsoleColor

    Private gTotalTestTimeForListingingInSeconds As Decimal = 5

    Sub Main()

        Dim ReturnCode As Integer = 0
        Dim CommandLine As String = ""

#If DEBUG Then

        'uncomment one and only one of the follow lines to run a test in debug mode

        'CommandLine = "garbage"
        CommandLine = "?"
        'CommandLine = "34"
        'CommandLine = "50%"
        'CommandLine = "+20"
        'CommandLine = "+200"
        'CommandLine = "-10"
        'CommandLine = "10"
        'CommandLine = "-15"
        'CommandLine = "mute report"
        'CommandLine = "unmute report"
        'CommandLine = "+ two"
        'CommandLine = "seventy-five"
        'CommandLine = "seventy four"
        'CommandLine = "seventy-three percent"
        'CommandLine = "1.5"
        'CommandLine = "75"
        'CommandLine = "beep device"
        'CommandLine = "unmute"
        'CommandLine = "report"
        'CommandLine = "setvol device"
        'CommandLine = "12 balance 50:75 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "35 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "50 device ASUS VP249 (NVIDIA High Definition Audio)"
        'CommandLine = "40 balance 50:100 report"
        'CommandLine = "100 balance 80:20:10 report"
        'CommandLine = "50 align report"
        'CommandLine = "12 device Spexxxxxxxxxxx)"
        'CommandLine = "45 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "12 balance 50:75 device Microphone (Realtek High Definition Audio)"
        'CommandLine = "Report"
        'CommandLine = "45 Report"
        'CommandLine = "Report device BenQ BL3200 (NVIDIA High Definition Audio)"
        'CommandLine = "Report listen device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "listen"
        'CommandLine = "listen -5"
        'CommandLine = "listen 10"
        'CommandLine = "listen 1 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "listen 32 device Microphone (Yeti Stereo Microphone)"
        'CommandLine = "Report listen 12 device Microphone (Yeti Stereo Microphone)"
        'CommandLine = "makedefault device Microphone (Yeti Stereo Microphone)"
        'CommandLine = "makedefault device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "75 beep report device Microphone (Yeti Stereo Microphone)"
        'CommandLine = "device"
        'CommandLine = "device Microphone (Yeti Stereo Microphone)"
        'CommandLine = "listen device"
        'CommandLine = "device jjjjj"


        'CommandLine = Environment.CommandLine

#Else

          CommandLine = Environment.CommandLine

#End If

        gStartingColour = Console.ForegroundColor
        gCurrentColour = gStartingColour

        Try
            gDev = gEenumer.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)
        Catch ex As Exception
        End Try

        gDeviceTable = InventoryDevices(CommandLine)

        If gDeviceTable Is Nothing Then

            Console_WriteLineInColour("No sound devices found", ConsoleColor.Red)

        Else
            ReturnCode = MainLine(CommandLine)

        End If

        Console.ForegroundColor = gStartingColour

        Environment.Exit(ReturnCode)

    End Sub
    Private Function MainLine(CommandLine As String) As Integer

        Dim ReturnCode As Integer = 0

        Dim OriginalCommandLine As String = CommandLine

        Dim StartingLevel As Integer = 0
        Dim EndingLevel As Integer = 0

        Dim DeviceChannelValue() As Integer

        Dim MasterVolumeLevelChanged As Boolean = False

        Dim AChangeInBalanceIsRequired As Boolean = False

        Try

#If DEBUG Then

            Console_WriteLineInColour("[ Command line = " & CommandLine & " ]", ConsoleColor.Cyan)

#End If

            CommandLine = CommandLine.Trim.ToUpper

            If CommandLine.Length > 0 Then

                If CommandLine.Contains("SETVOL.EXE") Then
                    CommandLine = CommandLine.Remove(0, CommandLine.IndexOf("SETVOL.EXE") + 10).TrimStart("""") ' remove the path in the command line "....SETVOL.EXE" (including the quotes)
                End If

                CommandLine = CommandLine.Replace("SETVOL", "")
                CommandLine = CommandLine.Replace("PERCENT", "")
                CommandLine = CommandLine.Replace("%", "")
                CommandLine = CommandLine.Trim

            End If

            If (CommandLine = "?") OrElse (CommandLine = "HELP") OrElse (CommandLine = String.Empty) Then

                ShowHelp()

                GoTo AllGood

            End If

            ' select device

            Dim DeviceMatchFound As Boolean = False
            Dim IndentifiedDeviceName As String = String.Empty
            Dim CommandLineContainsAValidDevice As Boolean = False

            If CommandLine.Contains("DEVICE") Then

                Dim DeviceSpecifiedOnCommandLine As String = CommandLine.Remove(0, CommandLine.IndexOf("DEVICE") + 6).Trim

                Dim DeviceSpecifiedOnCommandLineInOriginalCase = OriginalCommandLine.Remove(0, OriginalCommandLine.Length - DeviceSpecifiedOnCommandLine.Length)

                If DeviceSpecifiedOnCommandLine = "S" Then

                    ' handles the case when the user types
                    ' setvol devices
                    ' rather than
                    ' setvol device

                    DeviceSpecifiedOnCommandLine = String.Empty

                End If

                CommandLine = CommandLine.Remove(CommandLine.IndexOf("DEVICE")).Trim

                If gDeviceTable IsNot Nothing Then

                    If DeviceSpecifiedOnCommandLine.Length > 0 Then 'look for a specified device

                        For Each DeviceEntry As DeviceTableStructure In gDeviceTable

                            If DeviceSpecifiedOnCommandLine = DeviceEntry.DeviceName.ToUpper Then
                                DeviceMatchFound = True
                                IndentifiedDeviceName = DeviceEntry.DeviceName
                                gDev = DeviceEntry.EndPoint
                                CommandLineContainsAValidDevice = True
                                Exit For
                            End If

                        Next

                    End If

                End If

                If DeviceMatchFound Then

#If DEBUG Then

                    StartingLevel = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100
                    Console_WriteLineInColour("[ Starting level = " & StartingLevel & " ]", ConsoleColor.Cyan)

#End If

                Else

                    'Device specified on command line was either missing or not found

                    If DeviceSpecifiedOnCommandLine.Length > 0 Then
                        Console_WriteLineInColour(" ")
                        Console_WriteLineInColour("Specified device, '" & DeviceSpecifiedOnCommandLineInOriginalCase & "', not found", ConsoleColor.Red)
                        Console_WriteLineInColour("", ConsoleColor.White)
                    End If

                    Dim DefaultAudioDevice As String = "unknown"
                    Dim DefaultRecordingDevice As String = "unknown"

                    If gDeviceTable IsNot Nothing Then

                        For Each DeviceEntry As DeviceTableStructure In gDeviceTable

                            If DeviceEntry.DefaultDevice Then

                                If DeviceEntry.PayOrRecord = PlayOrRecordEnum.Play Then
                                    DefaultAudioDevice = DeviceEntry.DeviceName
                                Else
                                    DefaultRecordingDevice = DeviceEntry.DeviceName
                                End If

                            End If

                        Next

                    End If

                    Console_WriteLineInColour(" ")
                    Console_WriteLineInColour("The default audio device is:")
                    Console_WriteLineInColour("    " & DefaultAudioDevice)
                    Console_WriteLineInColour(" ")
                    Console_WriteLineInColour("The default recording device is:")
                    Console_WriteLineInColour("    " & DefaultRecordingDevice)

                    If gDeviceTable IsNot Nothing Then

                        Console_WriteLineInColour(" ")
                        Console_WriteLineInColour("All available devices are:")
                        Console_WriteLineInColour(" ")

                        Console_WriteLineInColour("  Audio:")
                        For Each DeviceEntry As DeviceTableStructure In gDeviceTable

                            If DeviceEntry.PayOrRecord = PlayOrRecordEnum.Play Then
                                Console_WriteLineInColour("    " & DeviceEntry.DeviceName)
                            End If

                        Next

                        Console_WriteLineInColour(" ")
                        Console_WriteLineInColour("  Recording:")

                        For Each DeviceEntry As DeviceTableStructure In gDeviceTable

                            If DeviceEntry.PayOrRecord = PlayOrRecordEnum.Record Then
                                Console_WriteLineInColour("    " & DeviceEntry.DeviceName)
                            End If

                        Next

                    End If

                    If DeviceSpecifiedOnCommandLine.Length > 0 Then
                        GoTo ErrorFound
                    End If

                End If

                    Else

                'use the default input device
                Try

                    gDev = gEenumer.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)

#If DEBUG Then

                    StartingLevel = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100
                    Console_WriteLineInColour("[ Starting level = " & StartingLevel & " ]", ConsoleColor.Cyan)

#End If

                Catch ex As Exception

                    If CommandLine.Contains("NODEFAULT") Then
                        Console_WriteLineInColour("Error: no sound device specified.", ConsoleColor.Red)
                    Else
                        Console_WriteLineInColour("Error: no sound device specified and no default sound output available.", ConsoleColor.Red)
                    End If
                    gErrorExplained = True
                    GoTo ErrorFound

                End Try

            End If

            If CommandLine.Contains("WEBSITE") Then
                Process.Start("https://www.rlatour.com/setvol")
                CommandLine = CommandLine.Replace("WEBSITE", "").Trim
            End If

            If CommandLine.Contains("ALIGN") Then
                If CommandLine.Contains("BALANCE") Then
                    Console_WriteLineInColour("Error: Align and Balance options cannot be used at the same time", ConsoleColor.Red)
                    GoTo ErrorFound
                End If
                gAlignFlag = True
                CommandLine = CommandLine.Replace("ALIGN", "").Trim
            End If

            If CommandLine.Contains("BEEP") Then
                gBeepFlag = True
                CommandLine = CommandLine.Replace("BEEP", "").Trim
            End If

            If CommandLine.Contains("REPORT") Then
                CommandLine = CommandLine.Replace("REPORT", "").Trim
                gReportFlag = True
            End If

            If CommandLine.Contains("UNMUTE") Then
                gUnMuteFlag = True
                CommandLine = CommandLine.Replace("UNMUTE", "").Trim
            End If

            If CommandLine.Contains("MUTE") Then
                gMuteFlag = True
                CommandLine = CommandLine.Replace("MUTE", "").Trim
            End If

            If gMuteFlag AndAlso gUnMuteFlag Then
                Console_WriteLineInColour(" ")
                Console_WriteLineInColour("Error: mute and unmute cannot be used at the same time", ConsoleColor.Red)
                GoTo ErrorFound
            End If

            If CommandLine.Contains("NODEFAULT") Then
                CommandLine = CommandLine.Replace("NODEFAULT", "").Trim
            End If

            If CommandLine.Contains("MAKEDEFAULT") Then
                CommandLine = CommandLine.Replace("MAKEDEFAULT", "").Trim

                If CommandLineContainsAValidDevice Then

                    SetDefaultDevice(gDev.ID)

                Else
                    Console_WriteLineInColour(" ")
                    Console_WriteLineInColour("Error: makedevice was sepecified without device option or a valid device being identified", ConsoleColor.Red)
                    GoTo ErrorFound
                End If


            End If


            If CommandLine.Contains("LISTEN") Then

                Dim sSamplingPeriod As String = String.Empty

                Try

                    ' check if there is a listen threshold after the word "Listen"
                    ' if there is set the gTotalTestTimeForListingingInSeconds 

                    Dim LookForSamplingPeriod As String = CommandLine.Remove(0, CommandLine.IndexOf("LISTEN") + "LISTEN".Length).Trim

                    Dim SamplingPeriod() = LookForSamplingPeriod.Split(" ")

                    If SamplingPeriod.Count > 0 Then

                        If SamplingPeriod(0) > String.Empty Then

                            sSamplingPeriod = SamplingPeriod(0)

                            Dim CandiateSamplingPeriod As Decimal = CDec(sSamplingPeriod)

                            If CandiateSamplingPeriod > 0 Then

                                gTotalTestTimeForListingingInSeconds = CandiateSamplingPeriod

                            Else

                                Console_WriteLineInColour("Error: invalid Listen period, should be greater than zero", ConsoleColor.Red)
                                GoTo ErrorFound

                            End If

                        End If

                    End If

                Catch ex As Exception

                    Console_WriteLineInColour("Error: invalid Listen period", ConsoleColor.Red)
                    GoTo ErrorFound

                End Try

                CommandLine = CommandLine.Replace("LISTEN", "").Trim

                If CommandLine.Length > 0 Then
                    CommandLine = CommandLine.Replace(sSamplingPeriod, "").Trim
                End If

                Console_WriteLineInColour(" ")
                If gTotalTestTimeForListingingInSeconds = 1 Then
                    Console_WriteLineInColour("Listening on '" & gDev.FriendlyName & "' for up to 1 second ...")
                Else
                    Console_WriteLineInColour("Listening on '" & gDev.FriendlyName & "' for up to " & gTotalTestTimeForListingingInSeconds & " seconds ...")
                End If

                Console_WriteLineInColour(" ", gStartingColour) 'reset the colour to the default colour in case the user breaks processing ( Cntl-C )

                Dim SoundDetected As Boolean

                If gDev.DataFlow = DataFlow.Render Then
                    SoundDetected = DetectSoundFromADevice(0.0)
                Else
                    SoundDetected = DetectSoundInCaptureDevice()
                End If

                Console_WriteLineInColour("", ConsoleColor.White)

                If SoundDetected Then
                    Console_WriteLineInColour("Sound detected", ConsoleColor.Green)
                    ReturnCode = 1
                Else
                    Console_WriteLineInColour("Sound not detected", ConsoleColor.Yellow)
                    ReturnCode = 0
                End If

            End If

            Console_WriteLineInColour("", ConsoleColor.White) ' reset console colour to white (but don't print anything)

            CommandLine = CommandLine.Trim
            If CommandLine.Length = 0 Then GoTo AllGood

            If CommandLine.Contains(".") Then
                Console_WriteLineInColour("Error: an unexpected period (""."") was found in the command line", ConsoleColor.Red)
                GoTo ErrorFound
            End If

            If CommandLine.Contains("BALANCE") Then

                Dim WorkingLine As String = CommandLine

                Dim BalanceStartsAt As Integer = WorkingLine.IndexOf("BALANCE")

                WorkingLine = WorkingLine.Remove(0, BalanceStartsAt + "BALANCE".Length).Trim

                If WorkingLine.Contains(":") Then

                    While WorkingLine.Contains(" :")
                        WorkingLine = WorkingLine.Replace(" :", ":")
                    End While
                    While WorkingLine.Contains(":  ")
                        WorkingLine = WorkingLine.Replace(":  ", ": ")
                    End While

                    Dim EndOfWorkingLine As Integer = WorkingLine.IndexOf(" ", WorkingLine.LastIndexOf(":"))
                    If EndOfWorkingLine > 0 Then
                        WorkingLine = WorkingLine.Remove(EndOfWorkingLine)
                    End If

                    Dim Values() As String = WorkingLine.Split(":")

                    If Values.Count > gDev.AudioEndpointVolume.Channels.Count Then
                        Console_WriteLineInColour("Error: the number of channels specified (" & Values.Count & ") exceeeds the number of channels (" & gDev.AudioEndpointVolume.Channels.Count & ") supported by " & gDev.FriendlyName, ConsoleColor.Red)
                        GoTo ErrorFound
                    End If

                    ReDim DeviceChannelValue(gDev.AudioEndpointVolume.Channels.Count - 1)
                    For x As Integer = 0 To gDev.AudioEndpointVolume.Channels.Count - 1
                        DeviceChannelValue(x) = -1
                    Next

                    Dim WorkingValue As Integer = 0
                    ReDim Preserve Values(gDev.AudioEndpointVolume.Channels.Count - 1)

                    For x As Integer = 0 To Values.Count - 1

                        Try

                            If Values(x) > String.Empty Then

                                WorkingValue = CInt(Values(x))

                                If (WorkingValue >= 0) AndAlso (WorkingValue <= 100) Then
                                    DeviceChannelValue(x) = CInt(Values(x))
                                Else
                                    Console_WriteLineInColour("Error: when specified the value for channel " & x + 1 & " must be between 0 and 100 inclusive", ConsoleColor.Red)
                                    GoTo ErrorFound
                                End If

                            End If

                        Catch ex As Exception

                            Console_WriteLineInColour("Error: when specified the value for channel " & x + 1 & " must be between 0 and 100 inclusive", ConsoleColor.Red)
                            GoTo ErrorFound

                        End Try

                    Next

                    AChangeInBalanceIsRequired = True

                    WorkingLine = CommandLine.ToUpper

                    Dim BalanceEndsAt As Integer = WorkingLine.ToUpper.IndexOf("BALANCE")
                    BalanceEndsAt = WorkingLine.ToUpper.LastIndexOf(":")

                    Dim KeepGoing As Boolean = True
                    Dim test As String = String.Empty
                    While KeepGoing

                        If WorkingLine.Length = BalanceEndsAt Then

                            KeepGoing = False

                        Else

                            test = Mid(WorkingLine, BalanceEndsAt)

                            If ":01234567890".Contains(test) Then
                                KeepGoing = False
                            Else
                                BalanceEndsAt += 1
                            End If
                        End If

                    End While

                    CommandLine = CommandLine.Remove(BalanceStartsAt, BalanceEndsAt - BalanceStartsAt).Trim

                Else

                    Console_WriteLineInColour("Error: balance option requires at least one colon (':')", ConsoleColor.Red)
                    GoTo ErrorFound

                End If

            End If

            Dim NewVolume As Integer

            ' +/- n
            If (CommandLine.StartsWith("+") OrElse CommandLine.StartsWith("-")) Then

                Dim WorkingLine As String = CommandLine.Remove(0, 1).Trim

                If ConvertFromWords(WorkingLine) > -1 Then
                    WorkingLine = ConvertFromWords(WorkingLine)
                End If

                Dim DeltaVolume As Integer = 0

                Try

                    DeltaVolume = WorkingLine

                    If (DeltaVolume >= 0) AndAlso (DeltaVolume <= 100) Then

                        Dim CurrentVolume As Integer = StartingLevel

                        If CommandLine.StartsWith("+") Then
                            NewVolume = CurrentVolume + DeltaVolume
                            If NewVolume > 100 Then NewVolume = 100
                        Else
                            NewVolume = CurrentVolume - DeltaVolume
                            If NewVolume < 0 Then NewVolume = 0
                        End If

                        gDev.AudioEndpointVolume.MasterVolumeLevelScalar = NewVolume / 100
                        MasterVolumeLevelChanged = True

                        GoTo AllGood

                    Else

                        Console_WriteLineInColour("Error: +/- volume value must be a number between 0 and 100 inclusive", ConsoleColor.Red)
                        gErrorExplained = True
                        GoTo ErrorFound

                    End If

                Catch ex As Exception
                    Console_WriteLineInColour("Error: +/- volume value must be a number between 0 and 100 inclusive", ConsoleColor.Red)
                    gErrorExplained = True
                    GoTo ErrorFound
                End Try

            End If

            If gDev Is Nothing Then GoTo ErrorFound

            ' zero to one hundred
            NewVolume = ConvertFromWords(CommandLine)
            If ((NewVolume >= 0) AndAlso (NewVolume <= 100)) Then
                Dim NewVolSingle As Single = NewVolume / 100
                gDev.AudioEndpointVolume.MasterVolumeLevelScalar = NewVolSingle
                MasterVolumeLevelChanged = True
                GoTo AllGood
            End If

            Try

                Dim ThisShouldBeANumber = CInt(CommandLine)

            Catch ex As Exception

                Console_WriteLineInColour("", ConsoleColor.Red)
                Console_WriteLineInColour("Error: sorry, I don't understand.", ConsoleColor.Red)
                Console_WriteLineInColour("       To view the help, please enter ""setvol ?""", ConsoleColor.Red)
                gErrorExplained = True
                GoTo ErrorFound

            End Try



            ' 0 to 100
            If ((CommandLine.Length > 0) AndAlso (CommandLine.Length < 4)) Then

                Try
                    NewVolume = CommandLine
                Catch ex As Exception
                    Console_WriteLineInColour("Error: master volume value must be between 0 and 100 inclusive", ConsoleColor.Red)
                    gErrorExplained = True
                    GoTo ErrorFound
                End Try

                If (NewVolume >= 0) AndAlso (NewVolume <= 100) Then
                    Dim NewVolSingle As Single = NewVolume / 100
                    gDev.AudioEndpointVolume.MasterVolumeLevelScalar = NewVolSingle
                    MasterVolumeLevelChanged = True
                    GoTo AllGood
                Else
                    Console_WriteLineInColour("Error: master volume value must be between 0 and 100 inclusive", ConsoleColor.Red)
                    gErrorExplained = True
                    GoTo ErrorFound
                End If

                GoTo ErrorFound

            Else

                Console_WriteLineInColour("Error: master volume value must be between 0 and 100 inclusive", ConsoleColor.Red)
                gErrorExplained = True
                GoTo ErrorFound

            End If

            If AChangeInBalanceIsRequired Then
            Else
                GoTo ErrorFound
            End If

        Catch ex As Exception

            Console_WriteLineInColour(ex.ToString, ConsoleColor.Red)
            GoTo ErrorFound

        End Try

AllGood:

        If gDev Is Nothing Then GoTo WrapUp

        If gUnMuteFlag Then gDev.AudioEndpointVolume.Mute = False

        If gMuteFlag Then gDev.AudioEndpointVolume.Mute = True

        Dim FinalVolume As Single

        If MasterVolumeLevelChanged Then
            FinalVolume = gDev.AudioEndpointVolume.MasterVolumeLevelScalar
        Else
            FinalVolume = StartingLevel / 100
        End If

        If gAlignFlag Then
            For x As Integer = 0 To gDev.AudioEndpointVolume.Channels.Count - 1
                gDev.AudioEndpointVolume.Channels(x).VolumeLevelScalar = FinalVolume
            Next
        End If

        If AChangeInBalanceIsRequired Then

            For x As Integer = 0 To gDev.AudioEndpointVolume.Channels.Count - 1
                If DeviceChannelValue(x) > -1 Then
                    gDev.AudioEndpointVolume.Channels(x).VolumeLevelScalar = FinalVolume * (DeviceChannelValue(x) / 100)
                End If
            Next

        End If

        If gReportFlag Then

            Console_WriteLineInColour("", ConsoleColor.White)

            ReturnCode = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100

            Console.WriteLine("")
            Console.WriteLine("Report for " & gDev.FriendlyName)
            Console.WriteLine("")

            If gDev.AudioEndpointVolume.Mute Then
                Console.WriteLine("Volume = " & ReturnCode.ToString & " (Muted)")
                Console.WriteLine("")
                ReturnCode *= -1
            Else

                Console.WriteLine("Master volume level = " & ReturnCode.ToString)
                Console.WriteLine("")

                For x As Integer = 0 To gDev.AudioEndpointVolume.Channels.Count - 1
                    Console.WriteLine(" Channel " & x & " level = " & Math.Round(gDev.AudioEndpointVolume.Channels(x).VolumeLevelScalar * 100))
                Next

            End If

        End If

        If gBeepFlag Then

            If gErrorExplained Then
            Else
                Beep()
            End If

        End If

        GoTo ReturnNow

ErrorFound:

        ReturnCode = -999

        If gErrorExplained Then
        Else
            Console_WriteLineInColour("Error", ConsoleColor.Red)
        End If

ReturnNow:

#If DEBUG Then

        EndingLevel = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100

        Console.WriteLine("")
        Console_WriteLineInColour("[ Ending level = " & EndingLevel & " ]", ConsoleColor.Cyan)
        Console_WriteLineInColour("[ Return code = " & ReturnCode & " ]", ConsoleColor.Cyan)

        Console.ReadLine()

#End If

WrapUp:

        If (gCurrentColour <> gStartingColour) Then
            Console_WriteLineInColour(" ", Console.ForegroundColor)
        End If

        Return ReturnCode

    End Function


    Private Function InventoryDevices(ByVal CommandLine As String) As DeviceTableStructure()

        Dim ReturnValue() As DeviceTableStructure = Nothing

        Dim DefaultDeviceName_Record As String = String.Empty
        Dim DefaultDeviceName_Play As String = String.Empty

        Try

            Dim enumerator As MMDeviceEnumerator = New MMDeviceEnumerator()

            Try
                DefaultDeviceName_Record = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia).FriendlyName
            Catch ex As Exception
            End Try

            Try
                DefaultDeviceName_Play = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia).FriendlyName
            Catch ex As Exception
            End Try

            Dim NumberOfDevices_Record As Integer = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).Count
            Dim NumberOfDevices_Play As Integer = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).Count

            ReDim ReturnValue(Math.Max(NumberOfDevices_Record + NumberOfDevices_Play - 1, 0))

            Dim Index As Integer = 0

            Try

                If NumberOfDevices_Record > 0 Then

                    For Each endpoint In enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)

                        ReturnValue(Index).DeviceName = endpoint.FriendlyName.ToString
                        ReturnValue(Index).EndPoint = endpoint
                        ReturnValue(Index).PayOrRecord = PlayOrRecordEnum.Record
                        ReturnValue(Index).DefaultDevice = (endpoint.FriendlyName.ToString = DefaultDeviceName_Record)

                        Index += 1

                    Next

                End If

            Catch ex As Exception

            End Try

            Try


                If NumberOfDevices_Play > 0 Then

                    For Each endpoint In enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)

                        ReturnValue(Index).DeviceName = endpoint.FriendlyName.ToString
                        ReturnValue(Index).EndPoint = endpoint
                        ReturnValue(Index).PayOrRecord = PlayOrRecordEnum.Play
                        ReturnValue(Index).DefaultDevice = (endpoint.FriendlyName.ToString = DefaultDeviceName_Play)

                        Index += 1

                    Next

                End If

            Catch ex As Exception

            End Try

            If (NumberOfDevices_Record > 0) OrElse (NumberOfDevices_Play > 0) Then
                ReturnValue = ReturnValue.OrderBy(Function(c) c.DeviceName).ToArray
            End If

        Catch ex As Exception

        End Try

        If CommandLine.ToUpper.Contains("NODEFAULT") Then

        Else

            If CommandLine.ToUpper.Contains("DEVICE") Then

            Else

                If (DefaultDeviceName_Play = String.Empty) Then

                    MsgBox("Error - it appears no default sound output device is set." & vbCrLf & vbCrLf &
                             "Do you have a default sound output device set on your system?" & vbCrLf & vbCrLf &
                             "Please right-click on the speaker icon in your systray, usually near your system clock at the bottom right-hand side, click on 'open sound settings' and then check what it says under 'choose your output device'.")

                End If

            End If

        End If

        Return ReturnValue

    End Function

#Region "Detect sound"

    Private Function DetectSoundFromADevice(ByVal Threshold As Decimal) As Boolean

        Dim Channels As Integer = gDev.AudioMeterInformation.PeakValues.Count

        Dim sampleCount As Integer = 0

        Dim StopAfter As DateTime = Now.AddMilliseconds(gTotalTestTimeForListingingInSeconds * 1000)

        While Now <= StopAfter

            sampleCount += 1

            For y = 0 To Channels - 1

                If gDev.AudioMeterInformation.PeakValues(y) > Threshold Then
                    Return True
                End If

            Next

            Threading.Thread.Sleep(10) ' 1/100th of a second

        End While

#If DEBUG Then
        Console_WriteLineInColour("[ Listen samples taken = " & sampleCount & " ]", ConsoleColor.Cyan)
#End If

        Return False

    End Function

    Private Function DetectSoundInCaptureDevice() As Boolean

        Dim ReturnValue As Boolean = False

        Try

            Dim DeviceID As Integer = 0

            For x As Integer = 0 To WaveIn.DeviceCount - 1

                Dim DeviceInfo = WaveIn.GetCapabilities(x)
                If gDev.FriendlyName.StartsWith(DeviceInfo.ProductName) Then
                    DeviceID = x
                    Exit For
                End If

            Next

            gRecorder = New WaveInEvent()

            gRecorder.WaveFormat = New WaveFormat(44100, gDev.AudioEndpointVolume.Channels.Count)
            gRecorder.DeviceNumber = DeviceID

            gRecorder.StartRecording()

            'even if a microphone is muted it will report some sound, so set the detection level to 1% of total volume
            ReturnValue = DetectSoundFromADevice(0.01)

            gRecorder.StopRecording()

            gRecorder.Dispose()

        Catch ex As Exception

            ex = ex

        End Try

        Return ReturnValue

    End Function


#End Region

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub ShowHelp()

        Dim StartingColour As ConsoleColor = Console.ForegroundColor

        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol v3.1 Help")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("Options:")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" ?                 help")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" n                 set the master volume to n where n = 0 to 100")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" +n                increase the master volume level by n where n = 1 to 100; maximum result = 100")
        Console_WriteLineInColour(" -n                decrease the master volume level by n where n = 1 to 100; minimum result = 0")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" x                 set the master volume level to x where x = 'zero' to 'one hundred' (without the quotes)")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" align             make all channel volume levels equal the master volume level")
        Console_WriteLineInColour("                   most devices will have at least two channel levels")
        Console_WriteLineInColour("                   with the first channel representing the left speaker and the second channel representing the right speaker")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" balance c1:c2:cz  set the audio/recording device's channel levels")
        Console_WriteLineInColour("                   cl is the desired volume level of the first channel where c1 = 0 to 100")
        Console_WriteLineInColour("                   c2 is the desired volume level of the second channel where c2 = 0 to 100")
        Console_WriteLineInColour("                   ...")
        Console_WriteLineInColour("                   cz is the desired volume level of the last channel where cz = 0 to 100")
        Console_WriteLineInColour("                   values may be included for as many channels as the device supports")
        Console_WriteLineInColour("                   for example:")
        Console_WriteLineInColour("                       50 balance 100:100")
        Console_WriteLineInColour("                        set the first and second channel (usually the left and right speakers) volume levels to 50")
        Console_WriteLineInColour("                       100 balance 50:50")
        Console_WriteLineInColour("                        also set the first and second channel volume levels to 50")
        Console_WriteLineInColour("                       100 balance 80:100:70:100:90:85:50:100")
        Console_WriteLineInColour("                        set the eight  channel volume levels to the values specified")
        Console_WriteLineInColour("                       50 balance 80:40")
        Console_WriteLineInColour("                        set the first channel volume level to 40 (50 x 80%) and the second channel volume level 20 (50 x 40%)")
        Console_WriteLineInColour("                    when the balance option is used the greatest of the channel volume levels becomes the new master volume level")
        Console_WriteLineInColour("                    in the examples above the master volume level was specified")
        Console_WriteLineInColour("                    if the master volume level is not specified the current master volume level value will be used")
        Console_WriteLineInColour("                    and the master and channel volume levels be updated as in the last example above")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" beep              play a beep")
        Console_WriteLineInColour("                   if any changes to volume levels are being made, they will be made before the beep is played")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" device name       audio/recording device name")
        Console_WriteLineInColour("                   when not used, SetVol will act upon the default audio device")
        Console_WriteLineInColour("                   when used with a valid device name, SetVol will act upon the named device")
        Console_WriteLineInColour("                   when used without a valid device name, all audio and recording devices will be displayed")
        Console_WriteLineInColour("                   when the device option is used it must be the last option specified on the command line")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" listen n          reports if a device is actively playing/recording sound")
        Console_WriteLineInColour("                   n represents the maximum number of seconds that SetVol will listen for a sound")
        Console_WriteLineInColour("                   using n is optional, but if used it should be a whole or decimal number greater than zero")
        Console_WriteLineInColour("                   when n is not specified, it will be defaulted to 5")
        Console_WriteLineInColour("                   SetVol only listens for as long as needed for it to hear a sound")
        Console_WriteLineInColour("                   SetVol will display the results and provide a return code of either 0 (nothing heard) or 1 (something heard)")
        Console_WriteLineInColour("                   this option may also be used with the device option")
        Console_WriteLineInColour("                   if the device option is not used, the default audio device will be used")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" makedefault       change the default audio/recording device")
        Console_WriteLineInColour("                   if the makedefault option is used, then the device option must be used too")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" mute              turn mute on")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" unmute            turn mute off")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" nodefault         a pop-up error message will be displayed if a device is not specified and there is no default audio device")
        Console_WriteLineInColour("                   the use of this option suppresses that message")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" report            displays a device's master and channel volume levels")
        Console_WriteLineInColour("                   the master volume level will be provided via SetVol's return code")
        Console_WriteLineInColour("                       for example: if the device's master volume level is 60, SetVol's return code will be 60")
        Console_WriteLineInColour("                   if the device is muted, the return code will be a negative number")
        Console_WriteLineInColour("                   if a volume level is changed, the changed volume level is reported")
        Console_WriteLineInColour("                   this option may also be used with the device option")
        Console_WriteLineInColour("                   if the device option is not used, the default audio device will be used")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" website           visit the SetVol website")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("Examples:")
        Console_WriteLineInColour(" setvol 75")
        Console_WriteLineInColour(" setvol +10")
        Console_WriteLineInColour(" setvol seventy-five")
        Console_WriteLineInColour(" setvol 50 balance 80:100")
        Console_WriteLineInColour(" setvol mute")
        Console_WriteLineInColour(" setvol device")
        Console_WriteLineInColour(" setvol 32 report")
        Console_WriteLineInColour(" setvol 75 device Speakers (Realtek USB2.0 Audio))")
        Console_WriteLineInColour(" setvol listen 2.5 device Speakers (Realtek USB2.0 Audio)")
        Console_WriteLineInColour(" setvol makedefault device Speakers (Realtek USB2.0 Audio)")
        Console_WriteLineInColour(" setvol makedefault device Microphone (Yeti Stereo Microphone)")
        Console_WriteLineInColour(" setvol website")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol v3.1", ConsoleColor.Yellow)
        Console_WriteLineInColour("Copyright © 2021, Rob Latour", ConsoleColor.Yellow, True)
        Console_WriteLineInColour("https://rlatour.com/setvol", ConsoleColor.Yellow)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol is licensed under the MIT License", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://opensource.org/licenses/MIT", ConsoleColor.Cyan)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol is open source", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://github.com/roblatour/setvol", ConsoleColor.Cyan)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol makes use of Fody and Costura.Fody", ConsoleColor.Cyan)
        Console_WriteLineInColour("Both are licensed under the MIT License", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://opensource.org/licenses/MIT", ConsoleColor.Cyan)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol makes use of NAudio by Mark Heath", ConsoleColor.Cyan)
        Console_WriteLineInColour("NAudio is licensed under the Microsoft Public License (MS-PL)", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://opensource.org/licenses/MS-PL", ConsoleColor.Cyan)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("Thanks to IronRazerz for a significant portion of the 'makedefault' functionality as excerpted from:", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/3c437708-0e90-483d-b906-63282ddd2d7b/change-audio-input", ConsoleColor.Cyan)
        Console_WriteLineInColour(" ")

        Console.ForegroundColor = StartingColour

    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Private Function ConvertFromWords(ByVal input As String) As Integer

        Dim NumberTable(100) As String

        NumberTable(0) = "zero"
        NumberTable(1) = "one"
        NumberTable(2) = "two"
        NumberTable(3) = "three"
        NumberTable(4) = "four"
        NumberTable(5) = "five"
        NumberTable(6) = "six"
        NumberTable(7) = "seven"
        NumberTable(8) = "eight"
        NumberTable(9) = "nine"
        NumberTable(10) = "ten"
        NumberTable(11) = "eleven"
        NumberTable(12) = "twelve"
        NumberTable(13) = "thirteen"
        NumberTable(14) = "fourteen"
        NumberTable(15) = "fifteen"
        NumberTable(16) = "sixteen"
        NumberTable(17) = "seventeen"
        NumberTable(18) = "eighteen"
        NumberTable(19) = "nineteen"
        NumberTable(20) = "twenty"
        NumberTable(21) = "twenty-one"
        NumberTable(22) = "twenty-two"
        NumberTable(23) = "twenty-three"
        NumberTable(24) = "twenty-four"
        NumberTable(25) = "twenty-five"
        NumberTable(26) = "twenty-six"
        NumberTable(27) = "twenty-seven"
        NumberTable(28) = "twenty-eight"
        NumberTable(29) = "twenty-nine"
        NumberTable(30) = "thirty"
        NumberTable(31) = "thirty-one"
        NumberTable(32) = "thirty-two"
        NumberTable(33) = "thirty-three"
        NumberTable(34) = "thirty-four"
        NumberTable(35) = "thirty-five"
        NumberTable(36) = "thirty-six"
        NumberTable(37) = "thirty-seven"
        NumberTable(38) = "thirty-eight"
        NumberTable(39) = "thirty-nine"
        NumberTable(40) = "forty"
        NumberTable(41) = "forty-one"
        NumberTable(42) = "forty-two"
        NumberTable(43) = "forty-three"
        NumberTable(44) = "forty-four"
        NumberTable(45) = "forty-five"
        NumberTable(46) = "forty-six"
        NumberTable(47) = "forty-seven"
        NumberTable(48) = "forty-eight"
        NumberTable(49) = "forty-nine"
        NumberTable(50) = "fifty"
        NumberTable(51) = "fifty-one"
        NumberTable(52) = "fifty-two"
        NumberTable(53) = "fifty-three"
        NumberTable(54) = "fifty-four"
        NumberTable(55) = "fifty-five"
        NumberTable(56) = "fifty-six"
        NumberTable(57) = "fifty-seven"
        NumberTable(58) = "fifty-eight"
        NumberTable(59) = "fifty-nine"
        NumberTable(60) = "sixty"
        NumberTable(61) = "sixty-one"
        NumberTable(62) = "sixty-two"
        NumberTable(63) = "sixty-three"
        NumberTable(64) = "sixty-four"
        NumberTable(65) = "sixty-five"
        NumberTable(66) = "sixty-six"
        NumberTable(67) = "sixty-seven"
        NumberTable(68) = "sixty-eight"
        NumberTable(69) = "sixty-nine"
        NumberTable(70) = "seventy"
        NumberTable(71) = "seventy-one"
        NumberTable(72) = "seventy-two"
        NumberTable(73) = "seventy-three"
        NumberTable(74) = "seventy-four"
        NumberTable(75) = "seventy-five"
        NumberTable(76) = "seventy-six"
        NumberTable(77) = "seventy-seven"
        NumberTable(78) = "seventy-eight"
        NumberTable(79) = "seventy-nine"
        NumberTable(80) = "eighty"
        NumberTable(81) = "eighty-one"
        NumberTable(82) = "eighty-two"
        NumberTable(83) = "eighty-three"
        NumberTable(84) = "eighty-four"
        NumberTable(85) = "eighty-five"
        NumberTable(86) = "eighty-six"
        NumberTable(87) = "eighty-seven"
        NumberTable(88) = "eighty-eight"
        NumberTable(89) = "eighty-nine"
        NumberTable(90) = "ninety"
        NumberTable(91) = "ninety-one"
        NumberTable(92) = "ninety-two"
        NumberTable(93) = "ninety-three"
        NumberTable(94) = "ninety-four"
        NumberTable(95) = "ninety-five"
        NumberTable(96) = "ninety-six"
        NumberTable(97) = "ninety-seven"
        NumberTable(98) = "ninety-eight"
        NumberTable(99) = "ninety-nine"
        NumberTable(100) = "one hundred"

        Dim TestCase As String

        TestCase = input.ToLower.Trim
        For x As Integer = 0 To 100
            If NumberTable(x) = TestCase Then
                Return x
            End If
        Next

        TestCase = TestCase.Replace(" ", "-")
        For x As Integer = 21 To 99
            If NumberTable(x) = TestCase Then
                Return x
            End If
        Next

        Return -1

    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub Console_WriteLineInColour(ByVal Message As String, Optional ByVal Colour As ConsoleColor = ConsoleColor.White, Optional PrintSpecialCharacters As Boolean = False)

        Dim OriginalTextEncoding As System.Text.Encoding

        Try

            If PrintSpecialCharacters Then
                OriginalTextEncoding = System.Console.OutputEncoding
                System.Console.OutputEncoding = System.Text.Encoding.Unicode
            End If

            If (Console.BackgroundColor = Colour) Then
                Colour = ConsoleColor.Black
            End If

            Console.ForegroundColor = Colour

            If Message.Length > 0 Then
                Console.WriteLine(Message)
            End If

            If Colour = ConsoleColor.Red Then
                gErrorExplained = True
            End If

            gCurrentColour = Colour

            If PrintSpecialCharacters Then
                System.Console.OutputEncoding = OriginalTextEncoding
            End If

        Catch ex As Exception

            MsgBox(Message & vbCrLf & ex.ToString)

        End Try

    End Sub

End Module