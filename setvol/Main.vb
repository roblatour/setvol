' SetVol 4.2, Copyright © 2025, Rob Latour  
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
Imports System
Imports System.Linq
Imports System.Diagnostics
Imports System.ServiceProcess
Imports System.Runtime.InteropServices

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
        Dim DefaultDeviceComm As Boolean

    End Structure

    Private gDeviceTable() As DeviceTableStructure

    Private gDebugFlag As Boolean = False
    Private gAlignFlag As Boolean = False
    Private gBeepFlag As Boolean = False
    Private gErrorExplained As Boolean = False
    Private gReportFlag As Boolean = False
    Private gUnMuteFlag As Boolean = False
    Private gMuteFlag As Boolean = False

    Private gStartingColour As ConsoleColor
    Private gCurrentColour As ConsoleColor

    Private gTotalTestTimeForListingingInSeconds As Decimal = 5

    Private gStartingLevel As Integer = 0
    Private gFadeTime As Single = -1

    Private gSessions As SessionCollection
    Private gSessionIndex As New List(Of Integer)

    Private gSetAppAudioVolume As Boolean = False
    Private gSetAppAudioVolumeForSystemSounds As Boolean = False

    Private Const specialCopyrightChar As Char = ChrW(&HA9)
    Private Const specialRegistratrionChar As Char = ChrW(&HAE)

    Sub Main()

        Dim ReturnCode As Integer = 0
        Dim CommandLine As String = ""

#If DEBUG Then

        'uncomment one and only one of the follow lines to run a test in debug mode

        'CommandLine = Environment.CommandLine

        'CommandLine = "garbage"
        'CommandLine = "?"
        'CommandLine = "35"
        'CommandLine = "50%"
        'CommandLine = "+20"
        'CommandLine = "+200"
        'CommandLine = "-10"
        'CommandLine = "50 over 15"
        'CommandLine = "50 over -25"
        'CommandLine = "71 over 10.55555"
        'CommandLine = "5c0 over 10.5"
        'CommandLine = "0 over 60"
        'CommandLine = "mute report"
        'CommandLine = "unmute report"
        'CommandLine = "+ two"
        'CommandLine = "seventy-five"
        'CommandLine = "seventy four"
        'CommandLine = "seventy-three percent"
        'CommandLine = "1.5"
        'CommandLine = "beep device"
        'CommandLine = "unmute"
        'CommandLine = "report"
        'CommandLine = "report device Microphone (Yeti Stereo Microphone)"
        'CommandLine = "report device Headphones (2- Realtek USB2.0 Audio)"
        'CommandLine = "device"
        'CommandLine = "12 balance 50:75 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "35 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "50 device ASUS VP249 (NVIDIA High Definition Audio)"
        'CommandLine = "40 balance 50:100 report"
        'CommandLine = "100 balance 80:20:10 report"
        'CommandLine = "50 align report"
        'CommandLine = "12 device Spexxxxxxxxxxx)"
        'CommandLine = "45 device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "12 balance 50:75 device Microphone (Realtek High Definition Audio)"
        'CommandLine = "50 Report"
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
        'CommandLine = "debug 50 device Microphone (Yeti Stereo Microphone) debug"
        'CommandLine = "listen device"
        'CommandLine = "device jjjjj"
        'CommandLine = "appaudio"
        'CommandLine = "50 appaudio snagiteditor"
        'CommandLine = "50 appaudio systemsounds"
        CommandLine = "mute appaudio systemsounds"
        'CommandLine = "-15 over 15 appaudio vlc"
        'CommandLine = "debug 32 appaudio vlc"
        'CommandLine = "32 appaudio vlc debug"
        'CommandLine = "25 appaudio vlc"
        'CommandLine = "mute appaudio vlc"
        'CommandLine = "unmute appaudio vlc"
        'CommandLine = "100"
        'CommandLine = "unmute device Speakers (Realtek USB2.0 Audio)"
        'CommandLine = "unmute appaudio vlc debug"
        'CommandLine = "appaudio"
        'CommandLine = "50 over 10 appaudio vlc"
        'CommandLine = "-12 over 10 appaudio SnagitEditor"
        'CommandLine = "100 appaudio vlc"
        'CommandLine = "debug"
        'CommandLine = "SetVol 40 balance 100:100:0:0:0:89"
        'CommandLine = "SetVol 40 balance 100:100:0:0:0:91"
        'CommandLine = "SetVol 100 balance 100:100:100:100:100:100"
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

    ' Function to check if the session is the system sounds session
    Private Function IsSystemSoundsSession(session As AudioSessionControl) As Boolean
        Try
            ' System sounds session typically does not have an associated process
            Dim processID = session.GetProcessID()
            Return processID = 0
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function MainLine(CommandLine As String) As Integer

        Dim ReturnCode As Integer = 0

        Dim OriginalCommandLine As String = CommandLine

        Dim EndingLevel As Integer = 0

        Dim DeviceChannelValue() As Integer

        Dim MasterVolumeLevelChanged As Boolean = False

        Dim AChangeInBalanceIsRequired As Boolean = False

        Dim ApplicationName As String
        Dim NewVolume As Integer

        Try

            'remove the path and program name for setvol from the command line

            If CommandLine.ToUpper.Contains("SETVOL.EXE") Then
                CommandLine = CommandLine.Remove(0, CommandLine.ToUpper.IndexOf("SETVOL.EXE") + 10).TrimStart("""") ' remove the path in the command line "....SETVOL.EXE" (including the quotes)
            End If

            If CommandLine.ToUpper.Contains("SETVOL") Then
                CommandLine = CommandLine.Remove(0, CommandLine.ToUpper.IndexOf("SETVOL") + 6).TrimStart("""") ' remove the path in the command line "....SETVOL.EXE" (including the quotes)
            End If

            CommandLine = CommandLine.Trim

#If DEBUG Then
            gDebugFlag = True
#Else
            gDebugFlag = CommandLine.ToUpper.Contains("DEBUG")
#End If

            If gDebugFlag Then

                Console_WriteLineInColour("[ Command line: " & CommandLine & " ]", ConsoleColor.Cyan)
                Console_WriteLineInColour(" ", ConsoleColor.Cyan)

            End If

            CommandLine = CommandLine.Trim.ToUpper

            If CommandLine.Length > 0 Then

                CommandLine = CommandLine.Replace("DEBUG", "")
                CommandLine = CommandLine.Replace("PERCENT", "")
                CommandLine = CommandLine.Replace("%", "")
                CommandLine = CommandLine.Trim

            End If

            If (CommandLine = "?") OrElse (CommandLine = "HELP") OrElse ((CommandLine = String.Empty) AndAlso (Not gDebugFlag)) Then

                ShowHelp()

                GoTo AllGood

            End If

            If CommandLine.Contains("DEVICE") AndAlso CommandLine.Contains("APPAUDIO") Then

                Console_WriteLineInColour("Error: the 'appaudio' and 'device' options cannot be used in the same command.", ConsoleColor.Red)
                GoTo ErrorFound

            End If

            If CommandLine.Contains("REPORT") AndAlso CommandLine.Contains("APPAUDIO") Then

                Console_WriteLineInColour("Error: the 'appaudio' and 'report' options cannot be used in the same command.", ConsoleColor.Red)
                GoTo ErrorFound

            End If

            ' select device

            Dim DeviceMatchFound As Boolean = False
            Dim IndentifiedDeviceName As String = String.Empty
            Dim CommandLineContainsAValidDevice As Boolean = False

            If CommandLine.Contains("DEVICE") Then

                Dim DeviceSpecifiedOnCommandLine As String = CommandLine.Remove(0, CommandLine.IndexOf("DEVICE") + "DEVICE".Length).Trim

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

                        Dim revisedDeviceName As String = String.Empty ' v4.2 added revised Device Name testing

                        For Each DeviceEntry As DeviceTableStructure In gDeviceTable

                            revisedDeviceName = DeviceEntry.DeviceName.Replace(specialCopyrightChar, "").Replace(specialRegistratrionChar, "").ToUpper()

                            If (DeviceSpecifiedOnCommandLine = DeviceEntry.DeviceName.ToUpper) OrElse (DeviceSpecifiedOnCommandLine = revisedDeviceName) Then
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

                    gStartingLevel = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100

                    If gDebugFlag Then
                        Console_WriteLineInColour("[ Starting master level = " & gStartingLevel & " ]", ConsoleColor.Cyan)
                    End If


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

                    gStartingLevel = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100

                    If gDebugFlag Then

                        Console_WriteLineInColour("[ Starting master level = " & gStartingLevel & " ]", ConsoleColor.Cyan)

                    End If

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

            ' this code setups to change the audio playback level for a specific application/program
            ' of note, there is no equivalent option for setting the recording level of a application/program

            If (CommandLine = "APPAUDIO") Then

                gSessions = gDev.AudioSessionManager.Sessions

                Dim count As Integer = 0

                Dim uniqueProcessNames As New List(Of String)

                For i As Integer = 0 To gSessions.Count - 1
                    Dim Process As Process = Process.GetProcessById(gSessions(i).GetProcessID)
                    If (Process.ProcessName.ToUpper = "IDLE") OrElse (Process.ProcessName.ToUpper = "SHELLEXPERIENCEHOST") Then
                    Else
                        If Not uniqueProcessNames.Contains(Process.ProcessName) Then
                            uniqueProcessNames.Add(Process.ProcessName)
                            count += 1
                        End If
                    End If
                Next

                Console_WriteLineInColour(" ", ConsoleColor.White)

                If count = 0 Then
                    Console_WriteLineInColour("Could not find any applications that can have their audio level changed by SetVol.", ConsoleColor.White)
                ElseIf count = 1 Then
                    Console_WriteLineInColour("Found one application that can have its audio level changed, it is:" & vbCrLf, ConsoleColor.White)
                Else
                    Console_WriteLineInColour("Found " & ConvertNumberToWords(count) & " applications that can have their audio levels changed, these are:" & vbCrLf, ConsoleColor.White)
                End If

                uniqueProcessNames.Clear()

                For i As Integer = 0 To gSessions.Count - 1
                    Dim Process As Process = Process.GetProcessById(gSessions(i).GetProcessID)
                    If (Process.ProcessName.ToUpper = "IDLE") OrElse (Process.ProcessName.ToUpper = "SHELLEXPERIENCEHOST") Then
                    Else
                        If Not uniqueProcessNames.Contains(Process.ProcessName) Then
                            Console_WriteLineInColour("   " & Process.ProcessName, ConsoleColor.Gray)
                            uniqueProcessNames.Add(Process.ProcessName)
                        End If
                    End If
                Next

                Console_WriteLineInColour(" ", ConsoleColor.White)

                If count = 0 Then
                    Console_WriteLineInColour("However, the level for Windows system sounds can be changed.", ConsoleColor.White)
                Else
                    Console_WriteLineInColour("Additionally, the level for Windows system sounds can be changed.", ConsoleColor.White)
                End If

                Console_WriteLineInColour(" ", ConsoleColor.White)
                Console_WriteLineInColour("To change one of the above use:", ConsoleColor.White)
                Console_WriteLineInColour("    setvol n appaudio x", ConsoleColor.White)
                Console_WriteLineInColour("    where n is the new level, between 0 and 100, or either mute or unmute", ConsoleColor.White)
                Console_WriteLineInColour("    and x is either the name of the application (as shown above) or SystemSounds", ConsoleColor.White)
                Console_WriteLineInColour(" ", ConsoleColor.White)
                Console_WriteLineInColour("note: Windows maintains the application audio levels and its system sound level relative to the master volume", ConsoleColor.White)
                Console_WriteLineInColour("      So, for example, setvol 50 appaudio Spotify", ConsoleColor.White)
                Console_WriteLineInColour("      will set Spotify's audio level to 50 percent of the current master volume", ConsoleColor.White)
                Console_WriteLineInColour("      and Windows will automatically update Spotify's audio level in relative proportions whenever the master volume changes", ConsoleColor.White)

                GoTo AllGood

            ElseIf CommandLine.Contains("APPAUDIO") Then

                Dim AppAudioSpecifiedOnCommandLine As String = CommandLine.Remove(0, CommandLine.IndexOf("APPAUDIO") + "APPAUDIO".Length).Trim

                Dim AppAudioSpecifiedOnCommandLineInOriginalCase = OriginalCommandLine.Remove(0, OriginalCommandLine.ToUpper.IndexOf("APPAUDIO") + "APPAUDIO".Length).Trim

                If AppAudioSpecifiedOnCommandLineInOriginalCase.ToUpper <> AppAudioSpecifiedOnCommandLine Then
                    AppAudioSpecifiedOnCommandLineInOriginalCase = AppAudioSpecifiedOnCommandLineInOriginalCase.Remove(AppAudioSpecifiedOnCommandLine.Length)
                End If

                Dim CommandLineContainsAValidApp As Boolean = False

                CommandLine = CommandLine.Remove(CommandLine.IndexOf("APPAUDIO")).Trim

                gSessionIndex.Clear()

                If gDeviceTable IsNot Nothing Then

                    If AppAudioSpecifiedOnCommandLine.Length > 0 Then 'look for a specified app

                        If AppAudioSpecifiedOnCommandLine = "SYSTEMSOUNDS" Then

                            gSetAppAudioVolumeForSystemSounds = True

                        Else

                            gSessions = gDev.AudioSessionManager.Sessions

                            Dim LastApplicationReported As String = ""

                            For i As Integer = 0 To gSessions.Count - 1

                                Dim Process As Process = Process.GetProcessById(gSessions(i).GetProcessID)

                                If Process.ProcessName.ToUpper = AppAudioSpecifiedOnCommandLine Then

                                    ApplicationName = AppAudioSpecifiedOnCommandLineInOriginalCase

                                    gSessionIndex.Add(i)

                                    CommandLineContainsAValidApp = True

                                    gSetAppAudioVolume = True

                                    If gDebugFlag Then
                                        If ApplicationName <> LastApplicationReported Then
                                            Console_WriteLineInColour("[ Starting audio level for application '" & ApplicationName & "' = " & gSessions(i).SimpleAudioVolume.Volume * 100 & " ]", ConsoleColor.Cyan)
                                        End If
                                    End If

                                    LastApplicationReported = ApplicationName

                                    'Exit For  'for is commented out so that we can check for multiple instances of the same application  ' testing here

                                End If

                            Next

                            If gSetAppAudioVolume Then
                            Else
                                Console_WriteLineInColour("Error: an application called '" & AppAudioSpecifiedOnCommandLine & "' does not appear to be running.", ConsoleColor.Red)
                                GoTo ErrorFound
                            End If

                        End If

                    End If

                End If

            End If

            If CommandLine.Contains("WEBSITE") Then
                Process.Start("https://github.com/roblatour/setvol")
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


            Dim ws As String = CommandLine
            If ws.Contains("MAKEDEFAULTCOMM") Then
                ws = ws.Replace("MAKEDEFAULTCOMM", "").Trim
                If ws.Contains("MAKEDEFAULT") Then
                    Console_WriteLineInColour(" ")
                    Console_WriteLineInColour("Error: makedefault and makedefaultcomm cannot be used at the same time", ConsoleColor.Red)
                    GoTo ErrorFound
                End If
            End If


            If CommandLine.Contains("MAKEDEFAULTCOMM") Then

                CommandLine = CommandLine.Replace("MAKEDEFAULTCOMM", "").Trim

                If CommandLineContainsAValidDevice Then

                    SetDefaultDevice(gDev.ID, True)

                Else
                    Console_WriteLineInColour(" ")
                    Console_WriteLineInColour("Error: makedevicecomm was specified without device option or a valid device being identified", ConsoleColor.Red)
                    GoTo ErrorFound
                End If

            End If


            If CommandLine.Contains("MAKEDEFAULT") Then
                CommandLine = CommandLine.Replace("MAKEDEFAULT", "").Trim

                If CommandLineContainsAValidDevice Then

                    SetDefaultDevice(gDev.ID, False)

                Else
                    Console_WriteLineInColour(" ")
                    Console_WriteLineInColour("Error: makedevice was specified without device option or a valid device being identified", ConsoleColor.Red)
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


            If CommandLine.Contains("OVER") Then

                Const ErrorMsg As String = "Error: the over option must be followed by a positive number greater than zero but less than or equal to 86,400."

                Dim WorkingLine As String = CommandLine.ToUpper

                Dim OverTimeNumberStartsAt As Integer = WorkingLine.IndexOf("OVER") + "OVER".Length

                Dim OverTimeNumberEndsAt As Integer = OverTimeNumberStartsAt

                Dim KeepGoing As Boolean = True
                Dim test As String = String.Empty

                While KeepGoing

                    If WorkingLine.Length = OverTimeNumberEndsAt Then

                        KeepGoing = False

                    Else

                        test = Mid(WorkingLine, OverTimeNumberEndsAt)

                        If " -.,01234567890".Contains(test) Then
                            KeepGoing = False
                        Else
                            OverTimeNumberEndsAt += 1
                        End If
                    End If

                End While

                WorkingLine = WorkingLine.Remove(0, OverTimeNumberStartsAt).Trim

                Dim Pos As Integer = WorkingLine.IndexOf(" ")

                Dim sFadeTime As String

                If Pos > 0 Then
                    sFadeTime = WorkingLine.Remove(Pos)
                Else
                    sFadeTime = WorkingLine
                End If

                Try

                    ' Convert.ToString (used below) allows a number with multiple inappropriately placed commas, such as "8,6,500"
                    ' the edit below weeds that out 

                    If sFadeTime.Split(",").Length - 1 > 1 Then
                        Console_WriteLineInColour(ErrorMsg, ConsoleColor.Red)
                        GoTo ErrorFound
                    End If

                    gFadeTime = Convert.ToSingle(sFadeTime)

                    ' Convert.ToString (used above) allows a number with an incorrectly placed comma, such as "8,6.5"
                    ' the edit below weeds that out 

                    If sFadeTime.Split(",").Length - 1 = 1 Then
                        If gFadeTime >= 1000 Then
                        Else
                            Console_WriteLineInColour(ErrorMsg, ConsoleColor.Red)
                            GoTo ErrorFound
                        End If
                    End If

                    gFadeTime = CInt(gFadeTime * 1000) / 1000  'Convert to three decimal places, can't adjust more accurately than that

                    If (gFadeTime <= 0) OrElse (gFadeTime > 86400) Then
                        Console_WriteLineInColour(ErrorMsg, ConsoleColor.Red)
                        GoTo ErrorFound
                    End If

                Catch ex As Exception

                    Console_WriteLineInColour(ErrorMsg, ConsoleColor.Red)
                    GoTo ErrorFound

                End Try

                WorkingLine = CommandLine.ToUpper

                CommandLine = CommandLine.Remove(OverTimeNumberStartsAt - "OVER".Length, OverTimeNumberEndsAt - OverTimeNumberStartsAt + "OVER".Length).Trim

            End If

            Console_WriteLineInColour("", ConsoleColor.White) ' reset console colour to white (but don't print anything)

            CommandLine = CommandLine.Trim

            If CommandLine.Length = 0 Then GoTo SetAppAudioVolume

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

                            test = Mid(WorkingLine, BalanceEndsAt, 1)

                            If ":01234567890".Contains(test) Then
                                BalanceEndsAt += 1
                            Else
                                KeepGoing = False
                            End If
                        End If

                    End While

                    CommandLine = CommandLine.Remove(BalanceStartsAt, BalanceEndsAt - BalanceStartsAt).Trim

                Else

                    Console_WriteLineInColour("Error: balance option requires at least one colon (':')", ConsoleColor.Red)
                    GoTo ErrorFound

                End If

            End If

            'determine if string starts with a number
            ' +/- n
            CommandLine = CommandLine.Trim

            If CommandLine.StartsWith("+") OrElse CommandLine.StartsWith("-") OrElse (IsNumeric(CommandLine.Substring(0, 1)) AndAlso gFadeTime > -1) Then

                Dim WorkingLine As String = CommandLine

                If WorkingLine.StartsWith("+") OrElse WorkingLine.StartsWith("-") Then
                    WorkingLine = WorkingLine.Remove(0, 1).Trim()
                End If

                If ConvertFromWords(WorkingLine) > -1 Then
                    WorkingLine = ConvertFromWords(WorkingLine)
                End If

                Dim DeltaVolume As Integer = 0

                Try

                    DeltaVolume = WorkingLine

                    If (DeltaVolume >= 0) AndAlso (DeltaVolume <= 100) Then

                        If gSetAppAudioVolume Then

                            ' in most cases if there are multiple sessions, they will all have the same starting volume
                            ' regardless, if there are multiple sessions, 
                            ' we will to use the highest volume of all them as the starting point for fading down, and
                            ' the lowest volume of all them as the starting point for fading up

                            If CommandLine.StartsWith("-") Then
                                gStartingLevel = 0
                                For Each i In gSessionIndex
                                    gStartingLevel = Math.Max(gStartingLevel, gSessions(i).SimpleAudioVolume.Volume * 100)
                                Next
                            Else
                                gStartingLevel = 100
                                For Each i In gSessionIndex
                                    gStartingLevel = Math.Min(gStartingLevel, gSessions(i).SimpleAudioVolume.Volume * 100)
                                Next
                            End If

                        End If

                        Dim CurrentVolume As Integer = gStartingLevel

                        If CommandLine.StartsWith("-") Then
                            NewVolume = CurrentVolume - DeltaVolume
                            If NewVolume < 0 Then NewVolume = 0
                        ElseIf CommandLine.StartsWith("+") Then
                            NewVolume = CurrentVolume + DeltaVolume
                            If NewVolume > 100 Then NewVolume = 100
                        Else
                            NewVolume = CInt(CommandLine)
                            If NewVolume > 100 Then NewVolume = 100
                        End If

                        Dim NewVolSingle As Single = NewVolume / 100
                        SetTheLevel(NewVolSingle)
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

SetAppAudioVolume:

            'set app audio level

            If gSetAppAudioVolume Then

                Dim LastApplicationReported As String = ""

                For Each i In gSessionIndex

                    If gDebugFlag AndAlso (ApplicationName <> LastApplicationReported) Then
                        Console.WriteLine("")
                        Console_WriteLineInColour("[ The original audio level for application '" & ApplicationName & "' = " & gSessions(i).SimpleAudioVolume.Volume.ToString * 100 & " ]", ConsoleColor.Cyan)
                    End If

                    If (CommandLine = "") AndAlso (Not gMuteFlag) AndAlso (Not gUnMuteFlag) Then
                        If (ApplicationName <> LastApplicationReported) Then
                            Console_WriteLineInColour("Error: a new audio level was not specified for application '" & ApplicationName & "'.", ConsoleColor.Red)
                        End If
                        gErrorExplained = True ' changed this from a warning to an error in version 4.4
                        ' NewVolume = gSessions(i).SimpleAudioVolume.Volume * 100 ' use the current level
                        GoTo ErrorFound
                    End If

                    If gMuteFlag Then

                        gSessions(i).SimpleAudioVolume.Mute = True
                        If gDebugFlag AndAlso (ApplicationName <> LastApplicationReported) Then
                            If (ApplicationName <> LastApplicationReported) Then
                                Console_WriteLineInColour("[ Application '" & ApplicationName & "' was muted ]", ConsoleColor.Cyan)
                            End If
                        End If
                        gMuteFlag = False

                    ElseIf gUnMuteFlag Then

                        gSessions(i).SimpleAudioVolume.Mute = False
                        If gDebugFlag AndAlso (ApplicationName <> LastApplicationReported) Then
                            If (ApplicationName <> LastApplicationReported) Then
                                Console_WriteLineInColour("[ Application '" & ApplicationName & "' was unmuted ]", ConsoleColor.Cyan)
                            End If
                        End If
                        gUnMuteFlag = False

                    Else

                        If Int(gSessions(i).SimpleAudioVolume.Volume * 100) = NewVolume Then

                            If gDebugFlag AndAlso (ApplicationName <> LastApplicationReported) Then
                                If (ApplicationName <> LastApplicationReported) Then
                                    Console_WriteLineInColour("[ The audio level for application '" & ApplicationName & "' did not need to be updated ]", ConsoleColor.Yellow)
                                End If
                            End If

                        ElseIf (NewVolume < 0 OrElse NewVolume > 100) Then

                            If gDebugFlag AndAlso (ApplicationName <> LastApplicationReported) Then
                                If (ApplicationName <> LastApplicationReported) Then
                                    Console_WriteLineInColour("[ The audio level for application '" & ApplicationName & "' was not updated as the update value specified was not between 0 and 100 inclusive ]", ConsoleColor.Red)
                                End If
                            End If

                        Else

                            Dim NewVolSingle As Single = (NewVolume / 100)
                            gSessions(i).SimpleAudioVolume.Volume = NewVolSingle

                            If gDebugFlag AndAlso (ApplicationName <> LastApplicationReported) Then
                                If (ApplicationName <> LastApplicationReported) Then
                                    Console_WriteLineInColour("[ The updated audio level for application '" & ApplicationName & "' = " & NewVolume & " ]", ConsoleColor.Cyan)
                                End If
                            End If

                        End If

                    End If

                    LastApplicationReported = ApplicationName

                Next

                GoTo AllGood

            End If

            If gSetAppAudioVolumeForSystemSounds Then

                If (CommandLine = "") AndAlso (Not gMuteFlag) AndAlso (Not gUnMuteFlag) Then
                    Console_WriteLineInColour("Error: a new audio level was not specified for Windows System sounds.", ConsoleColor.Red)
                    gErrorExplained = True
                    GoTo ErrorFound
                End If

                If (NewVolume < 0 OrElse NewVolume > 100) Then
                    Console_WriteLineInColour("Error: The audio level for Window's System sounds was not updated as the update value specified was not between 0 and 100 inclusive", ConsoleColor.Red)
                    gErrorExplained = True
                    GoTo ErrorFound
                End If

                If gMuteFlag OrElse gUnMuteFlag Then
                Else
                    Try
                        ' Create a new instance of the MMDeviceEnumerator
                        Dim enumerator As New MMDeviceEnumerator()

                        ' Get the default audio playback device
                        Dim defaultPlaybackDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)

                        ' Check if the default playback device is found
                        If defaultPlaybackDevice Is Nothing Then
                            Console.WriteLine("Default playback device not found.")
                            GoTo ErrorFound
                        End If

                        ' Get the audio session manager for the default playback device
                        Dim sessionManager = defaultPlaybackDevice.AudioSessionManager

                        ' Get the list of all active audio sessions
                        Dim sessions = sessionManager.Sessions

                        ' Find the session for system sounds
                        Dim SystemSoundsSession As AudioSessionControl = Nothing

                        For i As Integer = 0 To sessions.Count - 1
                            Dim session = sessions(i)
                            ' Check if the session is the system sounds session
                            If IsSystemSoundsSession(session) Then
                                SystemSoundsSession = session
                                Exit For
                            End If
                        Next

                        ' Check if the system sounds session is found
                        If SystemSoundsSession IsNot Nothing Then

                            Dim NewVolSingle As Single = (NewVolume / 100)
                            SystemSoundsSession.SimpleAudioVolume.Volume = NewVolSingle
                            CommandLine = ""
                        Else
                            Console.WriteLine("System sounds session not found")
                        End If

                    Catch ex As Exception
                        Console.WriteLine("An error occurred: " & ex.Message)
                    End Try

                End If

            End If

            If CommandLine.Length = 0 Then GoTo AllGood

            ' set master audio level

            If ((NewVolume >= 0) AndAlso (NewVolume <= 100)) Then
                Dim NewVolSingle As Single = NewVolume / 100
                SetTheLevel(NewVolSingle)
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


            ' absolute set from 0 to 100
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
                    SetTheLevel(NewVolSingle)
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

        If gMuteFlag OrElse gUnMuteFlag Then

            If gSetAppAudioVolumeForSystemSounds Then

                Dim enumerator As New MMDeviceEnumerator()
                Dim defaultPlaybackDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)
                Dim sessionManager = defaultPlaybackDevice.AudioSessionManager
                Dim sessions = sessionManager.Sessions
                Dim SystemSoundsSession As AudioSessionControl = Nothing
                For i As Integer = 0 To sessions.Count - 1
                    Dim session = sessions(i)
                    If IsSystemSoundsSession(session) Then

                        If gMuteFlag Then session.SimpleAudioVolume.Mute = True
                        If gUnMuteFlag Then session.SimpleAudioVolume.Mute = False

                        Exit For
                    End If
                Next

            Else

                If gUnMuteFlag Then gDev.AudioEndpointVolume.Mute = False
                If gMuteFlag Then gDev.AudioEndpointVolume.Mute = True

            End If

        End If

        Dim FinalVolume As Single

        If MasterVolumeLevelChanged Then
            FinalVolume = gDev.AudioEndpointVolume.MasterVolumeLevelScalar
        Else
            FinalVolume = gStartingLevel / 100
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

                Console.WriteLine("Master audio level = " & ReturnCode.ToString)
                Console.WriteLine("")

                For x As Integer = 0 To gDev.AudioEndpointVolume.Channels.Count - 1
                    Console.WriteLine(" Channel " & x & " level = " & Math.Round(gDev.AudioEndpointVolume.Channels(x).VolumeLevelScalar * 100))
                Next

            End If

        End If

Beep:

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

        EndingLevel = gDev.AudioEndpointVolume.MasterVolumeLevelScalar * 100

        If gDebugFlag Then

            Console.WriteLine("")

            Dim LastApplicationReported As String = ""

            If gSetAppAudioVolume Then
                For Each i In gSessionIndex
                    If ApplicationName <> LastApplicationReported Then
                        Console_WriteLineInColour("[ Ending audio level for application '" & ApplicationName & "' = " & gSessions(i).SimpleAudioVolume.Volume * 100 & " ]", ConsoleColor.Cyan)
                    End If
                    LastApplicationReported = ApplicationName
                Next
            End If

            Console_WriteLineInColour("[ Ending master level = " & EndingLevel & " ]", ConsoleColor.Cyan)
            Console_WriteLineInColour(" ", ConsoleColor.Cyan)
            Console_WriteLineInColour("[ Return code = " & ReturnCode & " ]", ConsoleColor.Cyan)

#If DEBUG Then
            Console.ReadLine()
#End If

        End If

WrapUp:

        If (gCurrentColour <> gStartingColour) Then
            Console_WriteLineInColour(" ", Console.ForegroundColor)
        End If

        Return ReturnCode

    End Function

    Private Sub SetTheLevel(ByVal NewLevel As Single)

        Try

            Dim StartingLevel As Single = gStartingLevel / 100

            If StartingLevel = NewLevel Then

                If gFadeTime > 0 Then
                    Console_WriteLineInColour("Note: the current level is already " & (NewLevel * 100) & ", it has been left unchanged.", ConsoleColor.Yellow)
                End If

                Exit Sub

            End If

            If gFadeTime < 0 Then

                'immediately set the new level

                If gSetAppAudioVolume Then
                    For Each i In gSessionIndex
                        gSessions(i).SimpleAudioVolume.Volume = NewLevel
                    Next
                Else
                    gDev.AudioEndpointVolume.MasterVolumeLevelScalar = NewLevel
                End If

            Else

                Dim OverallStopWatch As New Stopwatch

                If gDebugFlag Then
                    OverallStopWatch.Start()
                End If

                'set the new level in increments over a specified period of time

                Dim TotalTimeInMilliSeconds As Integer = gFadeTime * 1000
                Dim OptimalDelayTimeInMilliSeconds As Integer = TotalTimeInMilliSeconds / Math.Abs(Int((StartingLevel * 100) - (NewLevel * 100)))

                Dim Increment As Single
                If StartingLevel < NewLevel Then
                    Increment = 0.01  ' raise the level by 1/100th
                Else
                    Increment = -0.01 ' decrease the level by 1/100th
                End If

                If gFadeTime = 1 Then
                    Console.Write("Adjusting level from " & gStartingLevel & " to " & (NewLevel * 100) & " for the next 1 second.")
                Else
                    Console.Write("Adjusting level from " & gStartingLevel & " to " & (NewLevel * 100) & " for the next " & gFadeTime.ToString("#,##0.###") & " seconds.")
                End If

                Console.WriteLine(" Press [Ctrl][C] at any time to stop.")
                Console.ForegroundColor = gStartingColour

                Dim LoopProcessingStopWatch As New Stopwatch

                Dim NextDelayTime As Integer = OptimalDelayTimeInMilliSeconds

                For x As Single = StartingLevel + Increment To NewLevel Step Increment

                    If NextDelayTime > 0 Then System.Threading.Thread.Sleep(NextDelayTime)

                    LoopProcessingStopWatch.Reset()
                    LoopProcessingStopWatch.Start()

                    x = Math.Round(x, 2)
                    If x > 1 Then x = 1

                    If gSetAppAudioVolume Then
                        For Each i In gSessionIndex
                            gSessions(i).SimpleAudioVolume.Volume = x
                        Next
                    Else
                        gDev.AudioEndpointVolume.MasterVolumeLevelScalar = x
                    End If

                    Dim currentLineCursor As Integer = Console.CursorTop
                    Console.SetCursorPosition(0, Console.CursorTop)
                    Console.SetCursorPosition(0, currentLineCursor)
                    Console.Write("The current level is now " & x * 100 & "   ")

                    LoopProcessingStopWatch.Stop()

                    'The next delay time = the optimal delay time - the time it takes to change the audio level and update the display - a small fudge factor to account for calculating the delay time

                    NextDelayTime = OptimalDelayTimeInMilliSeconds - CInt(LoopProcessingStopWatch.ElapsedMilliseconds) - 4

                Next

                If gDebugFlag Then
                    OverallStopWatch.Stop()
                    Console.WriteLine("")
                    Console.WriteLine("")
                    Console_WriteLineInColour("[ Total time to complete was = " & (OverallStopWatch.ElapsedMilliseconds / 1000).ToString & " seconds ]", ConsoleColor.Cyan)
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Function InventoryDevices(ByVal CommandLine As String) As DeviceTableStructure()

        Dim ReturnValue() As DeviceTableStructure = Nothing

        Dim DefaultDeviceName_Record As String = String.Empty
        Dim DefaultDeviceName_Play As String = String.Empty
        Dim DefaultDeviceName_Comm_Record As String = String.Empty
        Dim DefaultDeviceName_Comm_Play As String = String.Empty

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

            Try
                DefaultDeviceName_Comm_Record = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications).FriendlyName
            Catch ex As Exception
            End Try

            Try
                DefaultDeviceName_Comm_Play = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Communications).FriendlyName
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
                        ReturnValue(Index).DefaultDeviceComm = (endpoint.FriendlyName.ToString = DefaultDeviceName_Comm_Record)

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
                        ReturnValue(Index).DefaultDevice = (endpoint.FriendlyName.ToString = DefaultDeviceName_Comm_Play)

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

            System.Threading.Thread.Sleep(10) ' 1/100th of a second

        End While

        If gDebugFlag Then
            Console_WriteLineInColour("[ Listen samples taken = " & sampleCount & " ]", ConsoleColor.Cyan)
        End If

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
        Console_WriteLineInColour("SetVol v4.4 Help")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("Options:")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" ?                 help")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" n                 set the master level to n where n = 0 to 100")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" +n                increase the master level by n where n = 1 to 100; maximum result = 100")
        Console_WriteLineInColour(" -n                decrease the master level by n where n = 1 to 100; minimum result = 0")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" x                 set the master to x where x = 'zero' to 'one hundred' (without the quotes)")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" align             make all channel levels equal the master level")
        Console_WriteLineInColour("                   most devices will have at least two channel levels")
        Console_WriteLineInColour("                   with the first channel representing the left speaker and the second channel representing the right speaker")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" appaudio app      used to set the audio level of a specific application/program or for SystemSounds")
        Console_WriteLineInColour("                   when used without an app name, apps that can have their audio levels changed will be listed")
        Console_WriteLineInColour("                   when used with an app name, this option should be proceeded by either a desired audio level, relative")
        Console_WriteLineInColour("                   audio level, mute, or unmute. For example:")
        Console_WriteLineInColour("                     75 appaudio Spotify")
        Console_WriteLineInColour("                    +10 appaudio Spotify")
        Console_WriteLineInColour("                     mute appaudio Spotify")
        Console_WriteLineInColour("                     unmute appaudio Spotify")
        Console_WriteLineInColour("                   SetVol can only change an app's audio level, it cannot change an app's recording level")
        Console_WriteLineInColour("                   use 'SetVol appaudio' for additional information, including a list of applications that can have their levels changed")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" balance c1:c2:cz  set the audio/recording device's channel levels")
        Console_WriteLineInColour("                   cl is the desired level of the first channel where c1 = 0 to 100")
        Console_WriteLineInColour("                   c2 is the desired level of the second channel where c2 = 0 to 100")
        Console_WriteLineInColour("                   ...")
        Console_WriteLineInColour("                   cz is the desired level of the last channel where cz = 0 to 100")
        Console_WriteLineInColour("                   values may be included for as many channels as the device supports")
        Console_WriteLineInColour("                   for example:")
        Console_WriteLineInColour("                     50 balance 100:100")
        Console_WriteLineInColour("                        set the first and second channel (usually the left and right speakers) levels to 50")
        Console_WriteLineInColour("                     100 balance 50:50")
        Console_WriteLineInColour("                        also set the first and second channel levels to 50")
        Console_WriteLineInColour("                     100 balance 80:100:70:100:90:85:50:100")
        Console_WriteLineInColour("                        set the eight channel levels to the values specified")
        Console_WriteLineInColour("                     50 balance 80:40")
        Console_WriteLineInColour("                        set the first channel level to 40 (50 x 80%) and the second channel level 20 (50 x 40%)")
        Console_WriteLineInColour("                   when the balance option is used the greatest of the channel levels becomes the new master level")
        Console_WriteLineInColour("                   in the examples above the master level was specified")
        Console_WriteLineInColour("                   if the master level is not specified the current master level value will be used")
        Console_WriteLineInColour("                   and the master and channel levels be updated as in the last example above")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" beep              play a beep")
        Console_WriteLineInColour("                   if any changes to levels are being made, they will be made before the beep is played")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" debug             show some additional displays")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" device name       used to set the audio/recording level of a specific device")
        Console_WriteLineInColour("                   when not used, SetVol will act upon the default audio device")
        Console_WriteLineInColour("                   when used with a valid device name, SetVol will act upon the named device")
        Console_WriteLineInColour("                   when used without a valid device name, all audio and recording devices will be displayed")
        Console_WriteLineInColour("                   when the device option is used it must be the last option specified on the command line")
        Console_WriteLineInColour("                   note: the device and appaudio options cannot be used in the same command line")
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
        Console_WriteLineInColour(" makedefault       change the default audio device for playing/recording")
        Console_WriteLineInColour("                   if the makedefault option is used, then the device option must be used too")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" makedefaultcomm   change the default communications device for playing/recording")
        Console_WriteLineInColour("                   when the makedefaultcomm option is used the device option must be used too")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("                   note: makedefault and makedefaultcomm may not be used in the same command line")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" mute              mute a device or application")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" unmute            unmute a device or application")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" nodefault         a pop-up error message will be displayed if a device is not specified and there is no default audio device")
        Console_WriteLineInColour("                   the use of this option suppresses that message")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" over n            evenly transition to the level being set over n seconds")
        Console_WriteLineInColour("                   must be greater than zero but less than or equal to 86,400 (one day)")
        Console_WriteLineInColour("                   when the current level does not need to be changed the program will exit without waiting the specified time")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" report            displays a device's master and channel levels")
        Console_WriteLineInColour("                   the master level will be provided via SetVol's return code")
        Console_WriteLineInColour("                     for example: if the device's master level is 60, SetVol's return code will be 60")
        Console_WriteLineInColour("                   when the device is muted the return code will be a negative number")
        Console_WriteLineInColour("                    for example: if the device's master level is 60 and the device is muted, SetVol's return code will be -60")
        Console_WriteLineInColour("                   when a level is changed, the changed level is reported")
        Console_WriteLineInColour("                   this option may also be used with the device option")
        Console_WriteLineInColour("                   when the device option is not used, the default audio device will be used")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour(" website           visit the SetVol website")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("Examples:")
        Console_WriteLineInColour(" setvol 55")
        Console_WriteLineInColour(" setvol +10")
        Console_WriteLineInColour(" setvol seventy-five")
        Console_WriteLineInColour(" setvol 60 over 30")
        Console_WriteLineInColour(" setvol 50 balance 80:100")
        Console_WriteLineInColour(" setvol mute")
        Console_WriteLineInColour(" setvol appaudio")
        Console_WriteLineInColour(" setvol 35 appaudio Spotify")
        Console_WriteLineInColour(" setvol 35 appaudio Spotify debug")
        Console_WriteLineInColour(" setvol 20 appaudio SystemSounds")
        Console_WriteLineInColour(" setvol device")
        Console_WriteLineInColour(" setvol 75 device Speakers (Realtek USB2.0 Audio))")
        Console_WriteLineInColour(" setvol report")
        Console_WriteLineInColour(" setvol report device Microphone (Yeti Stereo Microphone)")
        Console_WriteLineInColour(" setvol listen 2.5 device Speakers (Realtek USB2.0 Audio)")
        Console_WriteLineInColour(" setvol makedefault device Speakers (Realtek USB2.0 Audio)")
        Console_WriteLineInColour(" setvol makedefaultcomm device Microphone (Yeti Stereo Microphone)")
        Console_WriteLineInColour(" setvol website")
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol v4.4", ConsoleColor.Yellow)
        Console_WriteLineInColour("Copyright © 2025, Rob Latour", ConsoleColor.Yellow, True)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol is open source", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://github.com/roblatour/setvol", ConsoleColor.Cyan)
        Console_WriteLineInColour(" ")
        Console_WriteLineInColour("SetVol is licensed under the MIT License", ConsoleColor.Cyan)
        Console_WriteLineInColour("https://github.com/roblatour/setvol/blob/main/LICENSE", ConsoleColor.Cyan)
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
    Private Function ConvertNumberToWords(ByVal input As Integer) As String

        Select Case input
            Case 0
                Return "zero"
            Case 1
                Return "one"
            Case 2
                Return "two"
            Case 3
                Return "three"
            Case 4
                Return "four"
            Case 5
                Return "five"
            Case 6
                Return "six"
            Case 7
                Return "seven"
            Case 8
                Return "eight"
            Case 9
                Return "nine"
            Case 10
                Return "ten"
            Case Else
                Return input.ToString

        End Select

    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Private Function ConvertFromWords(ByVal input As String) As Integer

        Try

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

            Dim TestNumber As Integer = CInt(TestCase)
            If TestNumber >= 0 And TestNumber <= 100 Then
                Return TestNumber
            End If

        Catch ex As Exception

        End Try

        Return -1

    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub Console_WriteLineInColour(ByVal Message As String, Optional ByVal Colour As ConsoleColor = ConsoleColor.White, Optional PrintSpecialCharacters As Boolean = False)

        Dim OriginalTextEncoding As System.Text.Encoding = System.Console.OutputEncoding

        Try

            If PrintSpecialCharacters Then
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