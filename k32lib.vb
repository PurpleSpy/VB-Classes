Public Class Kernel32
    Public Class MusicMaker
        Inherits Collections.ArrayList
        Public Property playing As Boolean = False

        Enum BeepTones
            c = 261.6
            d = 293.7
            e = 329.6
            f = 349.2
            g = 392
            a = 440
            b = 493.9
            silence = 0
        End Enum

        Public Sub addnote(note As MusicNote)
            Me.Add(note)
        End Sub
        Public Sub addnote(note As MusicNote, addat As Integer)
            Me.Insert(addat, note)
        End Sub




        Public Sub addnote(note As MusicMaker.BeepTones, currentoctave As Integer, notelength As Integer)
            addnote(New MusicNote(note, currentoctave, notelength))
        End Sub

        Public Sub addnote(note As MusicMaker.BeepTones, currentoctave As Integer, notelength As Integer, addat As Integer)
            addnote(New MusicNote(note, currentoctave, notelength), addat)
        End Sub

        Public Sub removenoteatindex(index As Integer)
            Me.RemoveAt(index)
        End Sub
        Public Sub removenote(note As MusicNote)
            Me.removenote(note)
        End Sub

        Public Sub stopPlaying()
            playing = False
        End Sub

        Public Sub startPlayingAsnc()
            Dim csg As New Threading.Thread(AddressOf startPlaying)
            csg.Start()

        End Sub




        Public Sub startPlaying()
            If playing Then
                Throw New Exception("Device already in use")
            End If

            playing = True
            RaiseEvent PlaystateChanged(Me, playing)
            Dim cstx As IEnumerator = Me.GetEnumerator

            While cstx.MoveNext And playing
                Dim ctx As MusicNote = cstx.Current

                Select Case ctx.tone
                    Case BeepTones.silence
                        Threading.Thread.Sleep(ctx.milliseconds)
                    Case Else
                        Dim tonefreq As Double = ctx.tone * 1000
                        Dim baseoctave As Double = 4
                        Dim octaverange As Integer() = {3, 2, 1, 0, 1, 2, 3, 4}

                        If ctx.octave <> 4 Then
                            If ctx.octave > 4 Then
                                tonefreq *= Math.Pow(2, octaverange(ctx.octave - 1))
                            Else
                                tonefreq /= Math.Pow(2, octaverange(ctx.octave - 1))
                            End If
                        End If

                        Beep(tonefreq, ctx.milliseconds)
                End Select

            End While

            playing = False
            RaiseEvent PlaystateChanged(Me, playing)
        End Sub

        Public Event PlaystateChanged(Sender As Object, playstate As Boolean)
        Public Class MusicNote
            Public tone As MusicMaker.BeepTones
            Public octave As Integer = 4
            Public milliseconds As Integer
            Sub New(note As MusicMaker.BeepTones, currentoctave As Integer, notelength As Integer)
                If currentoctave < 1 Or currentoctave > 8 Then
                    Throw New ArgumentException("Octave must be between 1 and 8")
                End If

                Me.tone = note
                octave = currentoctave
                milliseconds = notelength

            End Sub
        End Class

        Public Class Song
            Public Property voices As New ArrayList

            Public Sub createVoices(ByVal numberofvoices As Integer)

                For i As Integer = 0 To numberofvoices
                    voices.Add(New MusicMaker)
                Next

            End Sub

            Public Sub addVoice(ByRef Voice As MusicMaker)
                voices.Add(Voice)
            End Sub
            Public Sub removeVoice(ByRef Voice As MusicMaker)
                voices.Remove(Voice)
            End Sub

            Public Sub removeVoice(ByVal removeindex As Integer)
                voices.RemoveAt(removeindex)
            End Sub

            Public Sub stopsong()
                For Each voice As MusicMaker In voices
                    voice.playing = False
                Next

            End Sub

            Public Sub playsong()
                For Each voice As MusicMaker In voices
                    voice.startPlayingAsnc()
                Next
            End Sub
        End Class
    End Class


    Public Structure powerstatusdata
        Public AClineStatus As Byte
        Public BatteryFlag As Byte
        Public BatteryLifePercent As Byte
        Public SystemStatusFlag As Byte
        Public BatteryLifeTime As Integer
        Public BatteryFullLifeTime As Integer

        Enum acstatusflag
            Offline = 0
            Online = 1
            Unknown = 255
        End Enum

        Enum BatteryFlagInfo
            High = 1
            Low = 2
            Critical = 4
            Charging = 8
            NoBattery = 128
            Unknown = 255
        End Enum

        Enum SystemStatusFlagInfo
            BatterySaverOff = 0
            BatterySaverOn = 1
        End Enum

        Public Function getBatteryStatus() As String
            Return System.Enum.GetName(GetType(BatteryFlagInfo), BatteryFlag)
        End Function

        Public Function getACStatus() As String
            Return System.Enum.GetName(GetType(acstatusflag), AClineStatus)
        End Function

        Function getBatteryPercent() As Double
            If BatteryLifePercent = 255 Then
                Return -1
            Else
                Return BatteryLifePercent / 100
            End If
        End Function
    End Structure

    <Runtime.InteropServices.DllImport("Kernel32")> Shared Function GetSystemPowerStatus(ByRef dinfo As powerstatusdata) As Boolean

    End Function
    <Runtime.InteropServices.DllImport("Kernel32")> Shared Function Beep(tone As Integer, millseconds As Integer) As Boolean

    End Function


End Class