Imports System.IO

Public Class SOUND
    Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Private oName As String = Nothing

    Public Property Name As String
        Set(value As String)
            oName = value
        End Set
        Get
            Return oName
        End Get
    End Property

    Public Sub Play(ByVal id As Integer, ByVal repeat As Boolean, Optional vol As Integer = 1000)
        If repeat = True Then
            mciSendString("Open " & GetFile(id) & " alias " & oName, CStr(0), 0, 0)
            mciSendString("Play " & oName & " repeat", CStr(0), 0, 0)
        Else
            mciSendString("Open " & GetFile(id) & " alias " & oName, CStr(0), 0, 0)
            mciSendString("Play " & oName, CStr(0), 0, 0)
        End If
        'Optionally Set Volume
        mciSendString("setaudio" & oName & " volume to " & vol, CStr(0), 0, 0)
    End Sub

    'Media Library
    Private Function GetFile(ByVal Id As Integer) As String
        Dim path1 As String = ""

        ' Spaces cause failure to load (Add quotes)
        ' Dots in path can failure to load C:\this.is.my.folder\song.wav 
        ' Very Long Paths will cause failure to load (255)
        ' .wav files will fail to load if repeat = true
        Select Case Id
            'Case 0 'BackGroundSound
                'path1 = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(CurDir))) & "\Resources\backgroundmusic.mp3"
            Case 1 'Get Coin
                path1 = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(CurDir))) & "\Resources\coin.wav"
            Case 2 'Hit by Bomb
                path1 = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(CurDir))) & "\Resources\bomb.mp3"
            Case 3 'Button clicked
                path1 = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(CurDir))) & "\Resources\buttonclicked.wav"
            Case 4 'Dice sound
                path1 = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(CurDir))) & "\Resources\dice.mp3"
            Case 5 'coin drop sound
                path1 = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(CurDir))) & "\Resources\coindrop.wav"
        End Select

        path1 = Chr(34) & path1 & Chr(34)

        Return path1
    End Function
End Class
