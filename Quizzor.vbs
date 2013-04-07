' Quizzor - A MediaMonkey plugin for performing music quizzes
' Copyright (C) 2013 "m42i" 
' 
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

' Monkey Media Quizzor plugin
' Features:
' O Only show track info if space bar is pressed
' X Stop after current track
' O Play track again if length < 60s
' O Keep track of correctly guessed tracks

Const DEBUG_ON = True
Const BTN_MARGIN = 5 ' Defines the standard margin between buttons
Const BTN_HEIGHT = 25 ' Defines standard height of a button
Const BTN_WIDTH = 80 ' Defines standard width of a button
Const TIME_WIDTH = 50 ' Defines standard width of a time label

Const HTML_Style = "<style type='text/css'>body { overflow: auto; } </style>"

Const akLeft = 1
Const akTop = 2
Const akRight = 4
Const akBottom = 8

Const alNone = 0
Const alTop = 1
Const alBottom = 2
Const alLeft = 3
Const alRight = 4
Const alClient = 5

' Keep track of current quiz playlist
Dim Quiz_Playlist

' Keep track of important control elements
Dim QuizzorMainPanel, SongTrackBar, SongTimer
Dim SongTime, SongTimeLeft ' Labels for current song time
Dim CurrentSongLength

' Keep track of playlists between sessions
' [SectionName]
' key = description
' 
' [Quizzor]
' LastPlaylistID = Playlist.ID As Long
' NowPlayingSongs_Playlist.ID = "SongData.ID,SongData.ID,..." As String
'
Dim OptionsFile

Function GetFormattedDate()
    Today = Date
    This_Year = Year(Today)
    This_Month = Month(Today)
    If This_Month < 10 Then This_Month = "0" + CStr(This_Month) 
    Dim This_Day : This_Day = Day(Today)
    If This_Day < 10 Then This_Day = "0" + CStr(This_Day) 

    GetFormattedDate = CStr(This_Year) + "-" + CStr(This_Month) + "-" + CStr(This_Day)
End Function

' Formats a given time in seconds int mm:ss
Function GetFormattedTime(Time)
    Minutes = Int(Time / 60)
    If Minutes < 10 Then Minutes = "0" + CStr(Minutes)
    Seconds = Int(Time Mod 60)
    If Seconds < 10 Then Seconds = "0" + CStr(Seconds)

    GetFormattedTime = CStr(Minutes) + ":" + CStr(Seconds)
End Function

Sub DebugOutput(msg)
    SDB.MessageBox msg, mtInformation, Array(mbOk)
End Sub

' Thanks to Diddeleedoo from the MM forums
Sub RandomizePlaylist
    song_count = SDB.Player.CurrentPlaylist.Count
    If song_count <2 Then Exit Sub : SDB.Player.isShuffle = False
    If SDB.Player.isPlaying Or SDB.Player.isPaused Then
        Shuffle song_count 
        If Not SDB.Player.CurrentSongIndex = -1 Then
            SDB.Player.PlaylistMoveTrack _
            SDB.Player.CurrentSongIndex,0
            SDB.Player.CurrentSongIndex = 0
        End If
    Else
        Shuffle song_count : SDB.Player.CurrentSongIndex=0
    End If
End Sub

' Get a comma seperated string list of all IDs in a given SongList
Function GetSongIDList(SongList)
    If SongList.Count = 0 Then GetSongIDList = ""

    Dim Result : Result = CStr(SongList.Item(0).ID)
    For i = 1 To SongList.Count - 1
        Result = Result + "," + CStr(SongList.Item(i).ID)
    Next

    GetSongIDList = Result
End Function

' Create a new empty quiz playlist and prevent duplicates
Function CreateNewPlaylist()
    Dim NewBaseTitle : NewBaseTitle = SDB.Localize("Quiz of " + GetFormattedDate())
    
    ' If a playlist with that name doesn't exist, root is returned
    Set Playlist_Root = SDB.PlaylistByTitle(NewBaseTitle)

    Dim i : i = 1
    Dim NewTitle : NewTitle = NewBaseTitle
    While Not Playlist_Root.Title = ""
        NewTitle = NewBaseTitle + " (" + CStr(i) + ")"
        Set Playlist_Root = SDB.PlaylistByTitle(NewTitle)
        i = i + 1
    WEnd

    Set CreateNewPlaylist = Playlist_Root.CreateChildPlaylist(NewTitle)
End Function

Sub Shuffle(n)
    Randomize
    j = n - 1
    For i = 0 To n - 1
        SDB.Player.PlaylistMoveTrack i,Int(n*Rnd)
        SDB.Player.PlaylistMoveTrack j,Int(n*Rnd)
        j = j - 1
    Next
End Sub

' Check whether a quiz can be started and show a message box if not
Function IsQuizReady()
    Dim IsReady : IsReady = True
    
    If SDB.AllVisibleSongList.Count = 0 Then
        IsReady = False

        SDB.MessageBox SDB.Localize("No songs visible. Please fill the main window first."),_
        mtInformation, Array(mbOk)
    ElseIf IsObject(Quiz_Playlist) Then
        IsReady = False

        SDB.MessageBox SDB.Localize("Please create a new quiz first."), mtInformation, Array(mbOk)
    End If

    IsQuizReady = IsReady
End Function

Sub DestroyAllObjects
    Set PlayBtn = SDB.Objects("PlayBtn")
    If Not (PlayBtn Is Nothing) Then
        PlayBtn.Common.Visible = False
        PlayBtn = Nothing
    End If

    Set NextBtn = SDB.Objects("NextBtn")
    If Not (NextBtn Is Nothing) Then
        NextBtn.Common.Visible = False
        NextBtn = Nothing
    End If

    Set SongInfoLabel = SDB.Objects("SongInfoLabel")
    If Not (SongInfoLabel Is Nothing) Then
        SongInfoLabel.Common.Visible = False
        SongInfoLabel = Nothing
    End If

    Set ShowInfoBtn = SDB.Objects("ShowInfoBtn")
    If Not (ShowInfoBtn Is Nothing) Then
        ShowInfoBtn.Common.Visible = False
        ShowInfoBtn = Nothing
    End If

    Set QuizzorMainPanel = SDB.Objects("QuizzorMainPanel")
    If IsObject(QuizzorMainPanel) And Not (QuizzorMainPanel Is Nothing) Then
        QuizzorMainPanel.Common.Visible = False
        Set SDB.Objects("QuizzorMainPanel") = Nothing
    End If
    
    Set SDB.Objects("QuizBar") = Nothing
    Set SDB.Objects("NewQuizBtn") = Nothing
    Set SDB.Objects("StartQuizBtn") = Nothing
    Set SDB.Objects("StopQuizBtn") = Nothing

    Script.UnRegisterEvents SDB
End Sub

' Clears SongInfoHTML end ensures the style is preserved
Sub ClearSongInfoHTML
    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    Set HTMLDocument = SongInfoHTML.Interf.Document
    HTMLDocument.Write "<html>" & vbCrLf & HTML_Style & vbCrLf & "<body>&nbsp;</body></html>"
    HTMLDocument.Close
End Sub

Function GetSongInfoHTML(SongData)
    GetSongInfoHTML = "<html>" & vbCrLf & HTML_Style & vbCrLf & _
        "<table border='1' cellspacing='0' cellpaddin='2' rules='rows'" & _
        " frame='void' width='100%' height='100%'>" & vbCrLf & _
        "<colgroup>" & vbCrLf & _
            "<col width='20%'>" & vbCrLf & _
            "<col width='80%'>" & vbCrLf & _
        "</colgroup>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td align='right'>" & SDB.Localize("Album") & "</td>" & vbCrLf & _
            "<td>&nbsp;" & SongData.AlbumName & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td align='right'>" & SDB.Localize("Title") & "</td>" & vbCrLf & _
            "<td>&nbsp;" & SongData.Title & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td align='right'>" & SDB.Localize("Artist") & "</td>" & vbCrLf & _
            "<td>&nbsp;" & SongData.ArtistName & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td align='right'>" & SDB.Localize("Comment") & "</td>" & vbCrLf & _
            "<td>&nbsp;" & SongData.Comment & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td align='right'>" & SDB.Localize("File") & "</td>" & vbCrLf & _
            "<td>&nbsp;" & SongData.Path & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "</body></html>"
End Function

Sub CreateMainPanel()
    Set UI = SDB.UI
    
    Set QuizzorMainPanel = UI.NewDockablePersistentPanel("QuizzorMainPanel")
    If QuizzorMainPanel.IsNew Then
        QuizzorMainPanel.DockedTo = 4 
    End If
    ' Show panel to ensure the size and position of related elements
    QuizzorMainPanel.Common.Visible = True
    QuizzorMainPanel.ShowCaption = False

    Set PlayBtn = UI.NewButton(QuizzorMainPanel)
    PlayBtn.Common.ControlName = "PlayBtn"
    PlayBtn.Caption = SDB.Localize("Play")
    Script.RegisterEvent PlayBtn, "OnClick", "StartPlaying"

    Set PauseBtn = UI.NewButton(QuizzorMainPanel)
    PauseBtn.Common.ControlName = "PauseBtn"
    PauseBtn.Caption = SDB.Localize("Pause")
    Script.RegisterEvent PauseBtn, "OnClick", "PausePlayback"

    Set NextBtn = UI.NewButton(QuizzorMainPanel)
    NextBtn.Common.ControlName = "NextBtn"
    NextBtn.Caption = SDB.Localize("Next")
    Script.RegisterEvent NextBtn, "OnClick", "PlayNext"

    Set ShowInfoBtn = UI.NewButton(QuizzorMainPanel)
    ShowInfoBtn.Common.ControlName = "ShowInfoBtn"
    ShowInfoBtn.Caption = SDB.Localize("Show Information")
    Script.RegisterEvent ShowInfoBtn, "OnClick", "ShowSongInfo"

    Set SongInfoHTML = UI.NewActiveX(QuizzorMainPanel, "Shell.Explorer")
    SongInfoHTML.Common.ControlName = "SongInfoHTML"
    SongInfoHTML.Common.Align = 0
    SongInfoHTML.Common.Anchors = akLeft + akTop + akRight + akBottom
    SongInfoHTML.Interf.Navigate "about:" ' A trick to make sure document exists, from Wiki

    Set SongTime = UI.NewLabel(QuizzorMainPanel)
    SongTime.Common.ControlName = "SongTime"
    SongTime.Common.Anchors = akLeft + akBottom
    SongTime.Alignment = 2 ' Center
    SongTime.Autosize = False
    SongTime.Caption = "00:00"

    ' TODO: Change playback time, when TrackBar changes
    Set SongTrackBar = UI.NewTrackBar(QuizzorMainPanel)
    SongTrackBar.Common.ControlName = "SongTrackBar"
    SongTrackBar.Common.Anchors = akBottom 
    SongTrackBar.Common.Enabled = False
    SongTrackBar.Common.Anchors = akLeft + akBottom + akRight
    SongTrackBar.Value = 0
    SongTrackBar.Horizontal = True

    Set SongTimeLeft = UI.NewLabel(QuizzorMainPanel)
    SongTimeLeft.Common.ControlName = "SongTimeLeft"
    SongTimeLeft.Common.Anchors = akBottom + akRight
    SongTimeLeft.Alignment = 2 ' Center
    SongTimeLeft.Autosize = False
    SongTimeLeft.Caption = "00:00"

    Call ResizeMainPanel

    ' Always hide Main Panel if it is created
    QuizzorMainPanel.Common.Visible = False
End Sub

Sub ResizeMainPanel
    If Not IsObject(QuizzorMainPanel) Then
        Call CreateMainPanel
    End If

    Set PlayBtn = QuizzorMainPanel.Common.ChildControl("PlayBtn")
    PlayBtn.Common.SetRect BTN_MARGIN, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT

    Set PauseBtn = QuizzorMainPanel.Common.ChildControl("PauseBtn")
    PauseBtn.Common.SetRect 2*BTN_MARGIN + BTN_WIDTH,BTN_MARGIN, _
        BTN_WIDTH, BTN_HEIGHT

    Set NextBtn = QuizzorMainPanel.Common.ChildControl("NextBtn")
    NextBtn.Common.SetRect 3*BTN_MARGIN + 2*BTN_WIDTH, BTN_MARGIN, _
        BTN_WIDTH, BTN_HEIGHT

    Set ShowInfoBtn = QuizzorMainPanel.Common.ChildControl("ShowInfoBtn")
    ShowInfoBtn.Common.SetRect 4*BTN_MARGIN+3*BTN_WIDTH, BTN_MARGIN, _
        2*BTN_WIDTH + BTN_MARGIN, BTN_HEIGHT

    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    SongInfoHTML.Common.SetClientRect BTN_MARGIN, 2*BTN_MARGIN + BTN_HEIGHT, _
        QuizzorMainPanel.Common.Width - 3*BTN_MARGIN, _
        QuizzorMainPanel.Common.Height - 5*BTN_MARGIN - 2*BTN_HEIGHT

    Set SongTime = QuizzorMainPanel.Common.ChildControl("SongTime")
    SongTime.Common.SetRect 2*BTN_MARGIN, _
        QuizzorMainPanel.Common.Height - BTN_HEIGHT - BTN_MARGIN,_
        TIME_WIDTH, BTN_HEIGHT

    Set SongTrackBar = QuizzorMainPanel.Common.ChildControl("SongTrackBar")
    SongTrackBar.Common.SetRect BTN_MARGIN + TIME_WIDTH, _
        QuizzorMainPanel.Common.Height - BTN_HEIGHT - 3*BTN_MARGIN,_
        QuizzorMainPanel.Common.Width - 2*TIME_WIDTH - 2*BTN_MARGIN, BTN_HEIGHT

    Set SongTimeLeft = QuizzorMainPanel.Common.ChildControl("SongTimeLeft")
    SongTimeLeft.Common.SetRect BTN_MARGIN + TIME_WIDTH + SongTrackBar.Common.Width, _
        QuizzorMainPanel.Common.Height - BTN_HEIGHT - BTN_MARGIN,_
        TIME_WIDTH, BTN_HEIGHT
End Sub

' Open the playlists node
Sub SelectPlaylist(Playlist)
    Set Root = SDB.MainTree
    Set ParentPlaylistNode = Root.Node_Playlists
    ParentPlaylistNode.Expanded = True
End Sub

' Check whether the saved playlists already exist and delete if not
' This should be called whenever a playlist value is read
Sub UpdateOptionsFile
    If OptionsFile.ValueExists("Quizzor", "LastPlaylistID") Then
        Set Playlist = SDB.PlaylistByID(OptionsFile.IntValue("Quizzor", "LastPlaylistID")) 
        ' If no playlist exists, root (ID=0) is returned
        DebugOutput "Delete Playlist.ID = " + CStr(Playlist.ID)
        If Playlist.ID = 0 Then
            OptionsFile.DeleteKey "Quizzor", "LastPlaylistID"
        End If
    End If
    
    ' go through all saved playlists and check if they exist
    Set KeyList = OptionsFile.Keys("Quizzor")
    For i = 0 To KeyList.Count - 1
        KeyValue = KeyList.Item(i)
        ' Keys method returns each string as "key=value"
        Key = Left(KeyValue, InStr(KeyValue, "=") - 1)
        Value = Right(KeyValue, InStr(KeyValue, "=") + 1)
        DebugOutput Value
        IDPosition = InStrRev(Key, "_")
        If IDPosition > 0 Then
            ID = Mid(Key, IDPosition + 1)
            If SDB.PlaylistByID(ID).ID = 0 Then
                OptionsFile.DeleteKey "Quizzor", Key 
            End If
        End If
    Next

    OptionsFile.Flush
End Sub

' Creates a new quiz. 
' If a last session exists, the user is asked to use that.
' If not, a new quiz is created.
' Cancelling the dialog will change nothing.
Sub NewQuiz(Item)
    Call UpdateOptionsFile
    
    Dim NewQuizDialogAnswer
    If OptionsFile.ValueExists("Quizzor", "LastPlaylistID") Then
        NewQuizDialogAnswer = SDB.MessageBox( _
            SDB.Localize("Do you want to restore the last quiz session?") + vbCrLf + _
            SDB.Localize("Clicking No will create a new quiz. Click cancel to do nothing."), _ 
            mtWarning, Array(mbCancel, mbNo, mbYes))

        If NewQuizDialogAnswer = mrCancel Then
            Exit Sub
        ElseIf NewQuizDialogAnswer = mrYes Then
            Call RestoreLastSession
        ElseIf NewQuizDialogAnswer = mrNo Then
            Call CreateNewQuiz
        End If
    Else
        NewQuizDialogAnswer = SDB.MessageBox( SDB.Localize("Creating a new quiz replaces all  tracks") _
            + SDB.Localize(" in the current queue. This cannot be undone.") + vbCrLF _
            + SDB.Localize("Do you want to create a new quiz and lose the current queue?"), _
            mtWarning, Array(mbNo, mbYes))

        If NewQuizDialogAnswer = mrYes Then
            Call CreateNewQuiz
        ElseIf NewQuizDialogAnswer = mrNo Then
            Exit Sub
        End If
    End If

    ' Check whether a quiz playlist was created
    If IsObject(Quiz_Playlist) Then
        SDB.Objects("StartQuizBtn").Enabled = True
        SDB.Objects("StopQuizBtn").Enabled = True

        ' TODO: Automatically select newly created playlist 
        ' TODO: Automatically hide Now Playing List
        SDB.MessageBox SDB.Localize("Your new quiz playlist is ") + _
            Quiz_Playlist.Title + "." + vbCrLf _
            + SDB.Localize("Select it to see already played tracks."), _
            mtInformation, Array(mbOk)

        Call StartQuiz(Item)
    End If
End Sub

' Creates a new quiz without asking
Sub CreateNewQuiz
    If Not IsQuizReady Then Exit Sub

    ' Replace playing queue with current tracks from main window 
    Call SDB.Player.PlaylistClear()
    SDB.Player.PlaylistAddTracks SDB.AllVisibleSongList
    Call RandomizePlaylist

    ' Create new empty playlist, for played tracks
    Set Quiz_Playlist = CreateNewPlaylist()
    Call SelectPlaylist(Quiz_Playlist)

    ' Save new playlist data to ini file
    OptionsFile.IntValue("Quizzor", "LastPlaylistID") = Quiz_Playlist.ID
    OptionsFile.StringValue("Quizzor", "NowPlayingSongs_" + CStr(Quiz_Playlist.ID)) = _
        GetSongIDList(SDB.Player.CurrentSongList)
End Sub

Sub StartQuiz(Item)
    QuizzorMainPanel.Common.Visible = True
    ' Ensure that the elements are redrawn
    Call ResizeMainPanel
    
    SongTime.Caption = GetFormattedTime(0)
    SongTimeLeft.Caption = GetFormattedTime(0)

    SDB.Player.CurrentSongIndex = 0
End Sub

Sub StopQuiz(Item)
    If SDB.Player.isPlaying And IsObject(SongTimer) Then
        SongTimer.Enabled = False
        Script.UnRegisterEvents SongTimer
    End If

    QuizzorMainPanel.Common.Visible = False
    SDB.Player.Stop
    
    SongTime.Caption = GetFormattedTime(0)
    SongTimeLeft.Caption = GetFormattedTime(0)
    
    SDB.Objects("StartQuizBtn").Enabled = False
    SDB.Objects("StopQuizBtn").Enabled = False
End Sub

Sub StartPlaying
    If Not IsQuizReady() Then Exit Sub
    
    ' If the player is paused, just continue playing.
    If SDB.Player.isPaused Then
        Call SDB.Player.Pause
        Exit Sub
    End If

    SDB.Player.CurrentSongIndex = 0

    CurrentSongLength = SDB.Player.CurrentSong.SongLength / 1000
    SongTrackBar.MinValue = 0
    SongTrackBar.MaxValue = CurrentSongLength
    SongTrackBar.Value = 0
    
    SongTime.Caption = GetFormattedTime(0)
    SongTimeLeft.Caption = "- " + GetFormattedTime(CurrentSongLength)
    
    Set SongTimer = SDB.CreateTimer(100)
    Script.RegisterEvent SongTimer, "OnTimer", "UpdateSongTime"

    ' Disable playing next title
    ' Always play from the beginning
    SDB.Player.PlaybackTime = 0
    SDB.Player.Play
    SDB.Player.StopAfterCurrent = True
End Sub

' Pause and unpause playback
Sub PausePlayback
    Call SDB.Player.Pause
End Sub

Sub PlayNext
    If Not IsQuizReady() Then Exit Sub

    Call ClearSongInfoHTML

    If SDB.Player.CurrentPlaylist.Count <= 0 Then
        SDB.MessageBox SDB.Localize("Quiz has ended. Please create a new one."), _
            mtInformation, Array(mbOk)
        Call StopQuiz(Nothing)
        Exit Sub
    End If

    Quiz_Playlist.addTrack SDB.Player.PlaylistItems(0)
    SDB.Player.PlaylistDelete 0

    OptionsFile.StringValue("Quizzor", "NowPlayingSongs_" + CStr(Quiz_Playlist.ID)) = _
        GetSongIDList(SDB.Player.CurrentSongList)

    Call StartPlaying
End Sub

Sub ShowSongInfo
    Set CurrentSong = SDB.Player.CurrentSong
    If Not IsObject(CurrentSong) Then Exit Sub

    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    Set HTMLDocument = SongInfoHTML.Interf.Document
    HTMLDocument.Write GetSongInfoHTML(CurrentSong)
    HTMLDocument.Close
End Sub

Sub UpdateSongTime(Timer)
    PlaybackTime = SDB.Player.PlaybackTime / 1000
    SongTrackBar.Value = PlaybackTime
    SongTime.Caption = GetFormattedTime(PlaybackTime)
    SongTimeLeft.Caption = "- " + GetFormattedTime(CurrentSongLength - PlaybackTime)

    ' Update again in 100 ms
    Set SongTimer = SDB.CreateTimer(100)
End Sub

' Restores the last session
' Doesn't check if one exists and will not ask for permission
Sub RestoreLastSession
    Call SDB.Player.PlaylistClear()

    LastPlaylistID = OptionsFile.IntValue("Quizzor", "LastPlaylistID")
    Set Quiz_Playlist = SDB.PlaylistByID(LastPlaylistID)

    ' Fill Now Playing List
    ' TODO: Restore playlist order
    SongIDList = OptionsFile.StringValue("Quizzor", "NowPlayingSongs_" + CStr(LastPlaylistID))
    Set SongIter = SDB.Database.QuerySongs("ID in (" + SongIDList + ")")
    While Not SongIter.EOF
        SDB.Player.PlaylistAddTrack(SongIter.Item)
        Call SongIter.Next
    WEnd
End Sub

Sub OnStartup
    Set UI = SDB.UI
    
    ' Register new or get existing toolbar 
    Set QuizBar = SDB.Objects("QuizBar") 
    If QuizBar Is Nothing Then
        Set QuizBar = UI.AddToolbar("QuizBar")
        Set SDB.Objects("QuizBar") = QuizBar
    End If
       
    Set NewQuizBtn = SDB.Objects("NewQuizBtn")
    If NewQuizBtn Is Nothing Then
        Set NewQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
        NewQuizBtn.Caption = "New Quiz"
        Set SDB.Objects("NewQuizBtn") = NewQuizBtn  
    End If
       
    Set StartQuizBtn = SDB.Objects("StartQuizBtn")
    If StartQuizBtn Is Nothing Then
        Set StartQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
        StartQuizBtn.Caption = "Start Quiz"
        Set SDB.Objects("StartQuizBtn") = StartQuizBtn  
    End If
    StartQuizBtn.Enabled = False
       
    Set StopQuizBtn = SDB.Objects("StopQuizBtn")
    If StopQuizBtn Is Nothing Then
        Set StopQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
        StopQuizBtn.Caption = "Stop Quiz"
        Set SDB.Objects("StopQuizBtn") = StopQuizBtn  
    End If
    StopQuizBtn.Enabled = False

    Script.RegisterEvent NewQuizBtn, "OnClick", "NewQuiz"
    Script.RegisterEvent StartQuizBtn, "OnClick", "StartQuiz"
    Script.RegisterEvent StopQuizBtn, "OnClick", "StopQuiz"

    Script.RegisterEvent SDB, "OnShutdown", "OnShutdownHandler"
    
    Set OptionsFile = SDB.IniFile

    Call CreateMainPanel

    Call ClearSongInfoHTML
End Sub

' Hide the main player panel
Sub OnShutdownHandler
    OptionsFile.Flush
    QuizzorMainPanel.Common.Visible = False 
End Sub

Sub Uninstall 
    OptionsFile.DeleteSection "Quizzor"
    Call DestroyAllObjects
End Sub

