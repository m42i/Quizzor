' Monkey Media Quizzor plugin
' Features:
' O Only show track info if space bar is pressed
' X Stop after current track
' O Play track again if length < 60s
' O Keep track of correctly guessed tracks

Const DEBUG_ON = True
Const BTN_MARGIN = 5 ' Defines the standard margin between buttons

' Keep track of current quiz playlist
Dim Quiz_Playlist

Dim QuizzorMainPanel, NowPlayingLabel

Function GetFormattedDate()
    Dim Today : Today = Date
    Dim This_Year : This_Year = Year(Today)
    Dim This_Month : This_Month = Month(Today)
    If This_Month < 10 Then This_Month = "0" + CStr(This_Month) 
    Dim This_Day : This_Day = Day(Today)
    If This_Day < 10 Then This_Day = "0" + CStr(This_Day) 

    GetFormattedDate = CStr(This_Year) + "-" + CStr(This_Month) + "-" + CStr(This_Day)
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

' Create a new playlist and prevent duplicates
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


Sub CreateMainPanel()
    ' Set QuizzorMainPanel = SDB.Objects("QuizzorMainPanel")
    ' If QuizzorMainPanel Is Nothing Then
        Set UI = SDB.UI

        ' DEBUG: destroy panel and all buttons if it exists
        Set QuizzorMainPanel = SDB.Objects("QuizzorMainPanel")
        If DEBUG_ON And (Not (QuizzorMainPanel Is Nothing)) Then
            DebugOutput "Remove old stuff."
            QuizzorMainPanel.Common.Visible = False

            Set PlayBtn = QuizzorMainPanel.Common.ChildControl("PlayBtn")
            If DEBUG_ON And (Not (PlayBtn Is Nothing)) Then
                PlayBtn.Common.Visible = False
                PlayBtn = Nothing
            End If

            Set NextBtn = QuizzorMainPanel.Common.ChildControl("NextBtn")
            If DEBUG_ON And (Not (NextBtn Is Nothing)) Then
                NextBtn.Common.Visible = False
                NextBtn = Nothing
            End If

            Set SongInfoLabel = QuizzorMainPanel.Common.ChildControl("SongInfoLabel")
            If DEBUG_ON And (Not (SongInfoLabel Is Nothing)) Then
                SongInfoLabel.Common.Visible = False
                SongInfoLabel = Nothing
            End If

            Set ShowInfoBtn = QuizzorMainPanel.Common.ChildControl("ShowInfoBtn")
            If DEBUG_ON And (Not (ShowInfoBtn Is Nothing)) Then
                ShowInfoBtn.Common.Visible = False
                ShowInfoBtn = Nothing
            End If

            QuizzorMainPanel = Nothing
            Set SDB.Objects("QuizzorMainPanel") = Nothing
        End If


        Set QuizzorMainPanel = UI.NewDockablePersistentPanel("QuizzorMainPanel")
        QuizzorMainPanel.DockedTo = 4 
        QuizzorMainPanel.ShowCaption = True
        QuizzorMainPanel.Common.Visible = False

        Set PlayBtn = UI.NewButton(QuizzorMainPanel)
        PlayBtn.Common.ControlName = "PlayBtn"
        PlayBtn.Caption = SDB.Localize("Play")
        PlayBtn.Common.Anchors = akLeft + akTop
        Script.RegisterEvent PlayBtn, "OnClick", "StartPlaying"

        Set NextBtn = UI.NewButton(QuizzorMainPanel)
        NextBtn.Common.ControlName = "NextBtn"
        NextBtn.Caption = SDB.Localize("Next")
        NextBtn.Common.Anchors = akTop
        NextBtn.Common.Left = PlayBtn.Common.Width + BTN_MARGIN
        Script.RegisterEvent NextBtn, "OnClick", "PlayNext"

        Set ShowInfoBtn = UI.NewButton(QuizzorMainPanel)
        ShowInfoBtn.Common.ControlName = "ShowInfoBtn"
        ShowInfoBtn.Caption = SDB.Localize("Show Information")
        ShowInfoBtn.Common.Anchors = akTop + akLeft
        ShowInfoBtn.Common.Top = PlayBtn.Common.Height + BTN_MARGIN
        ShowInfoBtn.Common.Width = ShowInfoBtn.Common.Width * 2 + BTN_MARGIN
        Script.RegisterEvent ShowInfoBtn, "OnClick", "ShowSongInfo"

        Set SongInfoLabel = UI.NewLabel(QuizzorMainPanel)
        SongInfoLabel.Common.ControlName = "SongInfoLabel"
        SongInfoLabel.Alignment = 0
        SongInfoLabel.Autosize = True
        SongInfoLabel.Common.Anchors = akLeft + akTop
        SongInfoLabel.Common.Top = PlayBtn.Common.Height + ShowInfoBtn.Common.Height + 2*BTN_MARGIN
        SongInfoLabel.Common.FontSize = SongInfoLabel.Common.FontSize * 3
        SongInfoLabel.Caption = ""

    ' End If
End Sub

Sub NewQuiz(Item)
    ' Ask if a new quiz should really be started
    '  createNew = SDB.MessageBox( SDB.Localize("Creating a new quiz replaces all  tracks in the current queue. This cannot be undone. Do you want to create a new quiz and lose the old quiz?"), mtWarning, Array(mbNo, mbYes))
    '
    '  If createNew = mrNo then 
    '    Exit Sub 
    '  End If

    ' Replace playing queue with current tracks from main window 
    Call SDB.Player.PlaylistClear()
    SDB.Player.PlaylistAddTracks SDB.AllVisibleSongList
    Call RandomizePlaylist

    ' Create new empty playlist, for played tracks
    Set Quiz_Playlist = CreateNewPlaylist()

    ' TODO: Automaticallz select newly created playlist 
    ' TODO: Automatically hide Now Playing List
    ' SDB.MessageBox SDB.Localize("Please select the newly created playlist.") _
        ' + vbCrLf + SDB.Localize("Please hide the Now Playing playlist"), _
        ' mtInformation, Array(mbOk)
End Sub

Sub StartQuiz(Item)
    QuizzorMainPanel.Common.Visible = True

    SDB.Player.CurrentSongIndex = 0
End Sub

Sub StopQuiz(Item)
    QuizzorMainPanel.Common.Visible = False
    SDB.Player.Stop
End Sub

Sub StartPlaying
    ' Disable playing next title
    SDB.Player.Play
    SDB.Player.StopAfterCurrent = True
End Sub

Sub PlayNext
    Set SongInfoLabel = QuizzorMainPanel.Common.ChildControl("SongInfoLabel")
    SongInfoLabel.Caption = ""

    If SDB.Player.CurrentPlaylist.Count = 0 Then
        SDB.MessageBox SDB.Localize("Quiz has ended. Please create a new one."), _
            mtInformation, Array(mbOk)
        Exit Sub
    End If

    Quiz_Playlist.addTrack SDB.Player.CurrentSong
    SDB.Player.PlaylistDelete 0

    ' Disable playing next title
    SDB.Player.Play
    SDB.Player.StopAfterCurrent = True
End Sub

Sub ShowSongInfo
    Set SongInfoLabel = QuizzorMainPanel.Common.ChildControl("SongInfoLabel")
    Set CurrentSong = SDB.Player.CurrentSong
    SongInfoLabel.Caption = SDB.Localize("Album") + vbTab + CurrentSong.AlbumName + vbCrLf _
        + SDB.Localize("Title") + vbTab + CurrentSong.Title + vbCrLf _
        + SDB.Localize("Artist") + vbTab + CurrentSong.ArtistName + vbCrLf _
        + SDB.Localize("Comment") + vbTab + CurrentSong.Comment + vbCrLf 

    SongInfoLabel.Common.Visible = True
End Sub

Sub OnStartup
    Set UI = SDB.UI
    
    ' Register new or get existing toolbar 
    Set QuizBar = UI.AddToolbar("QuizBar")
    Set SDB.Objects("QuizBar") = QuizBar
       
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
       
    Set StopQuizBtn = SDB.Objects("StopQuizBtn")
    If StopQuizBtn Is Nothing Then
        Set StopQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
        StopQuizBtn.Caption = "Stop Quiz"
        Set SDB.Objects("StopQuizBtn") = StopQuizBtn  
    End If
    
    Script.RegisterEvent NewQuizBtn, "OnClick", "NewQuiz"
    Script.RegisterEvent StartQuizBtn, "OnClick", "StartQuiz"
    Script.RegisterEvent StopQuizBtn, "OnClick", "StopQuiz"

    Call CreateMainPanel
End Sub


Sub Uninstall 
    Set QuizBar = SDB.Objects("QuizBar")
    If Not QuizBar Is Nothing then
        SDB.Objects("QuizBar").Visible = False
        Set QuizBar = Nothing
        Set SDB.Objects("QuizBar") = Nothing
    End If
End Sub




