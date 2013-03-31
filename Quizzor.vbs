' Monkey Media Quizzor plugin
' Features:
' O Only show track info if space bar is pressed
' O Only show tracknumber and length in playlist
' O Stop after current track
' O Play track again if length < 60s
' O Keep track of correctly guessed tracks

Dim Quiz_Played_Playlist

Function GetFormattedDate()
    Dim Today : Today = Date
    Dim This_Year : This_Year = Year(Today)
    Dim This_Month : This_Month = Month(Today)
    If This_Month < 10 Then This_Month = "0" + CStr(This_Month) End If
    Dim This_Day : This_Day = Day(Today)
    If This_Day < 10 Then This_Day = "0" + CStr(This_Day) End If

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

Sub Shuffle(n)
    Randomize
    j = n - 1
    For i = 0 To n - 1
        SDB.Player.PlaylistMoveTrack i,Int(n*Rnd)
        SDB.Player.PlaylistMoveTrack j,Int(n*Rnd)
        j = j - 1
    Next
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
    Set Playlist_Root = SDB.PlaylistByTitle("")
    Set Quiz_Played_Playlist = Playlist_Root.CreateChildPlaylist(SDB.Localize("Quiz of " + GetFormattedDate()))
    Quiz_Played_Playlist.Selected = True
    ' TODO: Automatic playlist selection
    DebugOutput SDB.Localize("Select the newly created playlist."), mtInformation, Array(mbOk)

    ' TODO: Automaticly hide Now Playing List
    DebugOutput SDB.Localize("Please hide the Now Playing playlist"), mtInformation, Array(mbOk) 
    
End Sub

Sub StartQuiz(Item)
    ' Turn off shuffle
    ' Hide play controls
End Sub

Sub StopQuiz(Item)
End Sub

Sub OnStartup
    Set UI = SDB.UI
    
    Set QuizBar = SDB.Objects("QuizBar")
    If Not QuizBar Is Nothing then
        SDB.Objects("QuizBar").Visible = False
        Set QuizBar = Nothing
        Set SDB.Objects("QuizBar") = Nothing
    End If
    
    ' Register new menu item in tools menu
    Set QuizBar = UI.AddToolbar("QuizBar")
    Set SDB.Objects("QuizBar") = QuizBar
    
    Set NewQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
    NewQuizBtn.Caption = "New Quiz"
    
    Set StartQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
    StartQuizBtn.Caption = "Start Quiz"
    
    Set StopQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
    StopQuizBtn.Caption = "Stop Quiz"
    
    Script.RegisterEvent NewQuizBtn, "OnClick", "NewQuiz"
    Script.RegisterEvent StartQuizBtn, "OnClick", "StartQuiz"
    Script.RegisterEvent StopQuizBtn, "OnClick", "StopQuiz"
    
End Sub

Sub Uninstall 
    Set QuizBar = SDB.Objects("QuizBar")
    If Not QuizBar Is Nothing then
        SDB.Objects("QuizBar").Visible = False
        Set QuizBar = Nothing
        Set SDB.Objects("QuizBar") = Nothing
    End If
End Sub




