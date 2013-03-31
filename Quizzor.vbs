' Monkey Media Quizzor plugin
' Features:
' O Only show track info if space bar is pressed
' O Stop after current track
' O Play track again if length < 60s
' O Keep track of correctly guessed tracks

' Keep track of current quiz playlist
Dim Quiz_Playlist

Dim QuizzorMainPanel

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

        ' DEBUG: destroy panel if it exists
        Set QuizzorMainPanel = SDB.Objects("QuizzorMainPanel")
        If Not (QuizzorMainPanel Is Nothing) Then
            QuizzorMainPanel.Common.Visible = False
            Set SDB.Objects("QuizzorMainPanel") = Nothing
            QuizzorMainPanel = Nothing
        End If

        Set QuizzorMainPanel = UI.NewDockablePersistentPanel("QuizzorMainPanel")
        QuizzorMainPanel.DockedTo = 4 
        QuizzorMainPanel.ShowCaption = True
        QuizzorMainPanel.Common.Visible = False
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
    
    ' TODO: Automatic playlist selection
    ' TODO: Automaticly hide Now Playing List
    ' SDB.MessageBox SDB.Localize("Please select the newly created playlist.") _
        ' + vbCrLf + SDB.Localize("Please hide the Now Playing playlist"), _
        ' mtInformation, Array(mbOk)
End Sub

Sub StartQuiz(Item)
    QuizzorMainPanel.Common.Visible = True
End Sub

Sub StopQuiz(Item)
    QuizzorMainPanel.Common.Visible = False
End Sub

Sub OnStartup
    Set UI = SDB.UI
    
    Set QuizBar = SDB.Objects("QuizBar")
    If QuizBar Is Nothing then
        ' Register new menu item in tools menu
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




