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

' Check whether a quiz can be started and show a message box if not
Function IsQuizReady()
    Dim IsReady 
    IsReady = IsObject(Quiz_Playlist)
    If Not IsReady Then
        SDB.MessageBox SDB.Localize("Please create a new quiz first."), mtInformation, Array(mbOk)
    End If

    IsQuizReady = IsReady
End Function

Sub DetroyAllObjects
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

Function getSongInfoHTML(SongData)
    getSongInfoHTML = "<html><body>" & vbCrLf & _
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
    ' Set QuizzorMainPanel = SDB.Objects("QuizzorMainPanel")
    ' If QuizzorMainPanel Is Nothing Then
    ' DEBUG: destroy panel and all buttons if it exists
        Set UI = SDB.UI

        Set QuizzorMainPanel = UI.NewDockablePersistentPanel("QuizzorMainPanel")
        If QuizzorMainPanel.IsNew Then
            QuizzorMainPanel.DockedTo = 4 
            QuizzorMainPanel.Common.Visible = False
        End If
        QuizzorMainPanel.ShowCaption = False

        Set PlayBtn = UI.NewButton(QuizzorMainPanel)
        PlayBtn.Common.SetRect BTN_MARGIN, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT
        PlayBtn.Common.ControlName = "PlayBtn"
        PlayBtn.Caption = SDB.Localize("Play")
        Script.RegisterEvent PlayBtn, "OnClick", "StartPlaying"

        Set NextBtn = UI.NewButton(QuizzorMainPanel)
        NextBtn.Common.SetRect 2*BTN_MARGIN + PlayBtn.Common.Width,BTN_MARGIN, _
            BTN_WIDTH, BTN_HEIGHT
        NextBtn.Common.ControlName = "NextBtn"
        NextBtn.Caption = SDB.Localize("Next")
        Script.RegisterEvent NextBtn, "OnClick", "PlayNext"

        Set ShowInfoBtn = UI.NewButton(QuizzorMainPanel)
        ShowInfoBtn.Common.SetRect 3*BTN_MARGIN+2*BTN_WIDTH, BTN_MARGIN, _
            2*BTN_WIDTH + BTN_MARGIN, BTN_HEIGHT
        ShowInfoBtn.Common.ControlName = "ShowInfoBtn"
        ShowInfoBtn.Caption = SDB.Localize("Show Information")
        Script.RegisterEvent ShowInfoBtn, "OnClick", "ShowSongInfo"

        ' TODO: Hide vertical scrollbar and/or only show when needed
        Set SongInfoHTML = UI.NewActiveX(QuizzorMainPanel, "Shell.Explorer")
        SongInfoHTML.Common.ControlName = "SongInfoHTML"
        SongInfoHTML.Common.Align = 0
        SongInfoHTML.Common.SetClientRect BTN_MARGIN, 2*BTN_MARGIN + BTN_HEIGHT, _
            QuizzorMainPanel.Common.Width - 3*BTN_MARGIN, _
            QuizzorMainPanel.Common.Height - 4*BTN_MARGIN - 2*BTN_HEIGHT
        SongInfoHTML.Interf.Navigate "about:" ' A trick to make sure document exists, from Wiki

        Set QuizTrackBar = UI.NewTrackBar(QuizzorMainPanel)
        QuizTrackBar.Common.ControlName = "QuizTrackBar"
        QuizTrackBar.Common.SetRect 0, _
            QuizzorMainPanel.Common.Height - BTN_HEIGHT - 3*BTN_MARGIN,_
            QuizzorMainPanel.Common.Width, BTN_HEIGHT
        QuizTrackBar.Common.Anchors = akBottom + akLeft
        QuizTrackBar.Horizontal = True

    ' End If
End Sub

Sub NewQuiz(Item)
    ' Ask if a new quiz should really be started
    createNew = SDB.MessageBox( SDB.Localize("Creating a new quiz replaces all  tracks") _
        + SDB.Localize(" in the current queue. This cannot be undone.") + vbCrLF _
        + SDB.Localize("Do you want to create a new quiz and lose the old quiz?"), _
        mtWarning, Array(mbNo, mbYes))
    '
    If createNew = mrNo then 
       Exit Sub 
    End If

    ' Replace playing queue with current tracks from main window 
    Call SDB.Player.PlaylistClear()
    SDB.Player.PlaylistAddTracks SDB.AllVisibleSongList
    Call RandomizePlaylist

    ' Create new empty playlist, for played tracks
    Set Quiz_Playlist = CreateNewPlaylist()

    ' TODO: Automatically select newly created playlist 
    ' TODO: Automatically hide Now Playing List
    SDB.MessageBox SDB.Localize("Please select the newly created playlist.") _
        + vbCrLf + SDB.Localize("Please hide the Now Playing playlist"), _
        mtInformation, Array(mbOk)
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
    If Not IsQuizReady() Then Exit Sub

    Set QuizTrackBar = QuizzorMainPanel.Common.ChildControl("QuizTrackBar")
    QuizTrackBar.Value = SDB.Player.CurrentSong.ID

    ' Disable playing next title
    SDB.Player.Play
    SDB.Player.StopAfterCurrent = True
End Sub

Sub PlayNext
    If Not IsQuizReady() Then Exit Sub

    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    Set HTMLDocument = SongInfoHTML.Interf.Document
    HTMLDocument.Write ""
    HTMLDocument.Close

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
    Set CurrentSong = SDB.Player.CurrentSong
    If Not IsObject(CurrentSong) Then Exit Sub

    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    Set HTMLDocument = SongInfoHTML.Interf.Document
    HTMLDocument.Write getSongInfoHTML(CurrentSong)
    HTMLDocument.Close
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
    Call DetroyAllObjects
End Sub

