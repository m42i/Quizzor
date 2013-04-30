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
' "Big" Features:
' O Only show track info if space bar is pressed
' X Stop after current track
' O Play track again if length < 60s
' O Keep track of correctly guessed tracks

Const DEBUG_ON = True

' Defines the standard margin between buttons
' 9 is equal to the margin of a group box in an options sheet
Const BTN_MARGIN = 9
' Defines standard height of a button
' 24 equals the normal button height, e.g. 'Ok', 'Cancel' 
Const BTN_HEIGHT = 24
' Defines standard width of a button
' 74 equals the normal button width, e.g. 'Ok', 'Cancel' 
Const BTN_WIDTH = 74
Const BTN_LONG_WIDTH = 111 ' Defines width of a long button
Const TIME_WIDTH = 50 ' Defines standard width of a time label

' Node types
' see http://www.mediamonkey.com/wiki/index.php/MediaMonkey_Tree_structure
Const NODE_PLAYLIST_ROOT = 6
Const NODE_PLAYLIST_AUTO = 71
Const NODE_PLAYLIST_MANUAL = 61

' %font-size% should be replaced with the font size, e.g. 200%
Const HTML_Style = "<style type='text/css'> body { overflow: auto; } table { font-size: %font-size%; font-family: Verdana, sans-serif; } </style>" 

' Anchor constants, add them for multiple anchors
Const akLeft = 1
Const akTop = 2
Const akRight = 4
Const akBottom = 8
Const akAll = 15 ' Sum of all above

' Element alignment on panels/forms
Const alNone = 0
Const alTop = 1
Const alBottom = 2
Const alLeft = 3
Const alRight = 4
Const alClient = 5

' Text alignment
Const txtAlLeft = 0
Const txtAlRight = 1
Const txtAlCenter = 2

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
' LastSongIndexForPlaylist_<Playlist.ID> = <SDB.Player.CurrentSongIndex> As Long
'
Dim OptionsFile

' Creates a modal message box window, with the "Text".
' Buttons is an Array of Strings, arranged from right to left, aligned right
' Empty strings will not create buttons, usefule to specify the return value
' Return value is the String position in the array Buttons()
' If a button is not pressed (e.g. window closed), 
' the return value will be negative 
' and by 100 smaller than the default modal result
' e.g. -98 for Cancel (2)
Function FreeFormMessageBox(Text, Buttons())
    ' Construct form
    Set MsgWindow = SDB.UI.NewForm
    MsgWindow.Common.ClientWidth = 300
    MsgWindow.Common.ClientHeight = 150
    MsgWindow.BorderStyle = 3 ' non-resizable dialog
    MsgWindow.FormPosition = 4 ' screen center

    Set MsgText = SDB.UI.NewLabel(MsgWindow)
    MsgText.Alignment = txtAlLeft
    MsgText.Caption = Text
    MsgText.Common.Left = 2*BTN_MARGIN
    MsgText.Common.Top = 2*BTN_MARGIN
    MsgText.Common.ClientWidth = MsgWindow.Common.ClientWidth - 2*BTN_MARGIN
    MsgText.Common.ClientHeight = _
        MsgWindow.Common.ClientHeight - BTN_HEIGHT - 2*BTN_MARGIN
    MsgText.Multiline = True

    ' create buttons, set ModalResult
    ' Use extra variable to prevent big spaces between buttons
    btnNr = 0 
    For i = 0 To UBound(Buttons)
        If Buttons(i) <> "" Then 
            Set Button = SDB.UI.NewButton(MsgWindow)
            Button.Common.SetClientRect _
                MsgWindow.Common.ClientWidth - BTN_WIDTH*(btnNr+1) - BTN_MARGIN*(btnNr+1), _
                MsgWindow.Common.ClientHeight - BTN_HEIGHT - BTN_MARGIN, _
                BTN_WIDTH, BTN_HEIGHT
            Button.Caption = Buttons(i)

            ' We need to add 100, otherwise it would interfer with system defaults
            Button.ModalResult = 100 + i
            btnNr = btnNr + 1
        End If
    Next

    ' Return button's modal result
    FreeFormMessageBox = MsgWindow.ShowModal - 100
End Function

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

' Returns a new SDBSongList with all SongData in the current queue
Function ClonePlaylist(ByVal SongList)
    Set NewPlaylist = SDB.NewSongList

    If SongList.Count > 0 Then 
        For i = 0 To SongList.Count - 1
            NewPlaylist.Add SongList.Item(i)
        Next
    End If

    Set ClonePlaylist = NewPlaylist
End Function

Sub Shuffle(Playlist)
    Randomize
    Set Tracks = Playlist.Tracks
    n = Tracks.Count
    j = n - 1
    For i = 0 To n - 1
        Playlist.MoveTrack Tracks.Item(i), Tracks.Item(Int(n*Rnd))
        Playlist.MoveTrack Tracks.Item(j), Tracks.Item(Int(n*Rnd))
        j = j - 1
    Next
End Sub

' Thanks to Diddeleedoo from the MM forums
Sub RandomizePlaylist(Item)
    If Not IsPlaylistNode() Then Exit Sub

    DoShuffle = FreeFormMessageBox(SDB.Localize("Randomizing a playlist cannot be undone."), _
        Array(SDB.Localize("Randomize"), SDB.Localize("Cancel")))
    If DoShuffle <> 0 Then Exit Sub

    ' Because of OLE error 80020006 with wine, the queue is used for shuffling
    ' and restored afterwards
    Set NodePlaylist = SDB.PlaylistByID( SDB.MainTree.CurrentNode.RelatedObjectID )

    song_count = NodePlaylist.Tracks.Count
    If song_count > 1 Then 
        Shuffle NodePlaylist
        SDB.MainTracksWindow.Refresh
    End If
End Sub

' Get a comma seperated string list of all IDs in a given SongList
Function GetSongIDList(SongList)
    If SongList.Count = 0 Then 
        GetSongIDList = ""
        Exit Function
    End If

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
    While Not Playlist_Root.Title = "" ' Checking for root
        NewTitle = NewBaseTitle + " (" + CStr(i) + ")"
        Set Playlist_Root = SDB.PlaylistByTitle(NewTitle)
        i = i + 1
    WEnd

    Set CreateNewPlaylist = Playlist_Root.CreateChildPlaylist(NewTitle)
End Function

' Check whether a quiz exists, and displays a message if not
Function QuizExists()
    Dim QExists : QExists = True

    If (Not IsObject(Quiz_Playlist)) Or (Quiz_Playlist Is Nothing) Then
        QExists = False

        SDB.MessageBox SDB.Localize("Please create a new quiz first."), mtInformation, Array(mbOk)
    End If
    
    QuizExists = QExists
End Function
    
' Check whether songs are visible, and display message if not
Function SongsVisible()
    Dim SVisible : SVisible = True

    If SDB.AllVisibleSongList.Count = 0 Then
        SVisible = False

        SDB.MessageBox SDB.Localize("No songs visible. Please fill the main window first."),_
        mtInformation, Array(mbOk)
    End If

    SongsVisible = SVisible
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

    Set HideInfoBtn = SDB.Objects("HideInfoBtn")
    If Not (HideInfoBtn Is Nothing) Then
        HideInfoBtn.Common.Visible = False
        HideInfoBtn = Nothing
    End If

    Set QuizzorMainPanel = SDB.Objects("QuizzorMainPanel")
    If IsObject(QuizzorMainPanel) And Not (QuizzorMainPanel Is Nothing) Then
        QuizzorMainPanel.Common.Visible = False
        Set SDB.Objects("QuizzorMainPanel") = Nothing
    End If
    
    Set SDB.Objects("QuizBar") = Nothing
    Set SDB.Objects("NewQuizBtn") = Nothing

    Script.UnRegisterEvents SDB
End Sub

' Clears SongInfoHTML end ensures the style is preserved
Sub ClearSongInfoHTML
    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    Set HTMLDocument = SongInfoHTML.Interf.Document
    HTMLDocument.Write "<html>" & vbCrLf & Replace(HTML_Style,"%font-size%","100%") & vbCrLf & "<body>&nbsp;</body></html>"
    HTMLDocument.Close
End Sub

Function GetSongInfoHTML(SongData)
    ' Get window size, for reference to the font size
    Dim WindowHeight
    If IsObject(QuizzorMainPanel) Then
        WindowHeight = QuizzorMainPanel.Common.ClientHeight
    Else
        WindowHeight = 800
    End If 

    GetSongInfoHTML = "<html>" & vbCrLf & _
        Replace(HTML_Style, "%font-size%", CStr(WindowHeight / 2) & "%") & vbCrLf & _
        "<table border='1' cellspacing='0' cellpaddin='2' rules='rows'" & _
        " frame='void' width='100%' height='100%'>" & vbCrLf & _
        "<colgroup>" & vbCrLf & _
            "<col width='20%'>" & vbCrLf & _
            "<col width='80%'>" & vbCrLf & _
        "</colgroup>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td valign='top' align='left'>" & SDB.Localize("Album") & "</td>" & vbCrLf & _
            "<td valign='top'>" & SongData.AlbumName & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td valign='top' align='left'>" & SDB.Localize("Title") & "</td>" & vbCrLf & _
            "<td valign='top'>" & SongData.Title & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td valign='top' align='left'>" & SDB.Localize("Artist") & "</td>" & vbCrLf & _
            "<td valign='top'>" & SongData.ArtistName & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td valign='top' align='left'>" & SDB.Localize("Comment") & "</td>" & vbCrLf & _
            "<td valign='top'>" & SongData.Comment & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "<tr>" & vbCrLf & _
            "<td valign='top' align='left'>" & SDB.Localize("File") & "</td>" & vbCrLf & _
            "<td valign='top'>" & SongData.Path & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf & _
        "</body></html>"
End Function

Sub CreateMainPanel()
    Set UI = SDB.UI
    
    Set QuizzorMainPanel = UI.NewForm
    QuizzorMainPanel.BorderStyle = 2
    QuizzorMainPanel.Common.Visible = False
    QuizzorMainPanel.Common.Align = alClient
    Script.RegisterEvent QuizzorMainPanel, "OnClose", "StopQuiz"

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

    Set HideInfoBtn = UI.NewButton(QuizzorMainPanel)
    HideInfoBtn.Common.ControlName = "HideInfoBtn"
    HideInfoBtn.Common.Visible = False
    HideInfoBtn.Caption = SDB.Localize("Hide Information")
    Script.RegisterEvent HideInfoBtn, "OnClick", "HideSongInfo"
       
    ' TODO: Implement a close button, testing with wine gives an OLE error
    ' Set StopQuizBtn = UI.NewButton(QuizzorMainPanel)
    ' StopQuizBtn.Common.ControlName = "StopQuizBtn"
    ' StopQuizBtn.Caption = SDB.Localize("Stop Quiz")
    ' Script.RegisterEvent StopQuizBtn, "OnClick", "StopBtnClicked"

    Set SongTime = UI.NewLabel(QuizzorMainPanel)
    SongTime.Common.ControlName = "SongTime"
    SongTime.Common.Anchors = akLeft + akTop
    SongTime.Alignment = txtAlCenter
    SongTime.Autosize = False
    SongTime.Caption = "00:00"

    ' TODO: Change playback time, when TrackBar changes
    Set SongTrackBar = UI.NewTrackBar(QuizzorMainPanel)
    SongTrackBar.Common.ControlName = "SongTrackBar"
    SongTrackBar.Common.Anchors = akTop
    SongTrackBar.Common.Enabled = False
    SongTrackBar.Common.Anchors = akLeft + akBottom + akRight
    SongTrackBar.Value = 0
    SongTrackBar.Horizontal = True

    Set SongTimeLeft = UI.NewLabel(QuizzorMainPanel)
    SongTimeLeft.Common.ControlName = "SongTimeLeft"
    SongTimeLeft.Common.Anchors = akTop + akRight
    SongTimeLeft.Alignment = txtAlCenter
    SongTimeLeft.Autosize = False
    SongTimeLeft.Caption = "00:00"

    Set SongInfoHTML = UI.NewActiveX(QuizzorMainPanel, "Shell.Explorer")
    SongInfoHTML.Common.ControlName = "SongInfoHTML"
    SongInfoHTML.Common.Align = alNone
    SongInfoHTML.Common.Anchors = akLeft + akTop + akRight + akBottom
    SongInfoHTML.Interf.Navigate "about:" ' A trick to make sure document exists, from Wiki
    ResizeMainPanel

    ' Always hide Main Panel if it is created
    QuizzorMainPanel.Common.Visible = False
End Sub

Sub ResizeMainPanel
    If Not IsObject(QuizzorMainPanel) Then
        CreateMainPanel
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

    Set HideInfoBtn = QuizzorMainPanel.Common.ChildControl("HideInfoBtn")
    HideInfoBtn.Common.SetRect 4*BTN_MARGIN+3*BTN_WIDTH, BTN_MARGIN, _
        2*BTN_WIDTH + BTN_MARGIN, BTN_HEIGHT

    ' Set StopQuizBtn = QuizzorMainPanel.Common.ChildControl("StopQuizBtn")
    ' StopQuizBtn.Common.SetRect _
        ' QuizzorMainPanel.Common.ClientWidth - BTN_MARGIN - BTN_WIDTH, _
        ' BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT

    Set SongTime = QuizzorMainPanel.Common.ChildControl("SongTime")
    SongTime.Common.SetRect 2*BTN_MARGIN, BTN_HEIGHT + 3*BTN_MARGIN, _
        TIME_WIDTH, BTN_HEIGHT

    Set SongTrackBar = QuizzorMainPanel.Common.ChildControl("SongTrackBar")
    SongTrackBar.Common.SetRect BTN_MARGIN + TIME_WIDTH, _
        BTN_HEIGHT + 3*BTN_MARGIN,_
        QuizzorMainPanel.Common.Width - 2*TIME_WIDTH - 2*BTN_MARGIN, _
        BTN_HEIGHT

    Set SongTimeLeft = QuizzorMainPanel.Common.ChildControl("SongTimeLeft")
    SongTimeLeft.Common.SetRect _
        BTN_MARGIN + TIME_WIDTH + SongTrackBar.Common.Width, _
        BTN_HEIGHT + 3*BTN_MARGIN, TIME_WIDTH, BTN_HEIGHT

    Set SongInfoHTML = QuizzorMainPanel.Common.ChildControl("SongInfoHTML")
    SongInfoHTML.Common.SetClientRect BTN_MARGIN, _
        3*BTN_MARGIN + 2*BTN_HEIGHT, _
        QuizzorMainPanel.Common.Width - 3*BTN_MARGIN, _
        QuizzorMainPanel.Common.Height - 6*BTN_MARGIN - 2*BTN_HEIGHT
End Sub

' Open the playlists node
Sub SelectPlaylist(Playlist)
    Set Root = SDB.MainTree
    Set ParentPlaylistNode = Root.Node_Playlists
    
    ' Iterate through all nodes until Playlist is found
    Set PlaylistNode = Root.FirstChildNode(ParentPlaylistNode)
    While Not (PlaylistNode Is Nothing) 
        If PlaylistNode.RelatedObjectID = Playlist.ID Then
            Set Root.CurrentNode = PlaylistNode
            PlaylistNode.Expanded = True
            Exit Sub
        End If
        Set PlaylistNode = Root.NextSiblingNode(PlaylistNode)
    WEnd
End Sub

' Check whether the saved playlists already exist and delete if not
' This should be called whenever a playlist value is read
Sub UpdateOptionsFile
    ' Check if the last QuizPlaylist still exists and delete the key if not
    If OptionsFile.ValueExists("Quizzor", "LastPlaylistID") Then
        Set Playlist = SDB.PlaylistByID(OptionsFile.IntValue("Quizzor", "LastPlaylistID")) 
        ' If no playlist exists, root (ID=0) is returned
        If Playlist.ID = 0 Then
            OptionsFile.DeleteKey "Quizzor", "LastPlaylistID"
        End If
    End If
    
    ' Go through all saved playlists and check if they exist
    ' A playlist is saved with the key "LastSongIndexForPlaylist_<Playlist.ID>"
    Set KeyList = OptionsFile.Keys("Quizzor")
    For i = 0 To KeyList.Count - 1
        KeyValue = KeyList.Item(i)
        ' Keys method returns each string as "key=value"
        Key = Left(KeyValue, InStr(KeyValue, "=") - 1)
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
    UpdateOptionsFile
    
    Dim NewQuizDialogAnswer
    Dim OptionsArray
    Dim DialogText
    If OptionsFile.ValueExists("Quizzor", "LastPlaylistID") Then
        LastPlaylistID = OptionsFile.IntValue("Quizzor", "LastPlaylistID")
        Set LastPlaylist = SDB.PlaylistByID(LastPlaylistID)

        DialogText = SDB.Localize("A previous quiz exists. Do you want to restore the last quiz") & _
            SDB.LocalizedFormat(" %s or create a new quiz?", LastPlaylist.Title,0,0) & vbCrLf & _
            SDB.Localize("Either way the current queue will be lost.") 

        OptionsArray = Array(SDB.Localize("New Quiz"), _
            SDB.Localize("Restore Quiz"), SDB.Localize("Cancel"))
    Else
        DialogText = _
            SDB.Localize("Creating a new quiz means losing the current queue.")

        OptionsArray = Array(SDB.Localize("New Quiz"), "", _
            SDB.Localize("Cancel"))
    End If

    NewQuizDialogAnswer = FreeFormMessageBox(DialogText, OptionsArray)

    If NewQuizDialogAnswer = 1 Then
        RestoreLastSession
    ElseIf NewQuizDialogAnswer = 0 Then
        If SongsVisible() Then
            CreateNewQuiz
        End If
    Else
        Exit Sub
    End If
    
    If IsObject(Quiz_Playlist) Then
        StartQuiz(Item)
    End If
End Sub

' Creates a new quiz, resetting the current without asking
Sub CreateNewQuiz
    ' The user decided to create a new playlist, so we clear the current
    StopQuiz(Item)

    ' Replace playing queue with current tracks from main window 
    SDB.Player.PlaylistClear
    SDB.Player.PlaylistAddTracks SDB.AllVisibleSongList
    RandomizePlaylist

    ' Create new empty playlist, for played tracks
    Set Quiz_Playlist = CreateNewPlaylist()
    SelectPlaylist Quiz_Playlist

    ' Save new playlist data to ini file
    OptionsFile.IntValue("Quizzor", "LastPlaylistID") = Quiz_Playlist.ID
    OptionsFile.StringValue("Quizzor", "NowPlayingSongs_" + CStr(Quiz_Playlist.ID)) = _
        GetSongIDList(SDB.Player.CurrentSongList)
End Sub

' Check if the selected item is a playlist
' show message if not
Function IsPlaylistNode
    Set SelectedNode = SDB.MainTree.CurrentNode
    If (SelectedNode.NodeType = NODE_PLAYLIST_AUTO) _
            Or (SelectedNode.NodeType = NODE_PLAYLIST_MANUAL) Then
        IsPlaylistNode = True
    Else
        DebugOutput SDB.Localize("Please select a playlist.")
        IsPlaylistNode = False
    End If
End Function

Sub StartQuiz(Item)
    If Not IsPlaylistNode() Then Exit Sub

    Set SelectedNode = SDB.MainTree.CurrentNode
    Set Quiz_Playlist = SDB.PlaylistByID(SelectedNode.RelatedObjectID)
    Set SongList = Quiz_Playlist.Tracks
    If SongList.Count <= 0 Then
        DebugOutput SDB.Localize("Playlist is empty. Won't start quiz.")
        Exit Sub
    End If

    ' TODO: Set the first track in main track window focused
    ' Using the queue is a workaround, since the above TODO doesn't work
    If SDB.Player.PlaylistCount > 0 Then
        OverwriteQueue = FreeFormMessageBox( _
            SDB.Localize("The current queue is not empty. Do you want to replace all tracks?"), _
            Array(SDB.Localize("Replace queue"), SDB.Localize("Cancel")))
        If OverwriteQueue <> 0 Then Exit Sub
    End If

    If Not IsObject(QuizzorMainPanel) Then
        CreateMainPanel
    End If

    QuizzorMainPanel.Common.Visible = True
    ' Ensure that the elements are redrawn
    ResizeMainPanel
    
    SongTime.Caption = GetFormattedTime(0)
    SongTimeLeft.Caption = GetFormattedTime(0)

    SDB.Player.PlaylistClear
    SDB.Player.PlaylistAddTracks SongList

    ' Restore last index in playlist
    ResumeIndex = 0
    If OptionsFile.ValueExists("Quizzor", _ 
        "LastSongIndexForPlaylist_" + CStr(Quiz_Playlist.ID)) Then
        ResumeIndex = OptionsFile.IntValue("Quizzor", "LastSongIndexForPlaylist_" + _
            CStr(Quiz_Playlist.ID))
    End If

    SDB.Player.CurrentSongIndex = ResumeIndex
End Sub

Sub StopBtnClicked(Item)
End Sub

Sub StopQuiz(Item)
    UpdateOptionsFile

    If SDB.Player.isPlaying And IsObject(SongTimer) Then
        SongTimer.Enabled = False
        Script.UnRegisterEvents SongTimer
    End If

    SDB.Player.Stop
    
    SongTime.Caption = GetFormattedTime(0)
    SongTimeLeft.Caption = GetFormattedTime(0)

    If IsObject(Quiz_Playlist) Then Set Quiz_Playlist = Nothing
    SDB.ProcessMessages ' Ensure, that changes to Quiz_Playlist are applied

    HideSongInfo
End Sub

Sub StartPlaying
    If Not QuizExists() Then Exit Sub 

    If SDB.Player.CurrentPlaylist.Count <= 0 Then
        SDB.MessageBox SDB.Localize("Empty queue. Please create a new quiz."), _
            mtInformation, Array(mbOk)
        StopQuiz Nothing
        Exit Sub
    End If
    
    ' If the player is paused, just continue playing.
    If SDB.Player.isPaused Then
        SDB.Player.Pause
        Exit Sub
    End If

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
    SDB.Player.Pause
End Sub

Sub PlayNext
    If Not QuizExists() Then Exit Sub

    HideSongInfo

    NewIndex = SDB.Player.CurrentSongIndex + 1
    OptionsFile.IntValue("Quizzor", "LastSongIndexForPlaylist_" + CStr(Quiz_Playlist.ID)) = _
        NewIndex

    If SDB.Player.CurrentSongIndex = SDB.Player.CurrentSongList.Count - 1 Then
        SDB.MessageBox SDB.Localize("No more songs in queue.")
            mtInformation, Array(mbOk)
        Exit Sub
    End If

    SDB.Player.CurrentSongIndex = NewIndex
    StartPlaying
End Sub

Sub HideSongInfo
    Set ShowInfoBtn = QuizzorMainPanel.Common.ChildControl("ShowInfoBtn")
    ShowInfoBtn.Common.Visible = True
    Set HideInfoBtn = QuizzorMainPanel.Common.ChildControl("HideInfoBtn")
    HideInfoBtn.Common.Visible = False
    
    ClearSongInfoHTML
End Sub

Sub ShowSongInfo
    Set CurrentSong = SDB.Player.CurrentSong
    If Not IsObject(CurrentSong) Then Exit Sub
    
    Set ShowInfoBtn = QuizzorMainPanel.Common.ChildControl("ShowInfoBtn")
    ShowInfoBtn.Common.Visible = False
    Set HideInfoBtn = QuizzorMainPanel.Common.ChildControl("HideInfoBtn")
    HideInfoBtn.Common.Visible = True

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
    SDB.Player.PlaylistClear

    LastPlaylistID = OptionsFile.IntValue("Quizzor", "LastPlaylistID")
    Set Quiz_Playlist = SDB.PlaylistByID(LastPlaylistID)

    ' Fill Now Playing List
    ' TODO: Restore playlist order
    SongIDList = OptionsFile.StringValue("Quizzor", "NowPlayingSongs_" + CStr(LastPlaylistID))
    Set SongIter = SDB.Database.QuerySongs("ID in (" + SongIDList + ")")
    While Not SongIter.EOF
        SDB.Player.PlaylistAddTrack SongIter.Item
        SongIter.Next
    WEnd

    SelectPlaylist Quiz_Playlist
End Sub

Sub BuildSheet(Sheet)
    Set EnableRandomImages = SDB.UI.NewCheckBox(Sheet)
    EnableRandomImages.Caption = SDB.Localize("Enable random images")
    EnableRandomImages.Checked = True
    EnableRandomImages.Common.SetRect BTN_WIDTH,0,BTN_WIDTH,BTN_HEIGHT
    ' EnableRandomImages.Common.Align = 1
    EnableRandomImages.Common.ControlName = "EnableRandomImages"

    Set RandomImagesBox = SDB.UI.NewGroupBox(Sheet)
    RandomImagesBox.Common.SetRect 0,20,500,90
    RandomImagesBox.Common.ControlName = "EnableRandonImageBox"

    Set LoadImages = SDB.UI.NewButton(RandomImagesBox)
    LoadImages.Caption = SDB.Localize("Load Images")
    LoadImages.UseScript = Script.ScriptPath
    LoadImages.OnClickFunc = "LoadImages"
    LoadImages.Common.SetRect 10,10,75,20
    LoadImages.Common.ControlName = "LoadImagesBtn"

    Set ImageCount = SDB.UI.NewLabel(RandomImagesBox)
    ImageCount.Common.SetRect 90,10,200,25
    ImageCount.Common.ControlName = "ImagesLoadedLabel"
    ImageCount.Caption = SDB.LocalizedFormat("%d images loaded", 0,0,0)

    Set MinImageWaitTitles = SDB.UI.NewSpinEdit(RandomImagesBox)
    MinImageWaitTitles.MinValue = 1
    MinImageWaitTitles.Value = 1
    MinImageWaitTitles.Common.SetRect 120,40,50,20
    MinImageWaitTitles.Common.ControlName = "MinImageWaitTitles"

    Set MaxImageWaitTitles = SDB.UI.NewSpinEdit(RandomImagesBox)
    MaxImageWaitTitles.MinValue = 1
    MaxImageWaitTitles.Value = 1
    MaxImageWaitTitles.Common.SetRect 200,40,50,20
    MaxImageWaitTitles.Common.ControlName = "MaxImageWaitTitles"

    Set Label3 = SDB.UI.NewLabel(RandomImagesBox)
    Label3.Common.SetRect 0,0,65,17

    Set TranspRandomImagesBox = SDB.UI.NewTranspPanel(Sheet)
    TranspRandomImagesBox.Common.SetRect 427,287,100,90
    TranspRandomImagesBox.Common.ControlName = "TranspRandomImagesBox"
End Sub

Sub SaveSheet(Sheet)
    EnableRandomImages = _
    Sheet.Common.ChildControl("EnableRandomImages").Checked
End Sub

Sub OnStartup
    Set UI = SDB.UI

    ' Add right-click menu
    Set QuizMenuPopSeperator = SDB.Objects("QuizMenuPopSeperator") 
    If QuizMenuPopSeperator Is Nothing Then
        Set QuizMenuPopSeperator = UI.AddMenuItemSep(UI.Menu_Pop_Tree, 0, 0)
        SDB.Objects("QuizMenuPopSeperator") = QuizMenuPopSeperator
    End If

    Set BeginQuizMenuItem = SDB.Objects("BeginQuizMenuItem") 
    If BeginQuizMenuItem Is Nothing Then
        Set BeginQuizMenuItem = UI.AddMenuItem(UI.Menu_Pop_Tree, 0, -1)
        SDB.Objects("BeginQuizMenuItem") = BeginQuizMenuItem 
    End If
    BeginQuizMenuItem.Caption = SDB.Localize("Begin Quiz")
    Script.RegisterEvent BeginQuizMenuItem, "OnClick", "StartQuiz"
    
    Set RandomizePlaylistMenuItem = SDB.Objects("RandomizePlaylistMenuItem")
    If RandomizePlaylistMenuItem Is Nothing Then
        Set RandomizePlaylistMenuItem = UI.AddMenuItem(UI.Menu_Pop_Tree, 0, -1)
        SDB.Objects("RandomizePlaylistMenuItem") = RandomizePlaylistMenuItem  
    End If
    RandomizePlaylistMenuItem.Caption = SDB.Localize("Randomize")
    Script.RegisterEvent RandomizePlaylistMenuItem, _
                                    "OnClick", "RandomizePlaylist"
    
    Script.RegisterEvent SDB, "OnShutdown", "OnShutdownHandler"

    Set OptionsFile = SDB.IniFile

    ' Create options sheet
    If DEBUG_ON Then
        ' Create a frame with the options for rapid prototyping
        Set OptionsForm = UI.NewForm
        SDB.Objects("OptionsForm") = OptionsForm
        OptionsForm.Common.SetClientRect 200, 100, 600, 400
        OptionsForm.FormPosition = 4 ' screen center

        BuildSheet(OptionsForm)

        OptionsForm.ShowModal
    Else
        OptionsSheet = UI.AddOptionSheet("Quizzor", Script.ScriptPath, _
                "BuildSheet", "SaveSheet", 0)
    End If
End Sub

' Hide the main player panel
Sub OnShutdownHandler
    If IsObject(OptionsFile) Then
        OptionsFile.Flush
    End If
    If IsObject(QuizzorMainPanel) Then
        QuizzorMainPanel.Common.Visible = False 
    End If
End Sub

Sub Uninstall 
    OptionsFile.DeleteSection "Quizzor"
    DestroyAllObjects
End Sub

