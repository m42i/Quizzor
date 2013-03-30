' Monkey Media Quizzor plugin
' Features:
' O Only show track info if space bar is pressed
' O Only show tracknumber and length in playlist
' O Stop after current track
' O Play track again if length < 60s
' O Keep track of correctly guessed tracks

Sub NewQuiz(Item)
  ' Ask if a new quiz should really be started
  createNew = SDB.MessageBox( SDB.Localize("Creating a new quiz replaces all  tracks in the current queue. This cannot be undone. Do you want to create a new quiz and lose the old quiz?"), mtWarning, Array(mbNo, mbYes))

  If createNew = mrNo then 
    Exit Sub 
  End If

  ' Replace playing queue with current tracks from main window 
  Call SDB.Player.PlaylistClear()

  ' Randomize playing queue
  ' Create new empty playlist, for played tracks
  ' Select newly created playlist
End Sub

Sub StartQuiz(Item)
End Sub

Sub StopQuiz(Item)
End Sub

Sub OnStartup
  Set UI = SDB.UI

  ' Register new menu item in tools menu
  Set QuizBar = UI.AddToolbar("QuizBar")
  Set SDB.Objects("QuizBar") = QuizBar

  Set NewQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
  NewQuizBtn.Caption = "New Quiz"

  Set StartQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
  StartQuizBtn.Caption = "Start Quiz"
  
  Set StopQuizBtn = UI.AddMenuItem(QuizBar, 0, -1)
  StopQuizBtn.Caption = "Start Quiz"

  Script.RegisterEvent NewQuizBtn, "OnClick", "NewQuiz"
  Script.RegisterEvent StartQuizBtn, "OnClick", "StartQuiz"
  Script.RegisterEvent StopQuizBtn, "OnClick", "StopQuiz"

End Sub

Sub Uninstall 
  Set QuizBar = SDB.Objects("QuizBar")
  if Not QuizBar Is Nothing then
    QuizBar.Visible = False
    Set QuizBar = Nothing
  end if
End Sub


