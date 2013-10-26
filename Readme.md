This is a plugin for MediaMonkey 4.0 for performing music quizzes

Usage:
- Create a playlist with songs that should be guessed and select the playlist
- A right click on the playlist allows you to randomize it
- \<Begin Quiz\> in the right-click menu starts the quiz
- A fullscreen window appears with some control elements
- Pressing \<Play\> will play the first song in the Now Playing playlist
  again
- The player will stop after the song is played and hitting \<Play\> again will restart the song from the beginning
- With \<Show Information\> the song's information is shown, specifically 'Album', 'Title', 'Artist', 'Comment' and 'file path\file name'
- \<Next\> will clear the displayed information and play the next song
- The position in the playlist is retained between restarts

Version 1.2 additions:
- Ability to display random images between songs. For enabling this go to
  the options and select Quizzor at the bottom.
- You have to add images one by one, because the MediaMonkey API doesn't
  expose multi-select dialogs yet
- The image list is randomized at the beginning and every image is only
  shown once
- Currently the list advancement is not saved, meaning all images are
  loaded again if a quiz is resumed

Localization:
Because MM doesn't support localization for scripts yet, seperate releases are
packaged for different languages.
Poedit and Grom for Windows are used to edit and generate the .po files.
A Python script replaces english strings with the localization

TODO (from source):
- Change playback time, when TrackBar changes (might need buttons)

