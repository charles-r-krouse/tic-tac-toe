
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Automatically runs whenever a box is selected

If Target.Address = "$B$3" Or _
    Target.Address = "$C$3" Or _
    Target.Address = "$D$3" Or _
    Target.Address = "$B$4" Or _
    Target.Address = "$C$4" Or _
    Target.Address = "$D$4" Or _
    Target.Address = "$B$5" Or _
    Target.Address = "$C$5" Or _
    Target.Address = "$D$5" Then

'Exit the sub if the Start Game button has not been pressed
If Range("Q10").Value = "No" Then
    MsgBox ("You must start a new game first!")
    Range("C9").Select
    Exit Sub
End If

'Turn off screen updating for a faster run time
Application.ScreenUpdating = False

    'Check if there is already an entry
    If ActiveCell.Value = "" Then
        'Run if there is no previous X or O
        'Check who's move it is
        If Range("Q16").Value = "No" Then
            'Run if it is the computer's turn
            ActiveCell.Value = Range("P1").Value
            With ActiveCell.Font
                .Name = "Calibri"
                .Size = 150
                .Bold = True
            End With
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.VerticalAlignment = xlCenter
            'Format the color
            If Range("P2").Value = "X" Then
                ActiveCell.Font.Color = -1003520
                ActiveCell.Font.TintAndShade = 0
            Else
                ActiveCell.Font.Color = -16776961
                ActiveCell.Font.TintAndShade = 0
            End If
            'Check if the computer won
            Call Test_Win
            'Confirms that the computer has made a valid move
            Range("Q18").Value = "Yes"
        Else
            'Run if it is the player's turn
            ActiveCell.Value = Range("P2").Value
            With ActiveCell.Font
                .Name = "Calibri"
                .Size = 150
                .Bold = True
            End With
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.VerticalAlignment = xlCenter
            'Format the color
            If Range("P2").Value = "X" Then
                ActiveCell.Font.Color = -16776961
                ActiveCell.Font.TintAndShade = 0
            Else
                ActiveCell.Font.Color = -1003520
                ActiveCell.Font.TintAndShade = 0
            End If
            'Check if the player won
            Call Test_Win
            'Exit the sub if the player won and don't allow the computer to move
            If Range("Q20").Value = "Yes" Then
                Exit Sub
            End If
            'Tells the computer to move after the player moves
            'Sets Valid AI move to NO and hence allows the computer to move
            Range("Q18").Value = "No"
            Call Computer_Move
        End If
    Else
        'Run if there is already an X or O in this box
        If Range("Q15").Value = "No" Then
            'If it is the player's move display an error message
            MsgBox ("You can't move here!")
        Else
            'If it is the computer's move specify that the move was invalid
            'Will force the computer to try again
            Range("Q18").Value = "No"
        End If
    End If
    
'Move the cursor off the board
Range("C9").Select

End If

End Sub

Sub New_Game()

'Runs when the New Game button is hit by the player

Dim ws As Worksheet

'Disable automatic screen updating for a faster run speed
Application.ScreenUpdating = False

'Move the cursor off the tic-tac-toe board so that all moves are possible
Range("C9").Select

'Specify that the game is in progress
Range("Q10").Value = "Yes"

'Set the number of times the computer has moved to zero
Range("Q35").Value = 0

'Set the number of winning combinations to zero
Range("Q37").Value = 0

'Clear the board
Range("B3:D5").ClearContents

'Delete the WinBox if there was a previous winner or tie game
If (Range("Q20").Value = "Yes") Or (Range("Q33").Value = "Yes") Then
    ActiveSheet.Shapes("WinBox").Delete
End If

'Declare no winner yet
Range("Q20").Value = "No"

'Declare that there is not a tied game
Range("Q33").Value = "No"

'Format the Start Game box so that the player knows a game is in progress
ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.Text = "IN PROGRESS"
With ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Font.Fill
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
    .Solid
End With
With ActiveSheet.Shapes("TextBox 1").Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 255, 0)
    .Transparency = 0
    .Solid
End With
    
'Determine who goes first and respond accordingly
If ActiveSheet.Range("Q2").Value = "Second" Then
    'Sets Valid AI move to NO and hence allows the computer to move
    Range("Q18").Value = "No"
    Call Computer_Move
End If

'Allows the player to move
Call Player_Move


End Sub


Sub Clear_Board()

'Clear the board
Range("B3:D5").ClearContents

'Delete the WinBox if there was a previous winner or tie game
If (Range("Q20").Value = "Yes") Or (Range("Q33").Value = "Yes") Then
    ActiveSheet.Shapes("WinBox").Delete
End If

'Specify that there is no game in progress
Range("Q10").Value = "No"

'Format the Start Game button so that the player knows a game is not in progress
ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.Text = "START GAME"
With ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Font.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 255, 0)
    .Transparency = 0
    .Solid
End With
With ActiveSheet.Shapes("TextBox 1").Fill
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
    .Solid
End With

'Specify that there is no winner
Range("Q20").Value = "No"

'Specify that the game is not tied
Range("Q33").Value = "No"


End Sub

Sub OButton_Is_Clicked()

'Allows the player to be X or O
'Only runs if the game is not in progress

If Range("Q10").Value = "No" Then
    ActiveSheet.Range("P2").Value = "O"
    'Color the O box
    With ActiveSheet.Shapes("TextBox 38").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With

    'Color the X box
    With ActiveSheet.Shapes("TextBox 37").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 37").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
Else
    MsgBox ("You can't change to " & Range("P1").Value & " in the middle of the game!")
End If

End Sub

Sub XButton_Is_Clicked()

'Allows the player to be X or O
'Only runs if the game is not in progress

If Range("Q10").Value = "No" Then
    ActiveSheet.Range("P2").Value = "X"
    'Color the X box
    With ActiveSheet.Shapes("TextBox 37").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    'Color the O box
    With ActiveSheet.Shapes("TextBox 38").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 38").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
'Don't allow the player to switch between X and O in the middle of a game
Else
    MsgBox ("You can't change to " & Range("P1").Value & " in the middle of the game!")
End If


End Sub

Sub Move_First()

'Only allow the player to change the button if a game is not in progress
If Range("Q10").Value = "No" Then

'Specify that the player will move first
ActiveSheet.Range("Q2").Value = "First"

'Appropriately color the "First" and "Second" boxes

    'Color the Second box
    With ActiveSheet.Shapes("TextBox 5").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 5").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
    End With
    With ActiveSheet.Shapes("TextBox 5").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    'Color the First box
    With ActiveSheet.Shapes("TextBox 2").Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 2").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    With ActiveSheet.Shapes("TextBox 2").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
Else
    'If the game is in progress then give an error message
    MsgBox ("You can't change this option while a game is in progress!")
End If

End Sub

Sub Move_Second()

'Only allow the player to change the button if a game is not in progress
If Range("Q10").Value = "No" Then

'Specify that the player will move first
ActiveSheet.Range("Q2").Value = "Second"

'Appropriately color the "First" and "Second" boxes

    'Color the First box
    With ActiveSheet.Shapes("TextBox 2").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 2").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    With ActiveSheet.Shapes("TextBox 2").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    'Color the Second box
    With ActiveSheet.Shapes("TextBox 5").Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 5").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    With ActiveSheet.Shapes("TextBox 5").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
 
Else
    'If the game is in progress then give an error message
    MsgBox ("You can't change this option while a game is in progress!")
End If

End Sub

Sub Wimpy()

'Only allow the player to change the difficulty if a game is not in progress
If Range("Q10").Value = "No" Then
    
Range("Q8").Value = "Wimpy"

    'Format the Wimpy box
    With ActiveSheet.Shapes("TextBox 13").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 13").Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    
    'Format the other two difficulty boxes
    With ActiveSheet.Shapes("TextBox 15").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 15").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 14").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 14").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
Else
    'If the game is in progress then give an error message
    MsgBox ("You can't change the difficulty in the middle of a game!")
End If

End Sub

Sub Average()

'Only allow the player to change the difficulty if a game is not in progress
If Range("Q10").Value = "No" Then

Range("Q8").Value = "Average"

    'Format the Average box
    With ActiveSheet.Shapes("TextBox 14").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 14").Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    
    'Format the other two difficulty boxes
    With ActiveSheet.Shapes("TextBox 15").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 15").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 13").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 13").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With

Else
    'If the game is in progress then give an error message
    MsgBox ("You can't change the difficulty in the middle of a game!")
End If

End Sub

Sub Impossible()

'Only allow the player to change the difficulty if a game is not in progress
If Range("Q10").Value = "No" Then

Range("Q8").Value = "Impossible"

    'Format the Impossible box
    With ActiveSheet.Shapes("TextBox 15").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 15").Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    
    'Format the other two difficulty boxes
    With ActiveSheet.Shapes("TextBox 14").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 14").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 13").TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.Shapes("TextBox 13").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With

Else
    'If the game is in progress then give an error message
    MsgBox ("You can't change the difficulty in the middle of a game!")
End If

End Sub


Sub RandomNumber()

'Only allow the computer to move in the corners or middle on it's first move
'Other moves will allow for the AI to lose
'Only applies when difficulty is set to Impossible
If Range("Q8").Value = "Impossible" And (Range("Q35").Value = 0 Or Range("Q35").Value = 1) Then
    ActiveSheet.Range("Q12").Value = CInt((8 * Rnd))
    Do Until ActiveSheet.Range("Q12").Value Mod 2 = 0
        ActiveSheet.Range("Q12").Value = CInt((8 * Rnd))
    Loop
Else
    'If the difficulty is not set to Impossible then any random number will suffice
    ActiveSheet.Range("Q12").Value = CInt((8 * Rnd))
End If

End Sub

Sub Computer_Move()

Application.ScreenUpdating = False

Dim iii As Integer
Dim jjj As Integer
Dim nnn As Integer
Dim nCol As Integer
Dim nRow As Integer
Dim bRandomMove As Boolean
Dim ppp As Integer

'Specify that it is not the player's move
Range("Q16").Value = "No"

'Will run the random move algorithm if a 'smarter' move is not found
bRandomMove = True

'Reset to all possible moves
For iii = 0 To 8
    Range("Q23").Offset(iii, 0).Value = "No"
Next iii

'Only run for the Impossible difficutly selecion
'Selects the middle cell so that the computer won't lose
'Only applicable if the computer moves second and the middle cell is blank
If Range("Q8").Value = "Impossible" And Range("C4").Value = "" And _
    (Range("Q1").Value = "Second" Or Range("Q35").Value > 0) Then
    Range("C4").Select
End If

'Only run for the Average and Impossible difficult cases
If (Range("Q8").Value = "Average" Or Range("Q8").Value = "Impossible") And Range("Q18").Value = "No" Then
    'First check if the computer can win
    'Secondly check if the player can be blocked
    For ppp = 0 To 1
    'Don't need to check the last cell
    For jjj = 0 To 7
        'Loop to each square and determine if it has an entry
        'Only runs if there is an X or O entry
        If Range("B3").Offset(jjj \ 3, jjj Mod 3).Value = Range("P1").Offset(ppp, 0).Value Then
            'Determine the appropriate cells to test for possible three in a row
            For nnn = 0 To 8
                'Find the appropriate column and check for entries with an X
                'An X designates a possible 3 in a row
                'Check if the square which is being tested matches the other cell whick may result in three in a row
                If Range("T8").Offset((nnn + 1), jjj).Value = "X" And _
                    Range("B3").Offset(jjj \ 3, jjj Mod 3).Value = _
                    Range("B3").Offset(nnn \ 3, nnn Mod 3).Value Then
                        'If a win can be achieved or blocked then do so
                        'Run if the move is vertical
                        If (jjj Mod 3) = (nnn Mod 3) Then
                            nCol = jjj Mod 3
                            'Calculates the appropriate row by recognizing that the sum of "cell"\3 will always equal 3 for three in a row
                            nRow = 3 - ((((Range("T8").Offset(0, jjj).Value) - 1) \ 3) + (((Range("T8").Offset(0, nnn).Value) - 1) \ 3))
                        'Run if the move is horizontal
                        ElseIf (jjj \ 3) = (nnn \ 3) Then
                            nRow = jjj \ 3
                            nCol = 3 - (((Range("T8").Offset(0, jjj).Value) - 1) Mod 3) - (((Range("T8").Offset(0, nnn).Value) - 1) Mod 3)
                        'Run if the move is diagonal
                        Else
                            nRow = 3 - ((((Range("T8").Offset(0, jjj).Value) - 1) \ 3) + (((Range("T8").Offset(0, nnn).Value) - 1) \ 3))
                            nCol = 3 - (((Range("T8").Offset(0, jjj).Value) - 1) Mod 3) - (((Range("T8").Offset(0, nnn).Value) - 1) Mod 3)
                        End If
                        'Selects the appropriate cell if there is no other entry
                        If Range("B3").Offset(nRow, nCol).Value = "" Then
                            Range("B3").Offset(nRow, nCol).Select
                            'Prevents the code to continue to looping and looking for more moves
                            jjj = 8
                            nnn = 9
                            ppp = 2
                            bRandomMove = False
                        End If
                End If
            Next nnn
        End If
    Next jjj
    Next ppp
End If
    
'Runs only if a previous 'smarter' move was not made
'Makes a completely random move
If bRandomMove = True Then
    
    'Loop until the computer moves to a blank space
    'Or loop until there are no more available moves
    Do Until Range("Q18").Value = "Yes" Or (Range("Q23").Value = "Yes" And _
        Range("Q23") = Range("Q24") And _
        Range("Q23") = Range("Q25") And _
        Range("Q23") = Range("Q26") And _
        Range("Q23") = Range("Q27") And _
        Range("Q23") = Range("Q28") And _
        Range("Q23") = Range("Q29") And _
        Range("Q23") = Range("Q30") And _
        Range("Q23") = Range("Q31"))
    
        Call RandomNumber
    
        For jjj = 0 To 8
            'Determine what the random number is
            'Select the appropriate cell on the game board
            If Range("Q12").Value = jjj Then
                Range("Q23").Offset(jjj, 0).Value = "Yes"
                Range("B3").Offset(jjj \ 3, jjj Mod 3).Select
            End If
        Next jjj
    
    Loop

    'If there are no more moves left then specify that the game is over
    If (Range("Q23").Value = "Yes" And _
        Range("Q23") = Range("Q24") And _
        Range("Q23") = Range("Q25") And _
        Range("Q23") = Range("Q26") And _
        Range("Q23") = Range("Q27") And _
        Range("Q23") = Range("Q28") And _
        Range("Q23") = Range("Q29") And _
        Range("Q23") = Range("Q30") And _
        Range("Q23") = Range("Q31")) Then
        'Specify that the game is over
        Range("Q10").Value = "No"
        'Format the Start Game button so that the player knows a game is not in progress
        ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.Text = "START GAME"
        With ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 0)
            .Transparency = 0
            .Solid
        End With
        With ActiveSheet.Shapes("TextBox 1").Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        'Confirm that the game ended in a tie
        Range("Q33").Value = "Yes"
        Range("Q20").Value = "Yes"
        'Display the tie box
        Call Format_Win
    End If
    
    'Ends the random move algorithm
    End If
    
    'Keep a tally of many times the computer has moved
    'If the computer moves 5 times and there is still no winner, the game has ended in a tie
    Range("Q35").Value = (Range("Q35").Value + 1)
    If (Range("Q35").Value = 5) And (Range("Q20").Value = "No") Then
        'Specify that the game is over
        Range("Q10").Value = "No"
        'Format the Start Game button so that the player knows a game is not in progress
        ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.Text = "START GAME"
        With ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 0)
            .Transparency = 0
            .Solid
        End With
        With ActiveSheet.Shapes("TextBox 1").Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        'Confirm that the game ended in a tie
        Range("Q33").Value = "Yes"
        Call Format_Win
    End If

'Allows the player to move
Call Player_Move
    
End Sub

Sub Player_Move()

'Specify that it is the player's turn
Range("Q16").Value = "Yes"

End Sub

Sub Test_Win()

Application.ScreenUpdating = False

'Test if there is horizontal Tic-Tac-Toe
If (Range("B3").Value = Range("C3").Value) And _
    (Range("B3").Value = Range("D3").Value) And _
    Range("B3").Value <> "" And _
    Range("C3").Value <> "" And _
    Range("D3").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("B3").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

If (Range("B4").Value = Range("C4").Value) And _
    (Range("B4").Value = Range("D4").Value) And _
    Range("B4").Value <> "" And _
    Range("C4").Value <> "" And _
    Range("D4").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("B4").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

If (Range("B5").Value = Range("C5").Value) And _
    (Range("B5").Value = Range("D5").Value) And _
    Range("B5").Value <> "" And _
    Range("C5").Value <> "" And _
    Range("D5").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("B5").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

'Test if there is vertical Tic-Tac_Toe
If (Range("B3").Value = Range("B4").Value) And _
    (Range("B3").Value = Range("B5").Value) And _
    Range("B3").Value <> "" And _
    Range("B4").Value <> "" And _
    Range("B5").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("B3").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

If (Range("C3").Value = Range("C4").Value) And _
    (Range("C3").Value = Range("C5").Value) And _
    Range("C3").Value <> "" And _
    Range("C4").Value <> "" And _
    Range("C5").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("C3").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

If (Range("D3").Value = Range("D4").Value) And _
    (Range("D3").Value = Range("D5").Value) And _
    Range("D3").Value <> "" And _
    Range("D4").Value <> "" And _
    Range("D5").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("D3").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

'Test if there is diagonal Tic-Tac-Toe
If (Range("B3").Value = Range("C4").Value) And _
    (Range("B3").Value = Range("D5").Value) And _
    Range("B3").Value <> "" And _
    Range("C4").Value <> "" And _
    Range("D5").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("B3").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

'Only allow the following code to run ONCE
'Prevents a win or loss in multiple directions
If Range("Q37").Value > 0 Then
    Exit Sub
End If

If (Range("B5").Value = Range("C4").Value) And _
    (Range("B5").Value = Range("D3").Value) And _
    Range("B5").Value <> "" And _
    Range("C4").Value <> "" And _
    Range("D3").Value <> "" Then
    'Keep a tally of how many times there is 3 in a row
    Range("Q37").Value = (Range("Q37").Value) + 1
    If Range("B5").Value = Range("P2").Value Then
        Call You_Win
    Else
        Call You_Lose
    End If
End If

Application.ScreenUpdating = True

End Sub

Sub You_Win()

'Confirm that there is a winner and the game is over
Range("Q20").Value = "Yes"
Range("Q10").Value = "No"
'Format the Start Game button so that the player knows a game is not in progress
ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.Text = "START GAME"
With ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Font.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 255, 0)
    .Transparency = 0
    .Solid
End With
With ActiveSheet.Shapes("TextBox 1").Fill
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
    .Solid
End With

'Specify that the player has won
Range("Q21").Value = "Player"
Call Format_Win

End Sub

Sub You_Lose()

'Confirm that there is a winner and the game is over
Range("Q20").Value = "Yes"
Range("Q10").Value = "No"
'Format the Start Game button so that the player knows a game is not in progress
ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.Text = "START GAME"
With ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Font.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 255, 0)
    .Transparency = 0
    .Solid
End With
With ActiveSheet.Shapes("TextBox 1").Fill
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
    .Solid
End With

'Specify that the player has lost
Range("Q21").Value = "AI"
Call Format_Win

End Sub

Sub testing()

Range("D3").Select

End Sub


