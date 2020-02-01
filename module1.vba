Sub Format_Win()

Application.ScreenUpdating = False

'Formats the winning entry box

    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 35.25, 45.75, _
        459.75, 435).Name = "WinBox"
    ActiveSheet.Shapes("WinBox").TextFrame2.VerticalAnchor = msoAnchorMiddle
    
    'Color the win box
    If Range("Q33").Value = "Yes" Then
        'Color the tie game box
        With ActiveSheet.Shapes("WinBox").Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0.33
            .Solid
        End With
    ElseIf Range("Q21").Value = "Player" Then
        'Color the winning box
        With ActiveSheet.Shapes("WinBox").Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0.33
            .Solid
        End With
    Else
        'Color the losing box
        With ActiveSheet.Shapes("WinBox").Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0.33
            .Solid
        End With
    End If
    
    If Range("Q33").Value = "Yes" Then
        'Tie game words
        ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Characters.Text = "Tie Game"
    ElseIf Range("Q21").Value = "Player" Then
        'Winning words
        ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Characters.Text = "YOU WIN!!"
    Else
        'Losing words
        ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Characters.Text = "YOU LOSE"
    End If
    
    With ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    
    If Range("Q33").Value = "Yes" Then
        'Color the tie text blue
        With ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Font
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 255)
        End With
    ElseIf Range("Q21").Value = "Player" Then
        'Color the winning text yellow
        With ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Font
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 255, 0)
        End With
    Else
        'Color the losing text red
        With ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Font
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
    End If
    
    With ActiveSheet.Shapes("WinBox").TextFrame2.TextRange.Font
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 138
        .Name = "Mead Bold"
    End With
    
    If Range("Q33").Value = "Yes" Then
        'Color the tie lines blue
        With ActiveSheet.Shapes("WinBox").Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 255)
            .Transparency = 0
            .Visible = msoTrue
            .Weight = 4
            .Visible = msoTrue
            .DashStyle = msoLineLongDash
        End With
    ElseIf Range("Q21").Value = "Player" Then
        'Color the winning lines yellow
        With ActiveSheet.Shapes("WinBox").Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 0)
            .Transparency = 0
            .Visible = msoTrue
            .Weight = 4
            .Visible = msoTrue
            .DashStyle = msoLineLongDash
        End With
    Else
        'Color the losing lines red
        With ActiveSheet.Shapes("WinBox").Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
            .Visible = msoTrue
            .Weight = 4
            .Visible = msoTrue
            .DashStyle = msoLineLongDash
        End With
    End If

End Sub


