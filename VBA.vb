Option Explicit

Sub editCategory(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Type the study topic:", "Set Up Study Topic", oSh.TextFrame.TextRange.Text)
    If sText = "" Then
    Else:
        oSh.TextFrame.TextRange.Text = sText
    End If
End Sub

Sub star()
    If SlideShowWindows(1).View.Slide.Shapes("Star").Visible = False Then
        SlideShowWindows(1).View.Slide.Shapes("Star").Visible = True
    Else:
        SlideShowWindows(1).View.Slide.Shapes("Star").Visible = False
    End If
    SlideShowWindows(1).View.Slide.Shapes("Star").TextFrame.TextRange.Text = ""
End Sub

Sub revealDefinition()
    If SlideShowWindows(1).View.Slide.Shapes("RevealDefinition").Visible = True Then
        SlideShowWindows(1).View.Slide.Shapes("RevealDefinition").Visible = False
        SlideShowWindows(1).View.Slide.Shapes("Definition").Visible = True
    Else:
        SlideShowWindows(1).View.Slide.Shapes("RevealDefinition").Visible = True
        SlideShowWindows(1).View.Slide.Shapes("Definition").Visible = False
    End If
    SlideShowWindows(1).View.Slide.Shapes("RevealDefinition").TextFrame.TextRange.Text = "REVEAL"
End Sub

Sub revealTerm()
    If SlideShowWindows(1).View.Slide.Shapes("RevealTerm").Visible = True Then
        SlideShowWindows(1).View.Slide.Shapes("RevealTerm").Visible = False
        SlideShowWindows(1).View.Slide.Shapes("Term").Visible = True
    Else:
        SlideShowWindows(1).View.Slide.Shapes("RevealTerm").Visible = True
        SlideShowWindows(1).View.Slide.Shapes("Term").Visible = False
    End If
    SlideShowWindows(1).View.Slide.Shapes("RevealTerm").TextFrame.TextRange.Text = "REVEAL"
End Sub

Sub BGChange(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    ActivePresentation.Slides(7).Shapes("Green").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Green").Shadow.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Red").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Red").Shadow.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Blue").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Blue").Shadow.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Gray").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Gray").Shadow.Transparency = 1
    oSh.Fill.Transparency = 0
    oSh.Shadow.Transparency = 0.6
    Dim bgcolor As Long
    If oSh.Name = "Green" Then
        bgcolor = RGB(155, 187, 89)
    ElseIf oSh.Name = "Red" Then
        bgcolor = RGB(192, 80, 77)
    ElseIf oSh.Name = "Blue" Then
        bgcolor = RGB(75, 172, 198)
    Else:
        bgcolor = RGB(128, 128, 128)
    End If
    Dim i As Integer
    For i = 1 To ActivePresentation.Slides.Count
        ActivePresentation.Slides(i).Shapes("BackColor").Fill.ForeColor.RGB = bgcolor
    Next i
End Sub

Sub toggleShuffle(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    ActivePresentation.Slides(7).Shapes("Unstarred").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Unstarred").Shadow.Transparency = 1
    ActivePresentation.Slides(7).Shapes("All").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("All").Shadow.Transparency = 1
    oSh.Fill.Transparency = 0
    oSh.Shadow.Transparency = 0.6
End Sub

Sub toggleCover(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    ActivePresentation.Slides(7).Shapes("Definitions").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Definitions").Shadow.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Terms").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Terms").Shadow.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Random").Fill.Transparency = 1
    ActivePresentation.Slides(7).Shapes("Random").Shadow.Transparency = 1
    oSh.Fill.Transparency = 0
    oSh.Shadow.Transparency = 0.6
End Sub

Private Sub shuffleCards()
' Get number of cards
    Dim numberOfCards As Long
    numberOfCards = ActivePresentation.Slides.Count - 8
    
' Create number seed array containing flashcard slide numbers
    Dim numberSeed() As Variant
    ReDim numberSeed(numberOfCards - 1)
    Dim i As Long
    For i = 0 To UBound(numberSeed)
        numberSeed(i) = i + 9
    Next i
    
' Shuffle array in place http://www.cpearson.com/Excel/ShuffleArray.aspx
    Dim N As Long
    Dim Temp As Variant
    Dim j As Long
   
    Randomize
    For N = 0 To UBound(numberSeed)
        j = ((UBound(numberSeed) - N) * Rnd) + N
        If N <> j Then
            Temp = numberSeed(N)
            numberSeed(N) = numberSeed(j)
            numberSeed(j) = Temp
        End If
    Next N
    
' Save number seed array to slide
    ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text = "0"
    Dim k As Integer
    For k = 0 To UBound(numberSeed) - 1
        ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text = ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text + CStr(numberSeed(k)) + ","
    Next k
    ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text = ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text + CStr(numberSeed(k))
End Sub

Private Sub addShuffleNumber()
    ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text = CLng(ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text + 1)
End Sub

Sub macroCheck()
    ActivePresentation.Slides(1).Shapes("MacrosDisabled").Visible = False
    nextCard
End Sub

Sub nextCard()
    Dim theNextCard As Long
    Dim seedSplit As Variant
    
    ' Terminate if there are no flashcards
    If ActivePresentation.Slides.Count - 8 <= 0 Then
        MsgBox ("No flashcards found. Try undoing with Ctrl+Z to see if you can get a flashcard slide back. Otherwise, you'll have to re-download Flashcards for PowerPoint.")
        Exit Sub
    End If
    
    ' Form seed array
    seedSplit = Split(ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text, ",")
    
    ' If seed array is empty or all cards in seed have been used, shuffle cards
    If UBound(seedSplit) = -1 Or UBound(seedSplit) < CLng(ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text) Then
        shuffleCards
        seedSplit = Split(ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text, ",")
    End If

    ' Find the next flashcard to display, based on user setting
    
        ' If shuffle setting = unstarred, find first slide number in array not already used that is unstarred
        If ActivePresentation.Slides(7).Shapes("Unstarred").Fill.Transparency = 0 Then
            Dim foundUnstarred As Boolean
            foundUnstarred = False
            Dim shuffleCount As Integer
            shuffleCount = 0
            While foundUnstarred = False
                theNextCard = seedSplit(CLng(ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text))
                If ActivePresentation.Slides(theNextCard).Shapes("Star").Visible = True Then
                    addShuffleNumber
                    If UBound(seedSplit) < CLng(ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text) Then
                        If shuffleCount = 1 Then
                            goToCongratsStarred
                        Else:
                            shuffleCards
                            seedSplit = Split(ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text, ",")
                            shuffleCount = 1
                        End If
                    End If
                Else:
                    addShuffleNumber
                    foundUnstarred = True
                End If
            Wend
        ' If shuffle setting = All, take the first number from array
        Else:
            theNextCard = seedSplit(CLng(ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text))
            addShuffleNumber
        End If
        
        ' If PowerPoint 2007, shuffle cards here to prevent cards from accidentally revealing their terms/definitions
        If Val(Application.Version) <= 12 Then
        ' Adjust flashcard appearance, based on user setting
            ' If cover setting = definitions, cover the definitions
            If ActivePresentation.Slides(7).Shapes("Definitions").Fill.Transparency = 0 Then
                coverDefinitions (theNextCard)
            ' If cover setting = terms, cover the terms
            ElseIf ActivePresentation.Slides(7).Shapes("Terms").Fill.Transparency = 0 Then
                coverTerms (theNextCard)
            ' If cover setting = random
            Else:
                Dim randomNumber As Integer
                Randomize
                randomNumber = Int(2 * Rnd) + 1
                If randomNumber = 1 Then
                    coverTerms (theNextCard)
                Else:
                    coverDefinitions (theNextCard)
                End If
            End If
        End If
        
    SlideShowWindows(1).View.GotoSlide theNextCard
    
End Sub

Private Sub coverDefinitions(i As Long)
    ActivePresentation.Slides(i).Shapes("RevealTerm").Visible = False
    ActivePresentation.Slides(i).Shapes("Term").Visible = True
    ActivePresentation.Slides(i).Shapes("RevealDefinition").Visible = True
    ActivePresentation.Slides(i).Shapes("Definition").Visible = False
    ActivePresentation.Slides(i).Shapes("RevealDefinition").TextFrame.TextRange.Text = "REVEAL"
End Sub

Private Sub coverTerms(i As Long)
    ActivePresentation.Slides(i).Shapes("RevealTerm").Visible = True
    ActivePresentation.Slides(i).Shapes("Term").Visible = False
    ActivePresentation.Slides(i).Shapes("RevealDefinition").Visible = False
    ActivePresentation.Slides(i).Shapes("Definition").Visible = True
    ActivePresentation.Slides(i).Shapes("RevealTerm").TextFrame.TextRange.Text = "REVEAL"
End Sub

Sub goToSettings()
    settingsStarView
    SlideShowWindows(1).View.GotoSlide 7
End Sub

Sub goToCongratsStarred()
    Dim numberOfCards As Long
    numberOfCards = ActivePresentation.Slides.Count - 8
    ActivePresentation.Slides(8).Shapes("CongratsStarred").TextFrame.TextRange.Text = "Congrats, you've starred all " + CStr(numberOfCards) + " flashcards!"
    SlideShowWindows(1).View.GotoSlide 8
End Sub

Private Sub settingsStarView()
    Dim numberOfCards As Long
    numberOfCards = ActivePresentation.Slides.Count - 8
    
    Dim i As Long
    Dim numberOfStars As Long
    numberOfStars = 0
    For i = 9 To ActivePresentation.Slides.Count
        If ActivePresentation.Slides(i).Shapes("Star").Visible = True Then
            numberOfStars = numberOfStars + 1
        End If
    Next i
    
    Dim percentStarred As Double
    percentStarred = 0
    If numberOfCards > 0 Then
        percentStarred = Round(numberOfStars / numberOfCards, 2) * 100
    End If
    ActivePresentation.Slides(7).Shapes("FlashcardStats").TextFrame.TextRange.Text = CStr(numberOfCards) + " total flashcards          " + CStr(numberOfStars) + " starred (" + CStr(percentStarred) + "%)"
End Sub

Sub confirmUnstarAllCards()
    Dim unstarConfirm
    unstarConfirm = MsgBox("Are you sure you want to unstar all cards?", vbYesNo + vbDefaultButton2)
    If unstarConfirm = vbYes Then
        unstarAllCards
    Else:
        Exit Sub
    End If
End Sub

Sub unstarAllCardsAndReshuffle()
    unstarAllCards
    nextCard
End Sub

Private Sub unstarAllCards()
    Dim i As Long
    For i = 9 To ActivePresentation.Slides.Count
        ActivePresentation.Slides(i).Shapes("Star").Visible = False
    Next i
    settingsStarView
End Sub

Sub confirmResetAllCards()
    Dim resetConfirm
    resetConfirm = MsgBox("Are you ABSOLUTELY sure you want to reset all of your flashcards?" & vbNewLine & vbNewLine & _
    "Note: This will exit the slideshow and restore the default flashcards, assuming you only edited the term and definition text boxes. " & _
    "If you modified anything else, you may need to re-download Flashcards for PowerPoint.", vbYesNo + vbDefaultButton2)
    If resetConfirm = vbYes Then
        resetAllCards
    Else:
        Exit Sub
    End If
End Sub

Sub resetAllCards()
    Dim i As Long
    Dim slideCount As Long
    slideCount = ActivePresentation.Slides.Count
    
    If slideCount - 8 <= 0 Then
        MsgBox ("No flashcards found. Try undoing to see if you can get a flashcard slide back. Otherwise, you'll have to re-download.")
    End If
    
    If slideCount >= 10 Then
        For i = 10 To slideCount
            ActivePresentation.Slides(10).Delete
        Next i
    End If
    
    With ActivePresentation.Slides(9).Shapes("Term")
        .Left = 109.8644
        .Top = 52.5
        .Width = 499.2499
        .Height = 120
        .Fill.Solid
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.TextRange.Text = "Term 1"
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Font.Name = "Corbel"
        .TextFrame.TextRange.Font.Bold = True
        .TextFrame.TextRange.Font.Italic = False
        .TextFrame.TextRange.Font.Size = 36
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    End With
    
    With ActivePresentation.Slides(9).Shapes("Definition")
        .Left = 110.375
        .Top = 196.5
        .Width = 411.625
        .Height = 120
        .Fill.Solid
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.TextRange.Text = "Definition 1"
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Font.Name = "Corbel"
        .TextFrame.TextRange.Font.Bold = False
        .TextFrame.TextRange.Font.Italic = False
        .TextFrame.TextRange.Font.Size = 28
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    End With
    
    ActivePresentation.Slides(9).Shapes("Star").Visible = False

    ActivePresentation.Slides(9).Duplicate
    ActivePresentation.Slides(9).Duplicate
    ActivePresentation.Slides(10).Shapes("Term").TextFrame.TextRange.Text = "Term 2"
    ActivePresentation.Slides(11).Shapes("Term").TextFrame.TextRange.Text = "Term 3"
    ActivePresentation.Slides(10).Shapes("Definition").TextFrame.TextRange.Text = "Definition 2"
    ActivePresentation.Slides(11).Shapes("Definition").TextFrame.TextRange.Text = "Definition 3"
    
    ActivePresentation.Slides(1).Shapes("StudyTopic").TextFrame.TextRange.Text = "Type study topic here"
    
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
    ' If not PowerPoint 2007, adjust flashcard appearance here to partially fix display issues
    If Val(Application.Version) > 12 Then
        If SSW.View.CurrentShowPosition >= 9 Then
        ' Adjust flashcard appearance, based on user setting
            ' If cover setting = definitions, cover the definitions
            If ActivePresentation.Slides(7).Shapes("Definitions").Fill.Transparency = 0 Then
                coverDefinitions (SSW.View.CurrentShowPosition)
            ' If cover setting = terms, cover the terms
            ElseIf ActivePresentation.Slides(7).Shapes("Terms").Fill.Transparency = 0 Then
                coverTerms (SSW.View.CurrentShowPosition)
            ' If cover setting = random
            Else:
                Dim randomNumber As Integer
                Randomize
                randomNumber = Int(2 * Rnd) + 1
                If randomNumber = 1 Then
                    coverTerms (SSW.View.CurrentShowPosition)
                Else:
                    coverDefinitions (SSW.View.CurrentShowPosition)
                End If
            End If
        End If
    End If
End Sub

Sub OnSlideShowTerminate(oWn As SlideShowWindow)
    Dim i As Integer
    For i = 9 To ActivePresentation.Slides.Count
        ActivePresentation.Slides(i).Shapes("RevealTerm").Visible = False
        ActivePresentation.Slides(i).Shapes("Term").Visible = True
        ActivePresentation.Slides(i).Shapes("RevealDefinition").Visible = False
        ActivePresentation.Slides(i).Shapes("Definition").Visible = True
    Next i
    ActivePresentation.Slides(1).Shapes("ShuffleSeed").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(1).Shapes("ShuffleNumber").TextFrame.TextRange.Text = "0"
    ActivePresentation.Slides(7).Shapes("FlashcardStats").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(1).Shapes("MacrosDisabled").Visible = True
    ActivePresentation.Slides(8).Shapes("CongratsStarred").TextFrame.TextRange.Text = ""
End Sub
