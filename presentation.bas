Sub CreateGenerativeAIPresentation()
    ' Declare variables
    Dim pptApp As Object
    Dim pptPres As Object
    Dim slideIndex As Integer

    ' Create a new PowerPoint application and presentation
    Set pptApp = CreateObject("PowerPoint.Application")
    Set pptPres = pptApp.Presentations.Add

    ' Slide 1: Title Slide
    slideIndex = 1
    With pptPres.Slides.Add(slideIndex, ppLayoutTitle)
        .Shapes.Title.TextFrame.TextRange.Text = "Typical Use Cases for Generative AI"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Exploring Applications Across Various Industries"
    End With

    ' Slide 2: Content Creation
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Content Creation"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Text Generation: Automated article writing, blog posts, and news reports." & vbCrLf & _
        "Creative Writing: Generating poetry, short stories, and scripts." & vbCrLf & _
        "Summarization: Creating concise summaries of long documents."
    End With

    ' Slide 3: Image and Video Generation
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Image and Video Generation"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Art and Design: Creating original artworks, graphic designs, and digital content." & vbCrLf & _
        "Video Production: Generating realistic videos and animations from textual descriptions." & vbCrLf & _
        "Photo Editing: Enhancing or transforming images (e.g., colorizing black and white photos)."
    End With

    ' Slide 4: Chatbots and Virtual Assistants
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Chatbots and Virtual Assistants"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Customer Support: Providing automated responses to customer inquiries." & vbCrLf & _
        "Personal Assistants: Managing schedules, sending reminders, and handling basic tasks."
    End With

    ' Slide 5: Gaming
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Gaming"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Procedural Content Generation: Creating new levels, characters, and environments." & vbCrLf & _
        "Storytelling: Generating dynamic narratives based on player actions."
    End With

    ' Slide 6: Healthcare
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Healthcare"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Drug Discovery: Generating new molecular structures for potential medications." & vbCrLf & _
        "Medical Imaging: Enhancing or generating diagnostic images (e.g., MRI, CT scans)." & vbCrLf & _
        "Personalized Treatment Plans: Analyzing patient data to recommend treatments."
    End With

    ' Slide 7: Finance
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Finance"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Fraud Detection: Identifying unusual patterns that may indicate fraudulent activity." & vbCrLf & _
        "Algorithmic Trading: Generating trading strategies based on historical data." & vbCrLf & _
        "Risk Management: Predicting and managing financial risks."
    End With

    ' Slide 8: Education
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Education"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Personalized Learning: Creating customized educational materials for students." & vbCrLf & _
        "Automated Tutoring: Providing explanations and assistance on various subjects."
    End With

    ' Slide 9: Marketing and Advertising
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Marketing and Advertising"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Ad Copy Generation: Creating engaging and targeted advertising content." & vbCrLf & _
        "Market Analysis: Generating insights and reports from large datasets."
    End With

    ' Slide 10: Music and Audio
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Music and Audio"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Music Composition: Creating original music tracks in various genres." & vbCrLf & _
        "Sound Effects: Generating realistic sound effects for movies and games." & vbCrLf & _
        "Voice Synthesis: Producing natural-sounding speech for voiceovers and assistants."
    End With

    ' Slide 11: Software Development
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Software Development"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Code Generation: Automating parts of the coding process by generating code snippets." & vbCrLf & _
        "Automated Testing: Generating test cases and scenarios for software applications."
    End With

    ' Slide 12: Personalization
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Personalization"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Recommendations: Providing personalized product, content, or service recommendations." & vbCrLf & _
        "Customization: Generating personalized experiences for users based on their preferences."
    End With

    ' Slide 13: Scientific Research
    slideIndex = slideIndex + 1
    With pptPres.Slides.Add(slideIndex, ppLayoutText)
        .Shapes.Title.TextFrame.TextRange.Text = "Scientific Research"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Data Analysis: Generating hypotheses and insights from complex datasets." & vbCrLf & _
        "Simulation: Creating simulations to model scientific phenomena and predict outcomes."
    End With

    ' Make PowerPoint visible
    pptApp.Visible = True

    ' Clean up
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub
