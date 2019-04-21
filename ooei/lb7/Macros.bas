
Sub Макрос1()
'Макрос 
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman" 
    Selection.Font.Size = 14
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
	Selection.ParagraphFormat.LineSpacing = LinesToPoints(32948)
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait 'орієнтація сторінки
        .TopMargin = CentimetersToPoints(2) 'верхнє поле
        .BottomMargin = CentimetersToPoints(2) 'нижнє
        .LeftMargin = CentimetersToPoints(3) 'зліва
        .RightMargin = CentimetersToPoints(1)'справа
        .Gutter = CentimetersToPoints(0) 
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25) 
        .PageWidth = CentimetersToPoints(21) 'ширина сторінки
        .PageHeight = CentimetersToPoints(29.7)'висота сторінки
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    
    End With
      ' ВИДАЛЕННЯ ПРОБІЛІВ

'Функція пошуку
    With Selection.Find
        .Text = " {2;}"              'шукаємо 2 і більше пробілів
        .Replacement.Text = " "      'замінюємо на один
        .Forward = True
        .Wrap = wdQuestion
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
'Повторюємо для всього тексту
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    ' ПРОБІЛ ПІСЛЯ ЗНАКУ ПУНКТУАЦІЇ
    
    'Функція пошуку
    With Selection.Find
        .Text = "([.,:;\!\?])" 'Знаходимо символи
        .Replacement.Text = "\1 "    'Замінюємо
        .Forward = True
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
'Повторюємо для всього тексту
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
	
	
	'ЗАМІНА НЕРОЗРИВНОГО ПРОБІЛУ НА ЗВИЧАЙНИЙ

'Функція пошуку
    With Selection.Find
        .Text = "^s" 'Знаходимо символи
        .Replacement.Text = " "     'Змінюємо
        .Forward = True             
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True      
    End With
'Повторюємо для всього тексту
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
	
	
'ВИДІЛЕННЯ КЛЮЧОВИХ СЛІВ
Dim range As range 
Dim i As Long 
Dim TargetList 

TargetList = Array("а", "але", "щоб", "що" , "який", "котрий") ' масив слів для пошуку

For i = 0 To UBound(TargetList) ' для довжини масиву

    Set range = ActiveDocument.range 

    With range.Find ' 
    .Text = TargetList(i) 'знайти слова,які є в масиві'
    .Format = True 
    .MatchCase = False 
    .MatchWholeWord = True 
    .MatchAllWordForms = False 

    Do While .Execute(Forward:=True) 
    range.HighlightColorIndex = wdYellow 'виділити ключові слова жовтим

    Loop 

    End With 
Next 
'СЛОВА ВИДІЛЕНІ ЖИРНИМ ШРИФТОМ ПІДКРЕСЛИТИ ДВІЧІ ТА ВИДІЛИТИ КОЛЬОРОМ
Set Content = ActiveDocument.Content'текст поточного документа
Set Words = Content.Words 
For i = 1 To Words.Count
    Set currentWord = Words.Item(i) 'слова в ньому
    If currentWord.Font.Bold = True Then 'Перевіряємо чи жирний шрифт
        currentWord.Font.Color = wdColorRed  'Змінюємо колір даного слова на червоний
        currentWord.Font.Underline = wdUnderlineDouble  'двічі підкреслюємо
    End If
Next
'КІЛЬКІСТЬ СИМВОЛІВ,АБЗАЦІВ,СТОРІНОК,СЛІВ,РЯДКІВ
Dialogs.Item(wdDialogDocumentStatistics).Display
Selection.TypeParagraph
Selection.TypeText Text:=" кількість символів " + Str(ActiveDocument.ComputeStatistics(wdStatisticCharacters)) + " кількість абзаців" + Str(ActiveDocument.ComputeStatistics(wdStatisticParagraphs)) +" кількість сторінок " + Str(ActiveDocument.ComputeStatistics(wdStatisticPages)) + " кількість слів " + Str(ActiveDocument.ComputeStatistics(wdStatisticParagraph)) + " кількість рядків " + Str(ActiveDocument.ComputeStatistics(wdStatisticLines))
End Sub


