#weekly report maker

#$filepath = "C:\Users\*****\Desktop\WeeklyActivityReport.doc"
$Word = New-Object -ComObject "word.application"
$Word.Visible = $true
$Word.Documents.Add()
$Writer = $Word.Selection

$Writer.TypeText("Weekly Activity Report")
$Writer.TypeParagraph()
$Writer.TypeText("Hours This Week: ")
$Writer.TypeParagraph()
$Writer.TypeText("Accomplishments:")
$Writer.TypeParagraph()
$Writer.TypeText("Challenges:")
$Writer.TypeParagraph()
$Writer.TypeText("Housekeeping:")
$Writer.TypeParagraph()
$Writer.TypeText("Tasks/Updates/Backlog:")

$Document.SaveAs("C:\Users\jmclanp\Desktop\WeeklyActivityReport.doc")
$Word.Quit()
