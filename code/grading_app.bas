Attribute VB_Name = "Module1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Marc Niven Kumar
' Student ID: 000006972
' Date: 04-01-2024
' Program title: Assignment 5
' Description: Student Marking App
'===========================================================+

'Callback for analyzeBtn onAction
Sub main(control As IRibbonControl)
    Dim ws As Worksheet
    
    'Error Handling for existing analysis data
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Grades")
    On Error GoTo 0
    'Delete prev sheet
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    importForm.Show
End Sub

'Callback for exportBtn onAction
Sub export(control As IRibbonControl)
    
    Dim ws As Worksheet
    
    'Error handling for non existent data
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Grades")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Analyze grades first before exporting to word document.", vbExclamation
    Else
        Call createDoc
    End If

End Sub
'Callback for clearBtn onAction
Sub clearSht(control As IRibbonControl)
    Dim ws As Worksheet
    
    'Error Handling for existing analysis data
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Grades")
    On Error GoTo 0
    'Delete prev sheet
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox "Nothing to clear.", vbInformation
    End If
    
End Sub

Sub createDoc()
    Dim ws As Worksheet
    Dim wdApp As New Word.Application
    Dim wdDoc As Document
    Dim reportText As String
    Dim filePath As String
    
    wdApp.Visible = True
    wdApp.Activate
    Set wdDoc = wdApp.Documents.Add
    Set ws = ThisWorkbook.Sheets("Grades")
    
    With ws
        reportText = "Student Grade Analysis Report" & vbCrLf & vbCrLf & _
                     "The following is a report on the students’ " & .Range("B2").Value & " grades for the " & .Range("B1").Value & " course. " & _
                     "The information below was retrieved from the registrar database and analyzed in Excel." & vbCrLf & _
                     "Among the students, the highest grade is " & .Range("H3").Value & "%, achieved by " & .Range("I3").Value & ", " & _
                     "and the lowest grade is " & .Range("H4").Value & "%, achieved by " & .Range("I4").Value & ". " & _
                     "The average grade achieved is " & Round(.Range("H5").Value, 1) & "% and the standard deviation of the grades is " & Round(.Range("H6").Value, 1) & "." & vbCrLf & _
                     "The following is an analysis of all the students and their respective grades, where students " & _
                     "that failed are highlighted in red and students that scored distinctions are highlighted in green." & vbCrLf & _
                     "A histogram showing the frequency of each grade is also located below."
    
    End With
    
    'Write paragraph
    wdDoc.Content.Text = reportText
    
    'Format title
    With wdDoc.Content.Paragraphs(1).Range
        .Font.Size = 14
        .Font.Bold = True
        .ParagraphFormat.Alignment = 1
    End With
    
    
    'Paste table
    ws.Range("A1:D53").CopyPicture xlScreen, xlPicture
    With wdApp.Selection
        .EndKey Unit:=wdStory
        .TypeParagraph
        .Paste
        .ParagraphFormat.Alignment = 1
    End With
    
    'Paste freq table
    ws.Range("G2:J14").CopyPicture xlScreen, xlPicture
    With wdApp.Selection
        .EndKey Unit:=wdStory
        .TypeParagraph
        .Paste
        .ParagraphFormat.Alignment = 1
    End With
    
    'Paste histogram
    ws.ChartObjects("Chart 1").Copy
    With wdApp.Selection
        .EndKey Unit:=wdStory
        .TypeParagraph
        .TypeParagraph
        .PasteSpecial Link:=False, DataType:=wdPasteBitmap, _
        Placement:=wdInLine, DisplayAsIcon:=False
        .ParagraphFormat.Alignment = 1
    End With

    filePath = ActiveWorkbook.Path & "\Student Grade Analysis Report.docx"
    wdApp.ActiveDocument.SaveAs2 filePath
    MsgBox "Report saved to: " & filePath
    
    Set wdApp = Nothing
    Set wdDoc = Nothing

End Sub
