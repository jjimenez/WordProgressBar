Attribute VB_Name = "Module1"

Dim progressCheckTarget As Double
Dim runWhen As Double


Sub ProgressCheck()
    
    progressCheckTarget = CDbl(InputBox("Enter target wordcount", "Target", "10000"))
    NumberOfWords
    
End Sub

Sub StopProgressCheck()
    progressCheckTarget = 0
End Sub

Sub NumberOfWords()
    Dim lngWords As Double
    Dim progress As Double
    Dim done As Integer
    Dim report As String
    Dim breaks As Integer
    
    breaks = 50
    
    With Word.Application
        If .Windows.Count > 0 And progressCheckTarget > 0 Then
            lngWords = ActiveDocument.Content.Words.Count
            progress = lngWords / progressCheckTarget * 100#
            done = CInt(progress) / (100 / breaks)
            If done > breaks Then done = breaks
            report = Format(progress, "##0.00") & "% of Target  [" & String(done, "I") & String(breaks - done, " ") & "]"
            .StatusBar = report
            runWhen = Now + TimeValue("00:00:20")
            .OnTime runWhen, "NumberOfWords"
        End If
    End With

End Sub
