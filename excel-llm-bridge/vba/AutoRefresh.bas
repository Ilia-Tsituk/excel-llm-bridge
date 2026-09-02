Attribute VB_Name = "Module1"
' AutoRefresh.bas
'
' Watches cell I17 for a new question, writes it to question.txt, then
' polls answer.txt every 3 seconds and displays the result in cell I18.
' Meant to be paired with the Python watcher script (python/watcher.py).
'
' Setup:
'   1. Open the VBA editor (Alt+F11) in your workbook.
'   2. Import this module (or paste its contents into a new module).
'   3. Update `filePath` below to match the folder used by watcher.py
'      (this must match the BRIDGE_FOLDER value on the Python side).
'   4. Run AutoRefresh() once (e.g. from a button or Workbook_Open) to
'      start the polling loop. Run StopRefresh() to stop it.

Dim nextRun As Date
Dim lastQuestion As String

Sub AutoRefresh()
    Dim question As String
    Dim filePath As String
    Dim fileNum As Integer

    filePath = "C:\Users\hello\Desktop\"
    question = CStr(ThisWorkbook.Sheets(1).Range("I17").Value)

    ' Write the question to a file if it has changed
    If question <> "" And question <> lastQuestion Then
        fileNum = FreeFile
        Open filePath & "question.txt" For Output As #fileNum
        Print #fileNum, question
        Close #fileNum
        lastQuestion = question
        ' Clear the previous answer
        ThisWorkbook.Sheets(1).Range("I18").Value = "..."
        ' Remove the old answer.txt
        If Dir(filePath & "answer.txt") <> "" Then
            Kill filePath & "answer.txt"
        End If
    End If

    ' Read the answer from the file
    If Dir(filePath & "answer.txt") <> "" Then
        fileNum = FreeFile
        Dim line As String
        Dim fullText As String
        fullText = ""
        Open filePath & "answer.txt" For Input As #fileNum
        Do While Not EOF(fileNum)
            Line Input #fileNum, line
            fullText = fullText & line & Chr(10)
        Loop
        Close #fileNum
        If fullText <> "" Then
            ThisWorkbook.Sheets(1).Range("I18").Value = fullText
        End If
    End If

    nextRun = Now + TimeSerial(0, 0, 3)
    Application.OnTime nextRun, "AutoRefresh"
End Sub

Sub StopRefresh()
    On Error Resume Next
    Application.OnTime nextRun, "AutoRefresh", , False
End Sub
