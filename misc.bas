Attribute VB_Name = "Module1"
Public Sub ReadList(List As ListBox, Filename As String, Optional ClearList As Boolean)
    On Error GoTo Err
    Open Filename For Input As #1
    If ClearList = True Then List.Clear


    Do While Not EOF(1)
        Input #1, lstinput
        List.AddItem lstinput
    Loop
    Close #1
    Exit Sub
Err:
    MsgBox "Error in ReadList" & Chr(13) & Chr(13) & Err.Number _
    & " - " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub


Public Sub WriteList(List As ListBox, Filename As String)


    If List.ListCount <= 0 Then
        MsgBox "Listbox is empty - cannot write to file!", vbCritical, "Error"
        End
    End If
    On Error GoTo Err
    Open Filename For Output As #1


    For I = 0 To List.ListCount - 1
        Print #1, List.List(I)
    Next
    Close #1
    Exit Sub
Err:
    MsgBox "Error in WriteList" & Chr(13) & Chr(13) & Err.Number _
    & " - " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub
