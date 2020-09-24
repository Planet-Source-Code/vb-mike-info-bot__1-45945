Attribute VB_Name = "ModReg"
Public Function SaveLW(Lw As ListView, Fname As String)
    ' The function we need for saving.
    Dim FileId As Integer
    Dim X As Integer
    Dim sIdx As Integer
    sIdx = Lw.ColumnHeaders.Count - 1
    FileId = FreeFile
    On Error Resume Next
    Open Fname For Output As #FileId
    For i = 1 To Lw.ListItems.Count
        Write #FileId, Lw.ListItems.Item(i).Text
        For X = 1 To sIdx
            Write #FileId, Lw.ListItems.Item(i).SubItems(X)
        Next
    Next
    Close #FileId
    Lw.ListItems.Clear
End Function
Public Function LoadLW(Lw As ListView, Fname As String)
    ' The function we need for loading
    Dim FileId As Integer
    Dim LVI As ListItem
    Dim fData As New Collection
    Dim Buffer As String
    Dim X As Integer
    Dim i As Integer
    Dim sIdx As Integer
    sIdx = Lw.ColumnHeaders.Count - 1
    i = 0
    FileId = FreeFile
    On Error Resume Next
    Open Fname For Input As #FileId
    While Not EOF(FileId)
        i = i + 1
        For X = 0 To sIdx
            Input #FileId, Buffer
            fData.Add Buffer
        Next
        Set LVI = Lw.ListItems.Add
        LVI.Text = fData.Item(fData.Count - sIdx)
        For Y = 1 To sIdx
            LVI.SubItems(Y) = fData.Item(fData.Count + Y - sIdx)
        Next
    Wend
    Close #FileId
    
End Function
