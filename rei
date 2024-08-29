Sub ParseTextFileAndImportAdmissions()
    Dim FilePath As String
    Dim FileContent As String
    Dim MainKeywords As Variant
    Dim Keyword As String
    Dim StartPos As Long, AdmissionsPos As Long
    Dim AdmissionsNumber As String
    Dim ws As Worksheet
    Dim ColNum As Integer

    ' Define the file path and open the file
    FilePath = "C:\path\to\your\file.txt" ' Update with your actual file path
    FileContent = GetFileContent(FilePath)
    
    ' Define the main keywords
    MainKeywords = Array("k1", "k2", "k3", "k4")
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name

    ' Loop through each main keyword
    For ColNum = 0 To UBound(MainKeywords)
        Keyword = MainKeywords(ColNum)
        
        ' Find the keyword in the file content
        StartPos = InStr(FileContent, Keyword)
        If StartPos > 0 Then
            ' Find "TOTAL ADMISSIONS:" after the keyword
            AdmissionsPos = InStr(StartPos, FileContent, "TOTAL ADMISSIONS:")
            If AdmissionsPos > 0 Then
                ' Extract the second number after "TOTAL ADMISSIONS:"
                AdmissionsNumber = ExtractSecondNumber(FileContent, AdmissionsPos)
                
                ' Import the number into the appropriate cell (B7, C7, D7, or E7)
                ws.Cells(7, ColNum + 2).Value = AdmissionsNumber
            End If
        End If
    Next ColNum
End Sub

Function GetFileContent(FilePath As String) As String
    Dim FileNum As Integer
    Dim FileContent As String
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    FileContent = Input$(LOF(FileNum), FileNum)
    Close #FileNum
    GetFileContent = FileContent
End Function

Function ExtractSecondNumber(ByVal FileContent As String, ByVal StartPos As Long) As String
    Dim RegExp As Object
    Dim Matches As Object
    Dim Result As String
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Pattern = "\b\d+(\.\d+)?\b"
    RegExp.Global = True
    Set Matches = RegExp.Execute(Mid(FileContent, StartPos))
    
    ' Skip the first number and get the second one
    If Matches.Count > 1 Then
        Result = Matches(1).Value
    Else
        Result = ""
    End If
    ExtractSecondNumber = Result
End Function