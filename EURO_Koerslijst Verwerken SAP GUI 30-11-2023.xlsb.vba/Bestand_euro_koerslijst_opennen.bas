'Sub CheckFiles()
'Dim fso As Object
'Dim FolderPath As Object
'
'Dim Testworbook As Workbook
'
'Set Testworkbook = Nothing
'Set fso = CreateObject("Scripting.filesystemObject")
'Set FolderPath = fso.GetFolder("C:\Users\khali102\OneDrive - Universiteit Utrecht\Documents\koersLijst")
'Set Files = FolderPath.Files
''If fso.FolderExists(FolderPath) Then
'    For Each File In Files
'        Debug.Print fso.GetExtensionName(File)
'
'    Next File
'
'
'End Sub
Sub Open_file_oud()
Dim fso As Object
Dim NewFolderPath
Dim OlderFolderPath
OlderFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G5").Value 'map oud_verwerkt of afgehandeld
NewFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value 'map koerslijst verwerken
Set fso = CreateObject("Scripting.FileSystemObject")

'FolderPath = fso.GetFolder("C:\Users\khali102\OneDrive - Universiteit Utrecht\Documents\koersLijst")
If Not fso.FolderExists(NewFolderPath) Then
    fso.CreateFolder NewFolderPath
    

End If

If fso.FolderExists(NewFolderPath) Then

Set NewFolderPath = fso.GetFolder(ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value) 'map oud_verwerkt of afgehandeld
Set Files = NewFolderPath.Files
For Each file In Files
    
    Set file = Workbooks.Open(file)
    Set wbname = ActiveWorkbook
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("KoersLijst_invoeren")
    ws.Activate
    ws.Range("G1").Value = wbname.Sheets(1).Range("K1").Value
    ws.Range("G2").Value = file.Name
Next file


Set fso = Nothing

End If

End Sub
Sub bestanden_verplaatsen01()

Dim wbtoMove As Workbook
For Each wb In Workbooks
    If wb.Name = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G2") Then
        wb.Close
        
    End If

Next wb

Dim sourceFolderPath As String, destinationFolderPath As String
Dim fso As Object, sourceFolder As Object, file As Object
Dim fileName As String, sourceFilePath As String, destinationFilePath As String

Application.ScreenUpdating = False

sourceFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value
destinationFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G5").Value

Set fso = CreateObject("Scripting.FileSystemObject")
Set sourceFolder = fso.GetFolder(sourceFolderPath)

For Each file In sourceFolder.Files
    fileName = file.Name
    If InStr(fileName, ".xls") Or InStr(fileName, ".xlsx") Then ' Only xlsx files will be moved
        sourceFilePath = file.Path
        
        destinationFilePath = destinationFolderPath & "\" & fileName
        fso.MoveFile Source:=sourceFilePath, Destination:=destinationFilePath
    End If ' If InStr(sourceFileName, ".xlsx") Then' Only xlsx files will be moved
Next


Set sourceFolder = Nothing
Set fso = Nothing

End Sub
Sub bestanden_verplaatsen()

    Dim wbtoMove As Workbook
    Dim wb As Workbook
    Dim sourceFolderPath As String, destinationFolderPath As String
    Dim fso As Object, sourceFolder As Object, file As Object
    Dim fileName As String, sourceFilePath As String, destinationFilePath As String

    Application.ScreenUpdating = False

    ' Ensure paths are correctly formatted with trailing backslashes
    sourceFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value
    destinationFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G5").Value

    If Right(sourceFolderPath, 1) <> "\" Then sourceFolderPath = sourceFolderPath & "\"
    If Right(destinationFolderPath, 1) <> "\" Then destinationFolderPath = destinationFolderPath & "\"

   

    ' Close the workbook specified in the "G2" cell if it is open
    For Each wb In Workbooks
        If wb.Name = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G2").Value Then
            wb.Close False ' Close without saving changes
        End If
    Next wb

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrorHandler
    ' Validate if source folder exists
    If Not fso.FolderExists(sourceFolderPath) Then
        MsgBox "Source folder does not exist: " & sourceFolderPath
        GoTo Cleanup
    End If

    ' Validate if destination folder exists
    If Not fso.FolderExists(destinationFolderPath) Then
        MsgBox "Destination folder does not exist: " & destinationFolderPath
        GoTo Cleanup
    End If

    Set sourceFolder = fso.GetFolder(sourceFolderPath)

    For Each file In sourceFolder.Files
        fileName = file.Name
        If InStr(fileName, ".xls") > 0 Or InStr(fileName, ".xlsx") > 0 Then ' Only xls/xlsx files will be moved
            sourceFilePath = file.Path
            destinationFilePath = destinationFolderPath & fileName
            ' Debugging: Print the file paths to the Immediate Window
'            Debug.Print "Moving file: " & sourceFilePath & " to " & destinationFilePath
            fso.MoveFile Source:=sourceFilePath, Destination:=destinationFilePath
        End If
    Next file

    On Error GoTo 0

Cleanup:
    Set sourceFolder = Nothing
    Set fso = Nothing
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    Resume Cleanup

End Sub


Sub Open_file_01()
    Dim fso As Object
    Dim NewFolderPath As String
    Dim OlderFolderPath As String
    Dim NewFolder As Object
    Dim file As Object
    Dim wbname As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    
    OlderFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G5").Value 'map oud_verwerkt of afgehandeld
    NewFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value 'map koerslijst verwerken
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(NewFolderPath) Then
        fso.CreateFolder NewFolderPath
    End If

    If fso.FolderExists(NewFolderPath) Then
        Set NewFolder = fso.GetFolder(NewFolderPath)
        If NewFolder.Files.Count > 0 Then
            ' Assuming there's always one file
            For Each file In NewFolder.Files
                Set wbname = Workbooks.Open(file.Path)
                Set wb = ThisWorkbook
                Set ws = wb.Sheets("KoersLijst_invoeren")
                ws.Activate
                ws.Range("G1").Value = wbname.Sheets(1).Range("K1").Value
                ws.Range("G2").Value = file.Name
'                wbname.Close False ' Close the opened workbook without saving changes
                Exit For ' Exit the loop after processing the first file
            Next file
        End If
    End If

    Set fso = Nothing
End Sub
Sub Open_file()
    Dim fso As Object
    Dim NewFolderPath As String
    Dim OlderFolderPath As String
    Dim NewFolder As Object
    Dim file As Object
    Dim wbname As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileExtension As String
    
    OlderFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G5").Value 'map oud_verwerkt of afgehandeld
    NewFolderPath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value 'map koerslijst verwerken
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(NewFolderPath) Then
        fso.CreateFolder NewFolderPath
    End If

    If fso.FolderExists(NewFolderPath) Then
        Set NewFolder = fso.GetFolder(NewFolderPath)
        If NewFolder.Files.Count > 0 Then
            ' Assuming there's always one file
            For Each file In NewFolder.Files
                fileExtension = fso.GetExtensionName(file.Path)
                If LCase(fileExtension) = "xlsx" Or LCase(fileExtension) = "xls" Then
                    Set wbname = Workbooks.Open(file.Path)
                    Set wb = ThisWorkbook
                    Set ws = wb.Sheets("KoersLijst_invoeren")
                    ws.Activate
                    ws.Range("G1").Value = wbname.Sheets(1).Range("K1").Value
                    ws.Range("G2").Value = file.Name
                    
                    Exit For ' Exit the loop after processing the first file
                End If
            Next file
        End If
    End If

    Set fso = Nothing
End Sub
