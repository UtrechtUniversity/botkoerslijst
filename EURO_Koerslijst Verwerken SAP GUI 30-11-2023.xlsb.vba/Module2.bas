Sub test_filepath()
filepath = ThisWorkbook.Sheets("KoersLijst_invoeren").Range("G4").Value

MsgBox filepath
If Right(filepath, 1) = "\" Then
   ActiveSheet.Save filepath & fileName
Else
ActiveSheet.Save filepath & "\" & fileName
End If
End Sub
