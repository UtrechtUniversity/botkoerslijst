
Private Sub MoveFiles_Click()
bestanden_verplaatsen
End Sub

Private Sub OpenFiles_Click()
answer = MsgBox("SAP UUP dient ingelogd te zijn", vbQuestion + vbYesNo + vbDefaultButton2, "SAP GUI LOGON")
If answer = vbYes Then
ActiveSheet.Range("C19:C22").ClearContents
Invoergegevens.Show
Else
Exit Sub

 
End If
 
End Sub


