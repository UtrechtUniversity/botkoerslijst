
Private Sub Annuleren_Click()
Unload Invoergegevens


End Sub

Private Sub ok_Click()
ActiveSheet.Range("G3").Value = Format(Me.TB1.Value, "dd" & "." & "mm" & "." & "yyyy")
Unload Invoergegevens
Call Open_file
Call koerslijsten_verwerken_in_SAP

End Sub

Private Sub TB1_Change()

End Sub

Private Sub UserForm_Click()

End Sub