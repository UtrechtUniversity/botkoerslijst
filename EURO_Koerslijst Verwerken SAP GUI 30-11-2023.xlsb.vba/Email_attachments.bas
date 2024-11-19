Sub Treasury_Email_Account()
Dim fso As Scripting.FileSystemObject
Dim FolderPath
Dim OlderFolderPath
Dim ol As Outlook.Application
Dim Ns As Outlook.Namespace
Dim get_Folder As folder
Dim olAccount As Object
Dim olmail As Outlook.MailItem
'FolderPath = ActiveSheet.Range("G4").Value
 
'Set fso = New Scripting.FileSystemObject
'If 'Not fso.FolderExists(FolderPath) Then

'End If
 

 
Set ol = Outlook.Application
Set Ns = Outlook.GetNamespace("MAPI")
Set get_Folder = Ns.GetDefaultFolder(olFolderInbox).Folders("WKR")

 
End Sub