Sub koers_lijsten()
Dim SapGuiAuto As Object
Dim SAPApp As Object
Dim SAPCon As Object
Dim Session As Object
Dim iRow As Long, Icol As Long
Dim Antwoord As VbMsgBoxResult
Dim SearchString As String
Dim Zoek_Wk As String
Dim Zoek_Wp As String
Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set Session = SAPCon.Children(0) 'Get the first session (window) on that connection
Dim lrow As Long
Dim lcol As Long
Dim i As Integer
Dim pos As Integer
Dim res As String
iRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
lcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
comp_code = ActiveSheet.Range("F1").Value



End Sub