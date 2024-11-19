Sub koerslijsten_verwerken_in_SAP()
Set SapGuiAuto = GetObject("SAPGUI")
Set SapApplication = SapGuiAuto.GetScriptingEngine
Set Connection = SapApplication.Children(0)
Set SAPSession = Connection.Children(0)
Dim wb1 As Workbook
Dim row As String
Application.DisplayAlerts = False


 Set wb = ThisWorkbook
    Set ws = wb.Sheets("Bijgehouden_valuta's")
    row = ThisWorkbook.Sheets("Bijgehouden_valuta's").Cells(2, 1).Value
    koers_date = wb.Sheets("KoersLijst_invoeren").Range("G3").Value
    wbsource = wb.Sheets("KoersLijst_invoeren").Range("G2").Value
    
    Set wb1 = Workbooks(wbsource)
    Set ws1 = wb1.Sheets("EURO_Koerslijst")

ws1.Activate
row = 2
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).row
SAPSession.findById("wnd[0]").maximize
SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/n OB08"
SAPSession.findById("wnd[0]").sendVKey 0

For i = 2 To lastrow
    r = ws.Cells(row, 1).Value
For j = 15 To ws1.Cells(Rows.Count, "m").End(xlUp).row
    If InStr(Cells(j, 13).Value, r) > 0 Then
     Debug.Print r & " = " & Cells(j, 16)
     Exit For
    
    End If
Next j

'SAP verwerken
    valuta = Cells(j, 16).Value
    valuta_per_eenheid = ws.Cells(row, 2).Value
    Invoer_koers = Application.WorksheetFunction.Round(valuta * valuta_per_eenheid, 5)
    SAPSession.findById("wnd[0]/usr/btnVIM_POSI_PUSH").press
    SAPSession.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = "M"
    SAPSession.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = r '"USD"
    SAPSession.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").SetFocus

    SAPSession.findById("wnd[1]/tbar[0]/btn[0]").press
    AbsoluteRow = SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR").VerticalScrollbar.Position
    SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR").getAbsoluteRow(AbsoluteRow).Selected = True
    SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-KURST[0,0]").SetFocus

    SAPSession.findById("wnd[0]/tbar[1]/btn[6]").press
    SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-GDATU[1,0]").Text = koers_date '"27.09.2023"
    SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text = Invoer_koers '"0.94742"
    SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").SetFocus
On Error GoTo m_err:
        SAPSession.findById("wnd[0]").sendVKey 0
        format_error = "Invoer alleen in de vorm _,___._____"
m_err:
    If SAPSession.findById("wnd[0]/sbar").messagetype = "E" And SAPSession.findById("wnd[0]/sbar").Text = format_error Then
        
        If SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text = Empty Then Exit Sub
        SearchString = SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text ' String to search in.
        Zoek_Wk = ","   ' Zoeken voor ","
        Zoek_Wp = "."  ' Zoeken voor "."
        pos = InStr(1, SearchString, Zoek_Wk)
        If pos = 0 Then
           pos = InStr(SearchString, Zoek_Wp)
           res = Replace(SearchString, Zoek_Wp, Zoek_Wk)
           SAPSession.findById("wnd[0]").sendVKey 0
         Else
            res = Replace(SearchString, Zoek_Wk, Zoek_Wp)
            SAPSession.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text = res 'het bedrag
            SAPSession.findById("wnd[0]").sendVKey 0 'Enter
            
            
         End If
    End If
        
        
        'On Error Resume Next
        SAPSession.findById("wnd[0]/tbar[0]/btn[11]").press
        
'         ses[0]/wnd[0]/tbar[0]/btn[0]
         'SAPSession.findbyid("wnd[0]/tbar[0]/btn[0]").press
                        
        
'       Debug.Print Cells(i, 13).Value & " = " & valuta
       Debug.Print r & " = " & Invoer_koers
       
       
    
row = row + 1
Next i
MsgBox "SCRIPT IS KLAAR"
'Hier komt code voor afsluiten koerslijst bestand
End Sub