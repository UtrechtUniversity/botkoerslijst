Sub koerslijsten1()
Set SapGuiAuto = GetObject("SAPGUI")
Set SapApplication = SapGuiAuto.GetScriptingEngine
Set Connection = SapApplication.Children(0)
Set Session = Connection.Children(0)
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
filepath = "C:\Users\khali102\OneDrive - Universiteit Utrecht\Documents\"

'Set wb = Workbooks(filepath & Filename)
row = 2

Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/n OB08"
Session.findById("wnd[0]").sendVKey 0

ws1.Activate

lastrow = ws1.Cells(Rows.Count, "p").End(xlUp).row
For i = 15 To lastrow
    r = ws.Cells(row, 1).Value
    valuta_per_eenheid = Application.WorksheetFunction.RoundDown(ws.Cells(row, 2).Value, 5)
   
    If ws1.Cells(i, 13).Value <> Empty And ws1.Cells(i, 13).Value = r And r <> Empty Then
       valuta = Cells(i, 16).Value
       valuta_per_eenheid = ws.Cells(row, 2).Value
       Invoer_koers = Application.WorksheetFunction.RoundDown(valuta * valuta_per_eenheid, 5)
        Session.findById("wnd[0]/usr/btnVIM_POSI_PUSH").press
        Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = "M"
        Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = r '"USD"
        Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").SetFocus

        Session.findById("wnd[1]/tbar[0]/btn[0]").press
        AbsoluteRow = Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR").VerticalScrollbar.Position
        Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR").getAbsoluteRow(AbsoluteRow).Selected = True
        Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-KURST[0,0]").SetFocus

        Session.findById("wnd[0]/tbar[1]/btn[6]").press
        Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-GDATU[1,0]").Text = koers_date '"27.09.2023"
        Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text = Invoer_koers '"0.94742"
        Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").SetFocus
On Error GoTo m_err:
        Session.findById("wnd[0]").sendVKey 0
        format_error = "Invoer alleen in de vorm _,___._____"
m_err:
    If Session.findById("wnd[0]/sbar").messagetype = "E" And Session.findById("wnd[0]/sbar").Text = format_error Then
        
        If Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text = Empty Then Exit Sub
        SearchString = Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text ' String to search in.
        Zoek_Wk = ","   ' Zoeken voor ","
        Zoek_Wp = "."  ' Zoeken voor "."
        pos = InStr(1, SearchString, Zoek_Wk)
        If pos = 0 Then
           pos = InStr(SearchString, Zoek_Wp)
           res = Replace(SearchString, Zoek_Wp, Zoek_Wk)
           Session.findById("wnd[0]").sendVKey 0
         Else
            res = Replace(SearchString, Zoek_Wk, Zoek_Wp)
            Session.findById("wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,0]").Text = res 'het bedrag
            Session.findById("wnd[0]").sendVKey 0 'Enter
            
            
         End If
    End If
        
        
        'On Error Resume Next
        Session.findById("wnd[0]/tbar[0]/btn[11]").press
        
'         ses[0]/wnd[0]/tbar[0]/btn[0]
         'Session.findbyid("wnd[0]/tbar[0]/btn[0]").press
                        
        
'       Debug.Print Cells(i, 13).Value & " = " & valuta
       Debug.Print r & " = " & Invoer_koers
       
       row = row + 1
    End If
'Stop:
'Session.findbyid("wnd[0]/tbar[0]/btn[12]").press
Next i





End Sub