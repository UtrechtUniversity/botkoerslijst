Sub test_valuta()
Dim wb As Workbook
Dim ws As Worksheet


 Set wb = ThisWorkbook
    Set ws = wb.Sheets("Bijgehouden_valuta's")
    row = ThisWorkbook.Sheets("Bijgehouden_valuta's").Cells(2, 1).Value
    koers_date = wb.Sheets("KoersLijst_invoeren").Range("G3").Value
    wbsource = wb.Sheets("KoersLijst_invoeren").Range("G2").Value
    
    Set wb1 = Workbooks(wbsource)
    Set ws1 = wb1.Sheets("EURO_Koerslijst")
    wb1.Activate





row = 2
Set ws = wb.Sheets("Bijgehouden_valuta's")
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).row
For i = 2 To lastrow
    r = ws.Cells(row, 1).Value
For j = 15 To ws1.Cells(Rows.Count, "m").End(xlUp).row
    If InStr(Cells(j, 13).Value, r) > 0 Then
    valuta = Cells(j, 16).Value
    valuta_per_eenheid = ws.Cells(row, 2).Value
    Invoer_koers = Application.WorksheetFunction.RoundUp(valuta * valuta_per_eenheid, 5)
'    Invoer_koers = valuta * valuta_per_eenheid
     Debug.Print r & " = " & Invoer_koers
     
     Exit For
    
    End If
    
Next j
row = row + 1
Next i
End Sub
Sub valuta_afronden()
Dim wb As Workbook
Dim ws As Worksheet


 Set wb = ThisWorkbook
 row = 2
    Set ws = wb.Sheets("Bijgehouden_valuta's")

    koers_date = wb.Sheets("KoersLijst_invoeren").Range("G3").Value
    wbsource = wb.Sheets("KoersLijst_invoeren").Range("G2").Value
    
    Set wb1 = Workbooks(wbsource)
    Set ws1 = wb1.Sheets("EURO_Koerslijst")
    wb1.Activate
    
    
        For j = 15 To ws1.Cells(Rows.Count, "m").End(xlUp).row
            r = ws.Cells(row, 1).Value
            If Cells(j, 13).Value <> Empty Then
            'valuta = Round(Cells(j, 16), 5)
            valuta = Cells(j, 16).Value
            valuta_per_eenheid = ws.Cells(row, 2).Value
            Invoer_koers = Application.WorksheetFunction.RoundDown(valuta * valuta_per_eenheid, 5)
             Debug.Print r & " = " & Invoer_koers
             
            
            row = row + 1
            End If
        Next j
        
       

End Sub

