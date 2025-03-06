Attribute VB_Name = "Module1"
Sub UpdateOrCreateMart2025Sheet()
    Dim ws As Worksheet
    Dim shtName As String
    shtName = "Mart 2025"
    
    ' "Mart 2025" sayfasý var mý kontrol et; varsa kullan, yoksa oluþtur.
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shtName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = shtName
    End If
    
    ' ***********************************************
    ' 1. AY BÝLGÝLERÝNÝN YER ALACAÐI ALANI (K2:L4)
    ' ***********************************************
    ws.Range("K2").Value = "Güncel Ay"
    ' L2: Sistem tarihine göre "mmmm yyyy" formatýnda ay (ilk harf büyük)
    ws.Range("L2").Value = Application.WorksheetFunction.Proper(Format(Date, "mmmm yyyy"))
    
    ws.Range("K3").Value = "Gün Sayýsý"
    Dim currentYear As Long, currentMonth As Long, dayCount As Long
    currentYear = Year(Date)
    currentMonth = Month(Date)
    dayCount = Day(DateSerial(currentYear, currentMonth + 1, 0))
    ws.Range("L3").Value = dayCount
    
    ws.Range("K4").Value = "Aylýk Hedef"
    ' L4 hücresi; kullanýcý tarafýndan manuel doldurulacak. (Makro L4'ü deðiþtirmiyor.)
    
    ' ***********************************************
    ' 2. GÜNLÜK TABLO ALANININ (A, B, D, F, G, H) GÜNCELLENMESÝ
    ' ***********************************************
    Dim startRow As Long, endRow As Long
    startRow = 3
    endRow = dayCount + 2
    ' Sadece otomatik hesaplanan alanlar temizlensin; E sütunu (Gerçekleþen Ciro) silinmesin.
    ws.Range("A" & startRow & ":B" & endRow).ClearContents
    ws.Range("D" & startRow & ":D" & endRow).ClearContents
    ws.Range("F" & startRow & ":F" & endRow).ClearContents
    ws.Range("G" & startRow & ":G" & endRow).ClearContents
    ws.Range("H" & startRow & ":H" & endRow).ClearContents
    ws.Range("C" & startRow & ":C" & endRow).ClearContents
    
    ' Günlük tablonun baþlýk satýrýný (A1) güncelleyelim:
    ws.Range("A1").Value = ws.Range("L2").Value & " Günlük Hedef Daðýlýmý"
    
    Dim i As Long, r As Long, monthName As String
    ' L2'den ay ismini al (örn. "Mart")
    monthName = Split(ws.Range("L2").Value, " ")(0)
    
    For i = 1 To dayCount
        r = i + 2   ' Veriler 3. satýrdan baþlýyor.
        
        ' A sütunu: "1 Mart", "2 Mart" vb.
        ws.Cells(r, 1).Value = i & " " & monthName
        ' B sütunu: Ýlgili günün hafta adý (ilk harfi büyük)
        ws.Cells(r, 2).Value = Application.WorksheetFunction.Proper(Format(DateSerial(currentYear, currentMonth, i), "dddd"))
        
        ' C sütunu boþ býrakýlýyor.
        
        ' D sütunu: Günlük hedef (planlanan) = Aylýk Hedef/L3; L4 boþsa boþ döndür.
        ws.Cells(r, 4).FormulaLocal = "=EÐER($L$4="""" ;"""" ; $L$4/$L$3)"
        
        ' E sütunu: Gerçekleþen Ciro; makro buraya dokunmuyor, kullanýcý manuel girecek.
        
        ' F sütunu: Hedef Gerçekleþme = Eðer E boþsa boþ, yoksa E/D
        ws.Cells(r, 6).FormulaLocal = "=EÐER(E" & r & "="""" ;"""" ; E" & r & "/D" & r & ")"
        
        ' G sütunu: Hedeften Sapma = Eðer E boþsa boþ, yoksa F-1
        ws.Cells(r, 7).FormulaLocal = "=EÐER(E" & r & "="""" ;"""" ; F" & r & "-1)"
        
        ' H sütunu: Güncel Hedef hesaplama
        ' - Ýlk gün (r=3) için boþ býrak.
        ' - Diðer günlerde: Eðer L4 boþsa boþ;
        '   Eðer bugüne kadar (E3:E(r-1)) henüz hiç deðer girilmemiþse, planlanan günlük hedefi (L4/L3) yaz;
        '   Aksi halde, (Aylýk Hedef - TOPLA(E3:E(r-1))) / (L3 - (SATIR()-3))
        If r = 3 Then
            ws.Cells(r, 8).Value = ""
        Else
            ws.Cells(r, 8).FormulaLocal = "=EÐER($L$4="""" ;"""" ; EÐER(SAY($E$3:E" & (r - 1) & ")=0; $L$4/$L$3; ($L$4-TOPLA($E$3:E" & (r - 1) & "))/($L$3-(SATIR()-3))))"
        End If
    Next i
    
    ' ***********************************************
    ' 3. SAYFA BÝÇÝMLENDÝRMESÝ
    ' ***********************************************
    ' F ve G sütunlarý yüzde biçiminde (ondalýk göstermeden)
    ws.Range("F" & startRow & ":F" & endRow).NumberFormat = "0%"
    ws.Range("G" & startRow & ":G" & endRow).NumberFormat = "0%"
    
    ' E ve H sütunlarý binlik ayracý, ondalýk göstermeden
    ws.Range("D" & startRow & ":D" & endRow).NumberFormat = "#,##0"
    ws.Range("E" & startRow & ":E" & endRow).NumberFormat = "#,##0"
    ws.Range("H" & startRow & ":H" & endRow).NumberFormat = "#,##0"
    
    MsgBox "Mart 2025 sayfasý, manuel girilen veriler korunarak güncellendi.", vbInformation
End Sub

