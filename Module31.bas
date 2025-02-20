Attribute VB_Name = "Module31"
Sub PTT_North()
    Application.ScreenUpdating = False
    
    Dim XMLReq As New MSXML2.XMLHTTP60
    Dim htmlDoc As New MSHTML.HTMLDocument
    Dim ws As Worksheet
    Dim url As String, fiyat As String
    Dim sonsat As Long, i As Long
    Dim urunFiyati As MSHTML.IHTMLElement
    Dim saticiAdi As MSHTML.IHTMLElement
    Dim urunAdi As MSHTML.IHTMLElement
    
    Set ws = ThisWorkbook.Sheets("PttAVM")
    sonsat = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Baþlýklarý ekleme
    With ws
        .Range("A1").Value = "Ürün Linki"
        .Range("B1").Value = "Ürün Adý"
        .Range("C1").Value = "Fiyat"
        .Range("D1").Value = "Satýcý Adý"
    End With
    
    ' "Lütfen Bekleyiniz…" mesajýný yazma
    ws.Range("F4").Value = "Lütfen Bekleyiniz…"
    
    ' Veri çekme iþlemi
    For i = 2 To sonsat
        url = Trim(ws.Range("A" & i).Value)
        
        ' URL kontrolü
        If Left(LCase(url), 8) <> "https://" Then
            ' Eðer geçerli bir URL deðilse, ilgili hücreleri boþalt
            ws.Range("B" & i & ":D" & i).ClearContents
            GoTo ContinueLoop
        End If
        
        On Error GoTo ErrorHandler
        
        ' XMLHTTP isteðini açma ve gönderme
        With XMLReq
            .Open "GET", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .send
        End With
        
        ' Ýstek cevabýný HTML belgesine dönüþtürme
        htmlDoc.body.innerHTML = XMLReq.responseText
        
        ' Ürün adýný alma
        Set urunAdi = htmlDoc.querySelector("h1")
        ws.Range("B" & i).Value = IIf(Not urunAdi Is Nothing, urunAdi.innerText, "")
        
        ' Ürün fiyatýný alma
        Set urunFiyati = htmlDoc.querySelector(".text-eGreen-700.font-semibold.md\:text-3xl.text-2xl")
        If Not urunFiyati Is Nothing Then
            fiyat = Trim(urunFiyati.innerText)
            fiyat = Replace(fiyat, "TL", "")
            ws.Range("C" & i).Value = CDbl(fiyat)
        Else
            ws.Range("C" & i).Value = ""
        End If
        
        ' Satýcý adýný alma
        Set saticiAdi = htmlDoc.querySelector("a.font-semibold")
        ws.Range("D" & i).Value = IIf(Not saticiAdi Is Nothing, saticiAdi.innerText, "")
        
        ' Ýstekler arasýna gecikme ekleme
        Application.Wait Now + TimeValue("00:00:01")
        
ContinueLoop:
    Next i
    
CleanExit:
    Set XMLReq = Nothing
    Set htmlDoc = Nothing
    Application.ScreenUpdating = True
    
    ' "Lütfen Bekleyiniz…" mesajýný temizleme
    ws.Range("F4").ClearContents
    ' Baþarý mesajý
    MsgBox "PttAVM'den güncel fiyatlar baþarý ile çekilmiþtir." & vbCrLf & vbCrLf & "Kuzey Uzun'dan sevgilerle :)", vbInformation, "Baþarýlý"
    Exit Sub
ErrorHandler:
    ws.Range("B" & i & ":D" & i).ClearContents
    Resume ContinueLoop
End Sub
