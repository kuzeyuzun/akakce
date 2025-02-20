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
    
    ' Ba�l�klar� ekleme
    With ws
        .Range("A1").Value = "�r�n Linki"
        .Range("B1").Value = "�r�n Ad�"
        .Range("C1").Value = "Fiyat"
        .Range("D1").Value = "Sat�c� Ad�"
    End With
    
    ' "L�tfen Bekleyiniz�" mesaj�n� yazma
    ws.Range("F4").Value = "L�tfen Bekleyiniz�"
    
    ' Veri �ekme i�lemi
    For i = 2 To sonsat
        url = Trim(ws.Range("A" & i).Value)
        
        ' URL kontrol�
        If Left(LCase(url), 8) <> "https://" Then
            ' E�er ge�erli bir URL de�ilse, ilgili h�creleri bo�alt
            ws.Range("B" & i & ":D" & i).ClearContents
            GoTo ContinueLoop
        End If
        
        On Error GoTo ErrorHandler
        
        ' XMLHTTP iste�ini a�ma ve g�nderme
        With XMLReq
            .Open "GET", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .send
        End With
        
        ' �stek cevab�n� HTML belgesine d�n��t�rme
        htmlDoc.body.innerHTML = XMLReq.responseText
        
        ' �r�n ad�n� alma
        Set urunAdi = htmlDoc.querySelector("h1")
        ws.Range("B" & i).Value = IIf(Not urunAdi Is Nothing, urunAdi.innerText, "")
        
        ' �r�n fiyat�n� alma
        Set urunFiyati = htmlDoc.querySelector(".text-eGreen-700.font-semibold.md\:text-3xl.text-2xl")
        If Not urunFiyati Is Nothing Then
            fiyat = Trim(urunFiyati.innerText)
            fiyat = Replace(fiyat, "TL", "")
            ws.Range("C" & i).Value = CDbl(fiyat)
        Else
            ws.Range("C" & i).Value = ""
        End If
        
        ' Sat�c� ad�n� alma
        Set saticiAdi = htmlDoc.querySelector("a.font-semibold")
        ws.Range("D" & i).Value = IIf(Not saticiAdi Is Nothing, saticiAdi.innerText, "")
        
        ' �stekler aras�na gecikme ekleme
        Application.Wait Now + TimeValue("00:00:01")
        
ContinueLoop:
    Next i
    
CleanExit:
    Set XMLReq = Nothing
    Set htmlDoc = Nothing
    Application.ScreenUpdating = True
    
    ' "L�tfen Bekleyiniz�" mesaj�n� temizleme
    ws.Range("F4").ClearContents
    ' Ba�ar� mesaj�
    MsgBox "PttAVM'den g�ncel fiyatlar ba�ar� ile �ekilmi�tir." & vbCrLf & vbCrLf & "Kuzey Uzun'dan sevgilerle :)", vbInformation, "Ba�ar�l�"
    Exit Sub
ErrorHandler:
    ws.Range("B" & i & ":D" & i).ClearContents
    Resume ContinueLoop
End Sub
