Attribute VB_Name = "Module2"
Sub ScrapePTTAVMProductLinks_Test()
    Dim xhr As Object
    Dim htmlDoc As Object
    Dim productElements As Object
    Dim i As Long
    Dim productURL As String
    Dim productName As String
    
    ' URL'yi tanýmla
    Dim url As String
    url = "https://www.pttavm.com/kampanyalar/ayin-en-cok-satan-tvleri/ses-sistemleri"
    
    ' Internet Explorer için HTML belge nesnesini oluþtur
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    xhr.Open "GET", url, False
    xhr.send
    
    Set htmlDoc = CreateObject("HTMLFile")
    htmlDoc.body.innerHTML = xhr.responseText
    
    ' Farklý HTML etiketlerini ve sýnýflarý kontrol et
    Set productElements = htmlDoc.getElementsByClassName("product-card")
    
    If productElements.Length = 0 Then
        Set productElements = htmlDoc.getElementsByTagName("a")
    End If
    
    ' Aktif çalýþma sayfasýný seç
    Worksheets("Ürünler").Activate
    
    ' Baþlýk satýrý
    Cells(1, 1).Value = "Ürün Adý"
    Cells(1, 2).Value = "Ürün Linki"
    
    ' Ürün linklerini Excel'e aktar
    For i = 0 To productElements.Length - 1
        On Error Resume Next
        productURL = productElements(i).getAttribute("href")
        productName = productElements(i).innerText
        
        ' Linki ve ismi kontrol et
        If Len(productURL) > 0 And Len(productName) > 0 Then
            Cells(i + 2, 1).Value = productName
            Cells(i + 2, 2).Value = "https://www.pttavm.com" & productURL
        End If
        On Error GoTo 0
    Next i
    
    MsgBox "Ürün linkleri çekildi. Toplam " & (i - 1) & " ürün bulundu.", vbInformation
End Sub
