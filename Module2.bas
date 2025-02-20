Attribute VB_Name = "Module2"
Sub ScrapePTTAVMProductLinks_Test()
    Dim xhr As Object
    Dim htmlDoc As Object
    Dim productElements As Object
    Dim i As Long
    Dim productURL As String
    Dim productName As String
    
    ' URL'yi tan�mla
    Dim url As String
    url = "https://www.pttavm.com/kampanyalar/ayin-en-cok-satan-tvleri/ses-sistemleri"
    
    ' Internet Explorer i�in HTML belge nesnesini olu�tur
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    xhr.Open "GET", url, False
    xhr.send
    
    Set htmlDoc = CreateObject("HTMLFile")
    htmlDoc.body.innerHTML = xhr.responseText
    
    ' Farkl� HTML etiketlerini ve s�n�flar� kontrol et
    Set productElements = htmlDoc.getElementsByClassName("product-card")
    
    If productElements.Length = 0 Then
        Set productElements = htmlDoc.getElementsByTagName("a")
    End If
    
    ' Aktif �al��ma sayfas�n� se�
    Worksheets("�r�nler").Activate
    
    ' Ba�l�k sat�r�
    Cells(1, 1).Value = "�r�n Ad�"
    Cells(1, 2).Value = "�r�n Linki"
    
    ' �r�n linklerini Excel'e aktar
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
    
    MsgBox "�r�n linkleri �ekildi. Toplam " & (i - 1) & " �r�n bulundu.", vbInformation
End Sub
