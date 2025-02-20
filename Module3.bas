Attribute VB_Name = "Module3"
Sub ScrapePTTAVMProductIDs()
    Dim xhr As Object
    Dim responseText As String
    
    ' URL'yi tan�mla
    Dim url As String
    url = "https://www.pttavm.com/kampanyalar/ayin-en-cok-satan-tvleri/tv-box---uydu-alici"
    
    ' WinHttpRequest kullan
    Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Detayl� HTTP ayarlar�
    xhr.Open "GET", url, False
    xhr.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    xhr.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    xhr.setRequestHeader "Accept-Language", "tr-TR,tr;q=0.8"
    
    On Error Resume Next
    xhr.send
    
    If xhr.Status <> 200 Then
        MsgBox "Sayfa y�klenemedi. Hata: " & xhr.Status, vbCritical
        Exit Sub
    End If
    
    responseText = xhr.responseText
    
    ' Geli�mi� RegEx deseni
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Global = True
    regex.Pattern = "(?<=p-)(\d+)(?=\?position=)"
    
    Dim matches As Object
    Set matches = regex.Execute(responseText)
    
    ' Excel'e aktarma
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("�r�nler")
    
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "�r�n ID"
    
    Dim i As Long, uniqueIDs As Object
    Set uniqueIDs = CreateObject("Scripting.Dictionary")
    
    For i = 0 To matches.Count - 1
        If Not uniqueIDs.Exists(matches(i).Value) Then
            uniqueIDs.Add matches(i).Value, True
            ws.Cells(uniqueIDs.Count + 1, 1).Value = matches(i).Value
        End If
    Next i
    
    MsgBox "Toplam " & uniqueIDs.Count & " benzersiz �r�n ID'si bulundu.", vbInformation
End Sub
