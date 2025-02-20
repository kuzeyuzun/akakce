Attribute VB_Name = "Module30"
Sub CreateLinks()
    Dim lastRow As Long
    Dim i As Long
    Dim productID As String
    Dim baseURL As String
    
    ' Temel URL
    baseURL = "https://www.pttavm.com/a-p-"
    
    ' D s�tunundaki son dolu sat�r� bul
    lastRow = Cells(Rows.Count, "E").End(xlUp).row
    
    ' D s�tunundaki her dolu h�cre i�in d�ng�
    For i = 1 To lastRow
        productID = Cells(i, "E").Value
        If productID <> "" Then
            ' Linki olu�tur ve A s�tununa yaz
            Cells(i, "A").Value = baseURL & productID
        End If
    Next i
    
    Call PTT_North
    
End Sub

