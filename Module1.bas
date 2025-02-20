Attribute VB_Name = "Module1"
Sub Temizle()
Attribute Temizle.VB_ProcData.VB_Invoke_Func = " \n14"

' Verileri Temizleme

    Sheets("PttAVM").Select
    Range("A2:D999").Select
    Selection.ClearContents
    Range("G6").Select
    
    Sheets("ID").Select
    Range("A2:A999").Select
    Selection.ClearContents
    Range("A2").Select
    
    
End Sub
