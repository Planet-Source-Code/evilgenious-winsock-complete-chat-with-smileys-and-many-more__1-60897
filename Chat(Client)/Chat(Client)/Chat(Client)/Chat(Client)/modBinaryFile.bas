Attribute VB_Name = "modBinaryFile"
Sub SaveBinaryArray(ByVal Filename As String, WriteData() As Byte)

    Dim t As Integer
    t = FreeFile
    Open Filename For Binary Access Write As #t
        
            Put #t, , WriteData()
        
    Close #t
    
End Sub

Function ReadBinaryArray(ByVal Source As String)

    Dim bytBuf() As Byte
    Dim intN As Long
    
    Dim t As Integer
    t = FreeFile
    
    Open Source For Binary Access Read As #t
    
    Dim n As Long
    
    ReDim bytBuf(1 To LOF(t)) As Byte
    Get #t, , bytBuf()
    
    ReadBinaryArray = bytBuf()
    
    Close #t
    
End Function

Public Function StripPath(t As String) As String

  Dim X As Integer
  Dim ct As Integer

    StripPath = t
    X = InStr(t, "\")
    Do While X
        ct = X
        X = InStr(ct + 1, t, "\")
    Loop
    If ct > 0 Then StripPath = Mid$(t, ct + 1)

End Function
