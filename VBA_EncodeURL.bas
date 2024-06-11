Attribute VB_Name = "VBA_EncodeURL"
Function EncodeURL(ByVal str As String) As String
    Dim i As Integer
    Dim c As String
    Dim encodedStr As String
    Dim byteArray() As Byte
    Dim stream As Object
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Mode = 3 ' Read/Write
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText str
    stream.Position = 0
    stream.Type = 1 ' Binary
    stream.Position = 3 ' Skip BOM (EF BB BF)
    byteArray = stream.Read()
    stream.Close
    Set stream = Nothing
    
    encodedStr = ""
    
    For i = 0 To UBound(byteArray)
        c = Chr(byteArray(i))
        If byteArray(i) >= 48 And byteArray(i) <= 57 Or _
           byteArray(i) >= 65 And byteArray(i) <= 90 Or _
           byteArray(i) >= 97 And byteArray(i) <= 122 Or _
           byteArray(i) = 45 Or byteArray(i) = 95 Or _
           byteArray(i) = 46 Or byteArray(i) = 126 Then
            encodedStr = encodedStr & c
        Else
            encodedStr = encodedStr & "%" & Right("0" & Hex(byteArray(i)), 2)
        End If
    Next
    
    EncodeURL = encodedStr
End Function

Sub cs()
Dim testStr
testStr = "Hello World! 这是一个测试。"

MsgBox EncodeURL(testStr)
End Sub
