# VBA Encoding , Encryption dan Hashing
kumpulan script VBA untuk kebutuhan Encoding , Encryption dan Hashing

### 1. Endcoding
#### Base64
```VB
Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As Object
  Dim objNode As Object

  Set objXML = CreateObject("MSXML2.DOMDocument")
  
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = Replace(objNode.text, vbLf, "")

  Set objNode = Nothing
  Set objXML = Nothing
End Function
```

### 2. Encryption
#### Caesar-Chipher
```VB
Public Function CaesarCipher(ByVal TextToEncrypt As String, ByVal CaesarShift As Long) As String

    Dim OutputText As String
    TextToEncrypt = UCase(TextToEncrypt)

    If CaesarShift > 26 Then
        CaesarShift = CaesarShift Mod 26
    End If

    If CaesarShift = 0 Then
        OutputText = TextToEncrypt
    ElseIf CaesarShift > 0 Then
        OutputText = ShiftRight(TextToEncrypt, CaesarShift)
    Else
        CaesarShift = Abs(CaesarShift)
        OutputText = ShiftLeft(TextToEncrypt, CaesarShift)
    End If

    CaesarCipher = OutputText
End Function

Private Function ShiftLeft(ByVal ShiftString As String, ByVal ShiftQuantity As Long) As String

    Dim TextLength As Long
    TextLength = Len(ShiftString)

    Dim CipherText As String
    Dim CharacterCode As Long
    Dim AsciiIndex As Long
    Dim AsciiIdentifier() As Long
    ReDim AsciiIdentifier(1 To TextLength)

    For AsciiIndex = 1 To TextLength
        CharacterCode = Asc(Mid(ShiftString, AsciiIndex, 1))
        If CharacterCode = 32 Then GoTo Spaces
        If CharacterCode - ShiftQuantity < 65 Then
            CharacterCode = CharacterCode + 26 - ShiftQuantity
        Else: CharacterCode = CharacterCode - ShiftQuantity
        End If
Spaces:
        AsciiIdentifier(AsciiIndex) = CharacterCode
    Next

        For AsciiIndex = 1 To TextLength
            CipherText = CipherText & Chr(AsciiIdentifier(AsciiIndex))
        Next
    ShiftLeft = CipherText
End Function

Private Function ShiftRight(ByVal ShiftString As String, ByVal ShiftQuantity As Long) As String

    Dim TextLength As Long
    TextLength = Len(ShiftString)

    Dim CipherText As String
    Dim CharacterCode As Long
    Dim AsciiIndex As Long
    Dim AsciiIdentifier() As Long
    ReDim AsciiIdentifier(1 To TextLength)

    For AsciiIndex = 1 To TextLength
        CharacterCode = Asc(Mid(ShiftString, AsciiIndex, 1))
        If CharacterCode + ShiftQuantity > 90 Then
            CharacterCode = CharacterCode - 26 + ShiftQuantity
        ElseIf CharacterCode = 32 Then GoTo Spaces
        Else:  CharacterCode = CharacterCode + ShiftQuantity
        End If
Spaces:
        AsciiIdentifier(AsciiIndex) = CharacterCode
    Next

        For AsciiIndex = 1 To TextLength
            CipherText = CipherText & Chr(AsciiIdentifier(AsciiIndex))
        Next
    ShiftRight = CipherText
End Function
```

### 3. Hashing
#### MD5
```VB
Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
        
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    MD5 = ConvToHexString(bytes)
        
    Set oT = Nothing
    Set oMD5 = Nothing
End Function
```
#### SHA-1
```VB
Public Function SHA1(sIn As String, Optional bB64 As Boolean = 0) As String
    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
            
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA1.ComputeHash_2((TextToHash))
        
    SHA1 = ConvToHexString(bytes)
            
    Set oT = Nothing
    Set oSHA1 = Nothing
    
End Function
```
#### SHA-256
```VB
Public Function SHA256(sIn As String, Optional bB64 As Boolean = 0) As String
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    
    SHA256 = ConvToHexString(bytes)
    
    Set oT = Nothing
    Set oSHA256 = Nothing
End Function
```
#### SHA-384
```vb
Public Function SHA384(sIn As String, Optional bB64 As Boolean = 0) As String
    Dim oT As Object, oSHA384 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA384.ComputeHash_2((TextToHash))
    
    SHA384 = ConvToHexString(bytes)

    Set oT = Nothing
    Set oSHA384 = Nothing
    
End Function
```
#### SHA-512
```VB
Public Function SHA512(sIn As String, Optional bB64 As Boolean = 0) As String
    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA512.ComputeHash_2((TextToHash))
    
    SHA512 = ConvToHexString(bytes)
    
    Set oT = Nothing
    Set oSHA512 = Nothing
    
End Function
```
