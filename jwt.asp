<!--#include file="utils.asp"-->
<%
' Accepts an ASP dictionary of key/value pairs and a secret and
' returns a signed JSON Web Token
Function JWTEncode(dPayload, sSecret)
  Dim sPayload, sHeader, sBase64Payload, sBase64Header
  Dim sSignature, sToken

  If Typename(dPayload) = "Dictionary" Then
    sPayload = DictionaryToJSONString(dPayload)
  Else
    sPayload = dPayload
  End If

  sHeader  = JWTHeaderDictionary()

  sBase64Payload = SafeBase64Encode(sPayload)
  sBase64Header  = SafeBase64Encode(sHeader)

  sPayload       = sBase64Header & "." & sBase64Payload
  sSignature     = SHA256SignAndEncode(sPayload, sSecret)
  sToken         = sPayload & "." & sSignature

  JWTEncode = sToken
End Function

' SHA256 HMAC
Function SHA256SignAndEncode(sIn, sKey)
  Dim sSignature

  'Open WSC object to access the encryption function
  Dim sha256
  Set sha256 = GetObject("script:"&Server.MapPath("/jwt/external/sha256.wsc"))

  'SHA256 sign data
  sSignature = sha256.b64_hmac_sha256(sKey, sIn)
  sSignature = Base64ToSafeBase64(sSignature)

  SHA256SignAndEncode = sSignature
End Function

' Returns a static JWT header dictionary
Function JWTHeaderDictionary()
  Dim dOut
  Set dOut = Server.CreateObject("Scripting.Dictionary")
  dOut.Add "typ", "JWT"
  dOut.Add "alg", "HS256"

  JWTHeaderDictionary = DictionaryToJSONString(dOut)
End Function

' Returns decoded payload (not verify)
Function JWTDecode(token)
    Dim tokenSplited, sPayload
    tokenSplited = Split(token, ".")
    If UBound(tokenSplited) <> 2 Then
        JWTDecode = "Invalid token"
    Else
        sPayload = tokenSplited(1)
        sPayload = SafeBase64ToBase64(sPayload)
        JWTDecode = Base64Decode(sPayload)
    End If
End Function

' Updated JWTVerify function
Function JWTVerify(token, sKey)
    Dim parts, header, payload, signature, signedData, computedSignature
    parts = Split(token, ".")
    
    If UBound(parts) <> 2 Then
        JWTVerify = False
        Exit Function
    End If
    
    header = parts(0)
    payload = parts(1)
    signature = SafeBase64ToBase64(parts(2))
    
    signedData = header & "." & payload
    computedSignature = SHA256SignAndEncode(signedData, sKey)
    computedSignature = SafeBase64ToBase64(computedSignature)
    
    JWTVerify = (signature = computedSignature)
End Function

' Helper function to convert Safe Base64 to regular Base64
Function SafeBase64ToBase64(input)
    SafeBase64ToBase64 = Replace(Replace(input, "-", "+"), "_", "/")
    Dim padding
    padding = Len(input) Mod 4
    If padding > 0 Then
        SafeBase64ToBase64 = SafeBase64ToBase64 & String(4 - padding, "=")
    End If
End Function
%>
