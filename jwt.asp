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
  Set sha256 = GetObject("script:"&Server.MapPath("./external/sha256.wsc"))

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

' Returns if token is valid
Function JWTVerify(token, sKey)
    Dim jsonPayload, reEncodingToken, tokenPayload
    tokenPayload = JWTDecode(token)
    reEncodingToken = JWTEncode(tokenPayload, sKey)

    JWTVerify = (token = reEncodingToken)
End Function
%>
