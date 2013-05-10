<!--#include file="utils.asp"-->
<%
' Accepts an ASP dictionary of key/value pairs and a secret and
' returns a signed JSON Web Token
Function JWTEncode(dPayload, sSecret)
  Dim sPayload, sHeader, sBase64Payload, sBase64Header
  Dim sSignature, sToken

  sPayload = DictionaryToJSONString(dPayload)
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
%>
