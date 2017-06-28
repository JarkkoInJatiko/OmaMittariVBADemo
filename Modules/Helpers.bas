Attribute VB_Name = "Helpers"
'***********************************************************************
'* OmaMittari palvelualustan demo - OmaMittariDemoCode.xlsm
'* Copyright (c) 2017, Jatiko Oy, email: info@jatiko.fi              Pvm:21.6.2017
'* T‰t‰ l‰hdekoodia saa k‰ytt‰‰ ja levitt‰‰ vapaasti, kunhan noudattaa
'* OmaMittari palvelun ehtoja ja modulikohtaisia rajoituksia
'***********************************************************************

'*** Find out the location of the Temp directory
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'*** Change these three credential constants to "" in the source code after running SaveCredentials() with real values on your computer
Private Const sUserName As String = "zzzz" '*** Replace this with your UserName
Private Const sAPItoken As String = "zzzz" '*** Replace this with your APItoken
Private Const sSubscription_Key As String = "zzzz" '*** Replace this with your Subscription_Key
Private Const MY_CRYPTO_KEY As String = "tdfuhwfnasdfisdflkaa"

'*** RECOMMENDATION: Encrypt credentials and save them to Windows registry, then change credentials to "" in this source code. Running this code saves credentials to registry.
Sub SaveCredentials()
EncryptionCSPConnect
SaveSetting "OmaMittari", "VBAExcelDemo", "UserName", EncryptData(sUserName, MY_CRYPTO_KEY)
SaveSetting "OmaMittari", "VBAExcelDemo", "APItoken", EncryptData(sAPItoken, MY_CRYPTO_KEY)
SaveSetting "OmaMittari", "VBAExcelDemo", "Subscription_Key", EncryptData(sSubscription_Key, MY_CRYPTO_KEY)
EncryptionCSPDisconnect
End Sub

'*** Get data from REST service
Function JsonDataFetch(URI_Start As String, URI_End As String) As String
Dim p As Object
Dim HttpReq As Object
Dim timestr As String, UserName As String, APItoken As String, Subscription_Key As String

EncryptionCSPConnect
UserName = DecryptData(GetSetting("OmaMittari", "VBAExcelDemo", "UserName", EncryptData(sUserName, MY_CRYPTO_KEY)), MY_CRYPTO_KEY)
APItoken = DecryptData(GetSetting("OmaMittari", "VBAExcelDemo", "APItoken", EncryptData(sAPItoken, MY_CRYPTO_KEY)), MY_CRYPTO_KEY)
Subscription_Key = DecryptData(GetSetting("OmaMittari", "VBAExcelDemo", "Subscription_Key", EncryptData(sSubscription_Key, MY_CRYPTO_KEY)), MY_CRYPTO_KEY)
EncryptionCSPDisconnect

Set HttpReq = CreateObject("MSXML2.XMLHTTP")
HttpReq.Open "GET", URI_Start & "/" & URI_End
HttpReq.setRequestHeader "Ocp-Apim-Subscription-Key", Subscription_Key

timestr = timeNow() 'Pick timestring for header
HttpReq.setRequestHeader "Authorization", authStr(APItoken, "/" & URI_End, UserName, timestr)

Call HttpReq.send
Do While HttpReq.readyState <> 4
    DoEvents
Loop
JsonDataFetch = HttpReq.responseText
Set p = Nothing
Set HttpReq = Nothing
End Function

'*** Kellonaika merkkijonoksi oikeassa formaatissa
Function timeNow() As String
    timeNow = Format(Now(), "yyyyMMddhhmmss")
End Function

Function authStr(APItoken As String, requestUrl As String, UserName As String, timeField) As String
Dim myPhrase As String
Dim SHA256Obj As New SHA256

myPhrase = APItoken & "|" & requestUrl & "|" & timeField
authStr = UserName & "|" & SHA256Obj.SHA256(myPhrase) & "|" & timeField
End Function

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False, Optional UTF8Encode As Boolean = True) As String
Dim StringValCopy As String: StringValCopy = IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsPlus Then Space = "+" Else Space = "%20"

  For i = 1 To StringLen
    Char = Mid$(StringValCopy, i, 1)
    CharCode = Asc(Char)
    Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        Result(i) = Char
      Case 32
        Result(i) = Space
      Case 0 To 15
        Result(i) = "%0" & Hex(CharCode)
      Case Else
        Result(i) = "%" & Hex(CharCode)
    End Select
  Next i
  URLEncode = Join(Result, "")
End If
End Function

Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
Dim lLength As Long
If UTF16 <> "" Then
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    sBuffer = Space$(lLength)
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, lLength - 1)
Else
    UTF16To8 = ""
End If
End Function


