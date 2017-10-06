'***********************************************************************
'* OmaMittari platform demo - OmaMittariDemoCode.xlsm
'* Copyright (c) 2017, Jatiko Oy, email: info@jatiko.fi   Date:6.10.2017
'* MIT License
'* OmaMittari Terms of usage, additional restrictions may be in each module
'***********************************************************************

'*** Find out the location of the Temp directory
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'*** Change these three credential constants to "" in the source code after running SaveCredentials() with real values on your computer
Private Const UserName As String = "testikuluttaja" '*** Demo UserName, replace this with your UserName
Private Const APItoken As String = "xyz123" '*** DemoAPItoken, replace this with your APItoken
Private Const Subscription_Key As String = "YourSubscriptionKeyFromDeveloperPortal" '*** Replace this with your Subscription Key from Developer Portal
Private Const MY_CRYPTO_KEY As String = "dkfoR5lsk#tpT" '*** Used only if saving encrypted credentials - not implemented

Private sUserName As String
Private sAPItoken As String
Private sSubscription_Key As String

'*** RECOMMENDATION: Encrypt credentials and save them to Windows registry or memory stick away from source code,
'*** then change credentials to "" in this source code. Running this code saves credentials to registry.
Sub SaveCredentials()
EncryptionCSPConnect
SaveSetting "OmaMittari", "VBAExcelDemo", "UserName", EncryptData(sUserName, MY_CRYPTO_KEY)
SaveSetting "OmaMittari", "VBAExcelDemo", "APItoken", EncryptData(sAPItoken, MY_CRYPTO_KEY)
SaveSetting "OmaMittari", "VBAExcelDemo", "Subscription_Key", EncryptData(sSubscription_Key, MY_CRYPTO_KEY)
EncryptionCSPDisconnect
End Sub

'*** RECOMMENDATION: Decrypt credentials from Windows registry or memory stick after clearing them from source code,
'*** Running this code gets credentials from registry.
Sub GetCredentials()
EncryptionCSPConnect
sUserName = DecryptData(GetSetting("OmaMittari", "VBAExcelDemo", "UserName", EncryptData(sUserName, MY_CRYPTO_KEY)), MY_CRYPTO_KEY)
sAPItoken = DecryptData(GetSetting("OmaMittari", "VBAExcelDemo", "APItoken", EncryptData(sAPItoken, MY_CRYPTO_KEY)), MY_CRYPTO_KEY)
sSubscription_Key = DecryptData(GetSetting("OmaMittari", "VBAExcelDemo", "Subscription_Key", EncryptData(sSubscription_Key, MY_CRYPTO_KEY)), MY_CRYPTO_KEY)
EncryptionCSPDisconnect

'*** In case Windows registry is empty and contants are not
If Len(sUserName) = 0 Then sUserName = UserName
If Len(sAPItoken) = 0 Then sAPItoken = APItoken
If Len(sSubscription_Key) = 0 Then sSubscription_Key = Subscription_Key
End Sub


'*** Get data from REST service
Function JsonDataFetch(URI_Start As String, URI_End As String) As String
Dim p As Object
Dim HttpReq As Object
Dim timestr As String, UserName As String, APItoken As String, Subscription_Key As String

SaveCredentials '*** After running this code at least once you can comment this row and clear credentials from the source code
GetCredentials
Set HttpReq = CreateObject("MSXML2.XMLHTTP")
HttpReq.Open "GET", URI_Start & "/" & URI_End
HttpReq.setRequestHeader "Ocp-Apim-Subscription-Key", sSubscription_Key

timestr = timeNow() 'Pick timestring for header
HttpReq.setRequestHeader "Authorization", authStr(sAPItoken, "/" & URI_End, sUserName, timestr)

Call HttpReq.send
Do While HttpReq.readyState <> 4
    DoEvents
Loop
JsonDataFetch = HttpReq.responseText
Set p = Nothing
Set HttpReq = Nothing
End Function

'*** Time to string in right format
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