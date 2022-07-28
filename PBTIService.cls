VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTIService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Popbill API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dllm. most of windows system has it
' on system32 by default.
'
' http://www.popbill.com
' Author : Kim Seongjun
' Written : 2014-04-22
' Contributor : Jeong Yohan (code@linkhubcorp.com)
' Updated : 2022-07-26
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID_REAL = "POPBILL"
Private Const ServiceID_TEST = "POPBILL_TEST"
Private Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Private Const ServiceURL_TEST = "https://popbill-test.linkhub.co.kr"
Private Const ServiceURL_Static_REAL = "https://static-popbill.linkhub.co.kr"
Private Const ServiceURL_Static_TEST = "https://static-popbill-test.linkhub.co.kr"
Private Const ServiceURL_GA_REAL = "https://ga-popbill.linkhub.co.kr"
Private Const ServiceURL_GA_TEST = "https://ga-popbill-test.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private Const adTypeBinary As Long = 1
Private Const adTypeText As Long = 2
Private Const adModeReadWrite As Long = 3

Private m_IsTest As Boolean
Private m_IPRestrictOnOff As Boolean
Private m_UseStaticIP As Boolean
Private m_UseGAIP As Boolean
Private m_UseLocalTimeYN As Boolean
Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String
Private m_token_Dic

'문서번호 형태
Public Enum MgtKeyType
    SELL '매출
    BUY '매입
    TRUSTEE '수탁
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Class_Initialize()
    Set m_token_Dic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    m_token_Dic.RemoveAll
    Set m_token_Dic = Nothing
End Sub

Private Function URLEncode(ByVal StringToEncode As String) As String
   Dim i                As Integer
   Dim iAsc             As Long
   Dim sTemp            As String
   
   Dim ByteArrayToEncode() As Byte
 
   ByteArrayToEncode = ADO_EncodeUTF8(StringToEncode)
   
   For i = 0 To UBound(ByteArrayToEncode)
      iAsc = ByteArrayToEncode(i)
      Select Case iAsc
         Case 32 'space
            sTemp = "+"
         Case 48 To 57, 65 To 90, 97 To 122
            sTemp = Chr(ByteArrayToEncode(i))
         Case Else
            Debug.Print iAsc
            sTemp = "%" & Hex(iAsc)
      End Select
      URLEncode = URLEncode & sTemp
   Next
 
End Function
  
'Purpose: UTF16 to UTF8 using ADO
Private Function ADO_EncodeUTF8(ByVal strUTF16 As String) As Byte()
 
   Dim objStream        As Object
   Dim Data()           As Byte
 
   Set objStream = CreateObject("ADODB.Stream")
   objStream.Charset = "utf-8"
   objStream.Mode = adModeReadWrite
   objStream.Type = adTypeText
   objStream.Open
   objStream.WriteText strUTF16
   objStream.Flush
   objStream.Position = 0
   objStream.Type = adTypeBinary
   objStream.Read 3 ' skip BOM
   Data = objStream.Read()
   objStream.Close
   ADO_EncodeUTF8 = Data
 
End Function

Private Function UniStrToUTF8(UniString) As Byte()
   ' Convert a Unicode string to a byte stream of UTF-8
   Dim BArray() As Byte
   Dim TempB() As Byte
   Dim i As Long
   Dim k As Long
   Dim TLen As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Dim UTF16 As Long
   Dim j
   TLen = Len(UniString)  ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function  ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2  ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
         ReDim Preserve BArray(k)
         BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray  ' Return the resultant UTF-8 byte array
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As Byte()
   ' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
   Dim BArray() As Byte
   If UTF16 < &H80 Then
      ReDim BArray(0)  ' one byte UTF-8
      BArray(0) = UTF16  ' Use number as is
   ElseIf UTF16 < &H800 Then
      ReDim BArray(1)  ' two byte UTF-8
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(0) = &HC0 + (UTF16 And &H1F)  ' Use 5 remaining bits
   Else
      ReDim BArray(2)  ' three byte UTF-8
      BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
   End If
   ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
'최근 오류코드 확인
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'최근오류메시지 확인
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'테스트 플래그
Public Property Let IsTest(ByVal value As Boolean)
    m_IsTest = value
End Property
Public Property Let IPRestrictOnOff(ByVal value As Boolean)
    m_IPRestrictOnOff = value
End Property
Public Property Let UseStaticIP(ByVal value As Boolean)
    m_UseStaticIP = value
End Property
Public Property Let UseGAIP(ByVal value As Boolean)
    m_UseGAIP = value
End Property
Public Property Let UseLocalTimeYN(ByVal value As Boolean)
    m_UseLocalTimeYN = value
End Property
'초기화
Public Sub Initialize(LinkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.LinkID = LinkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "member"
    ''세금계산서
    m_scope.Add "110"
    
    m_IsTest = False
    m_IPRestrictOnOff = True
    m_UseStaticIP = False
    m_UseGAIP = False
    m_UseLocalTimeYN = False
End Sub

Private Function getTargetURL() As String
    If m_UseGAIP Then
        getTargetURL = IIf(m_IsTest, ServiceURL_GA_TEST, ServiceURL_GA_REAL)
    ElseIf m_UseStaticIP Then
        getTargetURL = IIf(m_IsTest, ServiceURL_Static_TEST, ServiceURL_Static_REAL)
    Else
        getTargetURL = IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL)
    End If
End Function
Private Function getSession_token(CorpNum As String) As String
    Dim refresh As Boolean
    
    refresh = False
    Dim m_Token As LinkhubToken
    
    Set m_Token = Nothing
    
    If m_token_Dic.Exists(CorpNum) Then
        Set m_Token = m_token_Dic.Item(CorpNum)
    End If
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = m_Linkhub.GetTime(m_UseStaticIP, m_UseGAIP, m_UseLocalTimeYN)
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope, IIf(m_IPRestrictOnOff, "", "*"), m_UseStaticIP, m_UseGAIP, m_UseLocalTimeYN)
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "POPBILL", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
        If m_token_Dic.Exists(CorpNum) Then
            m_token_Dic.Remove (CorpNum)
        End If
        
        m_token_Dic.Add CorpNum, m_Token
        
    End If
    
    getSession_token = m_Token.session_token

End Function

Private Function httpGET(URL As String, BearerToken As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("GET", getTargetURL() + URL, False)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    Call winhttp1.setRequestHeader("User-Agent", "VB6 POPBiLL SDK")
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    winhttp1.Send
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function

Private Function httpPOST(URL As String, BearerToken As String, postData As String, Optional UserID As String) As Variant
    Set httpPOST = httpPOST_override(URL, BearerToken, "", postData, UserID)
    
End Function


Private Function httpPOST_override(URL As String, BearerToken As String, override As String, postData As String, Optional UserID As String, Optional contentstype As String = "") As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", getTargetURL() + URL, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    Call winhttp1.setRequestHeader("User-Agent", "VB6 POPBiLL SDK")
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    If contentstype <> "" Then
        Call winhttp1.setRequestHeader("Content-Type", contentstype)
    Else
        Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    End If
    
 
    winhttp1.Send (postData)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_override = m_Linkhub.parse(Response)

End Function

Private Function httpPOST_Bulk(URL As String, BearerToken As String, override As String, submitID As String, postData As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", getTargetURL() + URL, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    Call winhttp1.setRequestHeader("User-Agent", "VB6 POPBiLL SDK")
    Call winhttp1.setRequestHeader("x-pb-message-digest", m_Linkhub.Encode(m_Linkhub.SHA1(m_Linkhub.ToUTF8(postData)), edfBase64, efNoFolding) + Chr(10))
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If submitID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-submit-id", submitID)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postData)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_Bulk = m_Linkhub.parse(Response)

End Function

Private Function GetOnlyFileName(ByVal FilePath As String) As String

     Dim Temp() As String
     Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))

End Function

Private Function httpPOST_File(URL As String, BearerToken As String, postData As String, FilePath As String, Optional UserID As String) As Variant
     
  Dim winhttp1
    Dim boundary As String
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", getTargetURL() + URL, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("User-Agent", "VB6 POPBiLL SDK")
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    
    If postData <> "" Then
        'TODO 첨부전문 구성.
    
    End If
    
    Dim head, tail As String
    
    head = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""Filedata""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
           "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
           
    tail = vbCrLf & "--" & boundary & "--" & vbCrLf
    
    '### 이미지 바이너리를 얻는다.
    Dim binary() As Byte
    Open FilePath For Binary Access Read As #1
         ReDim binary(LOF(1))
         Get #1, , binary
    Close #1
    
    '### 전송할 데이터를 만든다.
    Dim BHeader1() As Byte, BHeader2() As Byte, Buffer() As Byte
    Dim Size1 As Long, Size2 As Long, Size3 As Long
    BHeader1 = UniStrToUTF8(head)
    BHeader2 = StrConv(tail, vbFromUnicode)
    
    Size1 = UBound(BHeader1)
    Size2 = UBound(binary)
    Size3 = UBound(BHeader2)
    
    ReDim Buffer(Size1 + Size2 + Size3 + 2)
    CopyMemory Buffer(0), BHeader1(0), Size1 + 1
    CopyMemory Buffer(Size1 + 1), binary(0), Size2 + 1
    CopyMemory Buffer(Size1 + Size2 + 2), BHeader2(0), Size3 + 1
    
    winhttp1.Send (Buffer)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_File = m_Linkhub.parse(Response)

End Function
'연동회원 아이디 중복 확인
Public Function CheckID(IDString As String) As PBResponse
On Error GoTo ErrHandler
    If IDString = "" Then
        Err.Raise -99999999, "POPBILL", "중복여부를 확인할 아이디가 입력되지 않았습니다."
    End If

    Dim result As Variant
    
    Set result = httpGET("/IDCheck?ID=" + IDString, "")
    
    Set CheckID = New PBResponse
    CheckID.code = result.Item("code")
    CheckID.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckID = Nothing
End Function

'연동회원 담당자 정보확인
Public Function GetContactInfo(CorpNum As String, ContactID As String, Optional UserID As String = "") As PBContactInfo
On Error GoTo ErrHandler
    Dim result As Variant
    Dim postData As String
    
    Dim tmp
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactID
            
    postData = m_Linkhub.toString(tmp)
    
    
    Set result = httpPOST("/Contact", getSession_token(CorpNum), postData, UserID)
    
    Set GetContactInfo = New PBContactInfo
    GetContactInfo.fromDictionary result
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetContactInfo = Nothing
End Function

'연동회원 담당자 목록조회
Public Function ListContact(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
        
    Dim result As Variant
       
    Set result = httpGET("/IDs", getSession_token(CorpNum), UserID)
    
    Set ListContact = New Collection
    
    Dim T As Variant
    Dim info As PBContactInfo
    
    For Each T In result
        Set info = New PBContactInfo
        
        info.fromDictionary T
        
        ListContact.Add info
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set ListContact = Nothing
End Function
'담당자 추가
Public Function RegistContact(CorpNum As String, ContactInfo As PBContactInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set RegistContact = Nothing
    
    Dim postData As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "pwd", ContactInfo.pwd
    tmp.Add "Password", ContactInfo.Password
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
    tmp.Add "searchRole", ContactInfo.searchRole
            
    postData = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs/New", getSession_token(CorpNum), postData, UserID)
    
    Set RegistContact = New PBResponse
    
    RegistContact.code = result.Item("code")
    RegistContact.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'담당자 정보 수정
Public Function UpdateContact(CorpNum As String, ContactInfo As PBContactInfo, UserID As String) As PBResponse
On Error GoTo ErrHandler
    Set UpdateContact = Nothing
    
    Dim postData As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
        
    tmp.Add "id", ContactInfo.id
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
    tmp.Add "searchRole", ContactInfo.searchRole
            
    postData = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs", getSession_token(CorpNum), postData, UserID)
    
    Set UpdateContact = New PBResponse
    
    UpdateContact.code = result.Item("code")
    UpdateContact.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'회사정보 수정
Public Function UpdateCorpInfo(CorpNum As String, CorpInfo As PBCorpInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set UpdateCorpInfo = Nothing
    
    Dim postData As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "ceoname", CorpInfo.ceoname
    tmp.Add "corpName", CorpInfo.corpName
    tmp.Add "addr", CorpInfo.addr
    tmp.Add "bizType", CorpInfo.bizType
    tmp.Add "bizClass", CorpInfo.bizClass
            
    postData = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/CorpInfo", getSession_token(CorpNum), postData, UserID)
    
    Set UpdateCorpInfo = New PBResponse
    
    UpdateCorpInfo.code = result.Item("code")
    UpdateCorpInfo.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'파트너 포인트 충전 URL
Public Function GetPartnerURL(CorpNum As String, TOGO As String) As String
    On Error GoTo ErrHandler
    
    GetPartnerURL = m_Linkhub.GetPartnerURL(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), TOGO, m_UseStaticIP, m_UseGAIP)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPartnerURL = m_LastErrCode
End Function

'회원잔액조회
Public Function GetBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), m_UseStaticIP, m_UseGAIP)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), m_UseStaticIP, m_UseGAIP)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPartnerBalance = m_LastErrCode
End Function
'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetPopbillURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopbillURL = ""
End Function


'팝빌 로그인 기본 URL
Public Function GetAccessURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=LOGIN", getSession_token(CorpNum), UserID)
    
    GetAccessURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetAccessURL = ""
End Function

'팝빌 연동회원 포인트 충전 URL
Public Function GetChargeURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=CHRG", getSession_token(CorpNum), UserID)
    
    GetChargeURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetChargeURL = ""
End Function

'팝빌 연동회원 포인트 결제내역 팝업 URL
Public Function GetPaymentURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=PAYMENT", getSession_token(CorpNum), UserID)
    
    GetPaymentURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPaymentURL = ""
End Function

'팝빌 연동회원 포인트 사용내역 팝업 URL
Public Function GetUseHistoryURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=USEHISTORY", getSession_token(CorpNum), UserID)
    
    GetUseHistoryURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetUseHistoryURL = ""
End Function

'팝빌 인감 및 첨부문서 등록 URL
Public Function GetSealURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
   Dim result As Variant
   
   Set result = httpGET("/?TG=SEAL", getSession_token(CorpNum), UserID)
   
   GetSealURL = result.Item("url")
   Exit Function
ErrHandler:
   m_LastErrCode = Err.Number
   m_LastErrMessage = Err.Description
   GetSealURL = ""
End Function

'공동인증서 등록 URL
Public Function GetTaxCertURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
   Dim result As Variant
   
   Set result = httpGET("/?TG=CERT", getSession_token(CorpNum), UserID)
   
   GetTaxCertURL = result.Item("url")
   Exit Function
ErrHandler:
   m_LastErrCode = Err.Number
   m_LastErrMessage = Err.Description
   GetTaxCertURL = ""
End Function


'회원가입 여부
Public Function CheckIsMember(CorpNum As String, LinkID As String) As PBResponse
On Error GoTo ErrHandler
    If CorpNum = "" Then
        Err.Raise -99999999, "POPBILL", "팝빌회원의 사업자번호가 입력되지 않았습니다."
    End If
    
    If LinkID = "" Then
        Err.Raise -99999999, "POPBILL", "링크아이디(LinkID)가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + LinkID, "")
    
    Set CheckIsMember = New PBResponse
    CheckIsMember.code = result.Item("code")
    CheckIsMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckIsMember = Nothing
End Function
'회원가입
Public Function JoinMember(JoinInfo As PBJoinForm) As PBResponse
On Error GoTo ErrHandler
    Set JoinMember = Nothing
    
    Dim postData As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "LinkID", JoinInfo.LinkID
    tmp.Add "CorpNum", JoinInfo.CorpNum
    tmp.Add "CorpName", JoinInfo.corpName
    tmp.Add "CEOName", JoinInfo.ceoname
    tmp.Add "Addr", JoinInfo.addr
    tmp.Add "ZipCode", JoinInfo.ZipCode
    tmp.Add "BizClass", JoinInfo.bizClass
    tmp.Add "BizType", JoinInfo.bizType
    tmp.Add "ContactName", JoinInfo.ContactName
    tmp.Add "ContactEmail", JoinInfo.ContactEmail
    tmp.Add "ContactFAX", JoinInfo.ContactFAX
    tmp.Add "ContactHP", JoinInfo.ContactHP
    tmp.Add "ContactTEL", JoinInfo.ContactTEL
    tmp.Add "ID", JoinInfo.id
    tmp.Add "PWD", JoinInfo.pwd
    tmp.Add "Password", JoinInfo.Password
    
    postData = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/Join", "", postData)
    
    Set JoinMember = New PBResponse
    
    JoinMember.code = result.Item("code")
    JoinMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'회사정보 조회
Public Function GetCorpInfo(CorpNum As String, Optional UserID As String = "") As PBCorpInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/CorpInfo", getSession_token(CorpNum), UserID)
    
    Set GetCorpInfo = New PBCorpInfo
        
    GetCorpInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetCorpInfo = Nothing
End Function

'과금정보 조회
Public Function GetChargeInfo(CorpNum As String, Optional UserID As String = "") As PBChargeInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/Taxinvoice/ChargeInfo", getSession_token(CorpNum), UserID)
    
    Set GetChargeInfo = New PBChargeInfo
        
    GetChargeInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetChargeInfo = Nothing
End Function


'''''''''''''''''''''''''' End Of PopbillBase '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'세금계산서 URL확인
Public Function GetURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Taxinvoice?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetURL = ""
End Function
'발행단가 확인
Public Function GetUnitCost(CorpNum As String) As Double
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Taxinvoice?cfg=UNITCOST", getSession_token(CorpNum))
    
    GetUnitCost = result.Item("unitCost")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetUnitCost = m_LastErrCode
End Function
'인증서 만료일자 확인
Public Function GetCertificateExpireDate(CorpNum As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Taxinvoice?cfg=CERT", getSession_token(CorpNum))
    
    GetCertificateExpireDate = result.Item("certificateExpiration")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetCertificateExpireDate = ""
End Function


'관려번호 사용여부 확인
Public Function checkMgtKeyInUse(CorpNum As String, KeyType As MgtKeyType, mgtKey As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey, getSession_token(CorpNum))
    
    Set checkMgtKeyInUse = New PBResponse
    
    checkMgtKeyInUse.code = IIf(IsNull(result.Item("itemKey")), 0, 1)
    checkMgtKeyInUse.message = IIf(IsNull(result.Item("itemKey")), "미사용중", "사용중")
    
    Exit Function
ErrHandler:
    If Err.Number = -11000005 Then
        Set checkMgtKeyInUse = New PBResponse
        checkMgtKeyInUse.code = 0
        checkMgtKeyInUse.message = "미사용중"
        Exit Function
    End If
    
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set checkMgtKeyInUse = Nothing
End Function
'즉시발행
Public Function RegistIssue(CorpNum As String, Taxinvoice As PBTaxinvoice, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler

    Dim result As Variant
    Dim tmpDic As Variant
    
    Set tmpDic = Taxinvoice.toDictionary
        
    Dim postData As String
    
    postData = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST_override("/Taxinvoice", getSession_token(CorpNum), "ISSUE", postData, UserID)
    
    Set RegistIssue = New PBResponse
    
    RegistIssue.code = result.Item("code")
    RegistIssue.message = result.Item("message")
    RegistIssue.ntsConfirmNum = result.Item("ntsConfirmNum")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set RegistIssue = Nothing
End Function
'초대량 발행 접수
Public Function BulkSubmit(CorpNum As String, submitID As String, taxinvoiceList As Collection, forceIssue As Boolean, Optional UserID As String = "") As PBBulkResponse
    On Error GoTo ErrHandler
    
    Dim result As Variant
    
    If submitID = "" Then
        Err.Raise -99999999, "POPBILL", "제출아이디가 입력되지 않았습니다.."
    End If
    
    If taxinvoiceList.Count = 0 Or IsEmpty(taxinvoiceList) Then
        Err.Raise -99999999, "POPBILL", "세금계산서 정보가 입력되지 않았습니다."
    End If
    
    Dim PBBulkSubmit As New PBBulkSubmit
    
    PBBulkSubmit.forceIssue = forceIssue
    
    Dim invoice As PBTaxinvoice

    For Each invoice In taxinvoiceList
    
        PBBulkSubmit.invoices.Add invoice
        
    Next
    
    Dim tmpDic As Variant
    Set tmpDic = PBBulkSubmit.toDictionary
    
    Dim postData As String
    
    postData = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST_Bulk("/Taxinvoice", getSession_token(CorpNum), "BULKISSUE", submitID, postData, UserID)
    
    Set BulkSubmit = New PBBulkResponse
    
    BulkSubmit.code = result.Item("code")
    BulkSubmit.message = result.Item("message")
    BulkSubmit.receiptID = result.Item("receiptID")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set BulkSubmit = Nothing
End Function
'초대량 접수결과 확인
Public Function GetBulkResult(CorpNum As String, submitID As String, Optional UserID As String = "") As PBBulkTaxinvoiceResult
    On Error GoTo ErrHandler
    
    Dim result As Variant
    
    If submitID = "" Then
        Err.Raise -99999999, "POPBILL", "제출아이디가 입력되지 않았습니다.."
    End If
    
    Set result = httpGET("/Taxinvoice/BULK/" + submitID + "/State", getSession_token(CorpNum), UserID)
    
    Set GetBulkResult = New PBBulkTaxinvoiceResult
    
    GetBulkResult.fromDictionary result
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetBulkResult = Nothing
End Function
'임시저장
Public Function Register(CorpNum As String, Taxinvoice As PBTaxinvoice, Optional writeSpecification As Boolean = False, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler

    Dim result As Variant
    Dim tmpDic As Variant
    
    Set tmpDic = Taxinvoice.toDictionary
    If writeSpecification Then
        tmpDic.Add "writeSpecification", True
    End If
    
    Dim postData As String
    
    postData = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST("/Taxinvoice", getSession_token(CorpNum), postData, UserID)
    
    Set Register = New PBResponse
    
    Register.code = result.Item("code")
    Register.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Register = Nothing

End Function
'수정
Public Function Update(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Taxinvoice As PBTaxinvoice, Optional writeSpecification As Boolean = False, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim tmpDic As Variant
    
    Set tmpDic = Taxinvoice.toDictionary
    If writeSpecification Then
        tmpDic.Add "writeSpecification", True
    End If
    
    Dim postData As String
    
    postData = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                    getSession_token(CorpNum), _
                                    "PATCH", _
                                    postData, _
                                    UserID)
    
    Set Update = New PBResponse
    
    Update.code = result.Item("code")
    Update.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Update = Nothing

End Function
'삭제
Public Function Delete(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, getSession_token(CorpNum), "DELETE", "", UserID)
    
    Set Delete = New PBResponse
    
    Delete.code = result.Item("code")
    Delete.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Delete = Nothing
End Function
'발행
Public Function Issue(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, memo As String, Optional emailSubject As String, Optional forceIssue As Boolean = False, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    Dim tmp As Variant
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "memo", memo
    If emailSubject <> "" Then
        tmp.Add "emailSubject", emailSubject
    End If
    
    If forceIssue Then
        tmp.Add "forceIssue", True
    Else
        tmp.Add "forceIssue", False
    End If
    
    postData = m_Linkhub.toString(tmp)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, getSession_token(CorpNum), "ISSUE", postData, UserID)
    
    Set Issue = New PBResponse
    
    Issue.code = result.Item("code")
    Issue.message = result.Item("message")
    Issue.ntsConfirmNum = result.Item("ntsConfirmNum")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행예정
Public Function Send(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional emailSubject As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """, ""emailSubject"":""" + emailSubject + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, getSession_token(CorpNum), "SEND", postData, UserID)
    
    Set Send = New PBResponse
    
    Send.code = result.Item("code")
    Send.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행예정 취소
Public Function CancelSend(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "CANCELSEND", postData, UserID)
    
    Set CancelSend = New PBResponse
    
    CancelSend.code = result.Item("code")
    CancelSend.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발헁예정 승인
Public Function Accept(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "ACCEPT", postData, UserID)
    
    Set Accept = New PBResponse
    
    Accept.code = result.Item("code")
    Accept.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행예정 거부
Public Function Deny(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "DENY", postData, UserID)
    
    Set Deny = New PBResponse
    
    Deny.code = result.Item("code")
    Deny.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행 취소
Public Function CancelIssue(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "CANCELISSUE", postData, UserID)
    
    Set CancelIssue = New PBResponse
    
    CancelIssue.code = result.Item("code")
    CancelIssue.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function

'즉시요청
Public Function RegistRequest(CorpNum As String, Taxinvoice As PBTaxinvoice, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler

    Dim result As Variant
    Dim tmpDic As Variant
    
    Set tmpDic = Taxinvoice.toDictionary
        
    Dim postData As String
    
    postData = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST_override("/Taxinvoice", getSession_token(CorpNum), "REQUEST", postData, UserID)
    
    Set RegistRequest = New PBResponse
    
    RegistRequest.code = result.Item("code")
    RegistRequest.message = result.Item("message")
    
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set RegistRequest = Nothing
End Function

'역)발행요청
Public Function Request(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "REQUEST", postData, UserID)
    
    Set Request = New PBResponse
    
    Request.code = result.Item("code")
    Request.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'역)발행요청 취소
Public Function CancelRequest(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "CANCELREQUEST", postData, UserID)
    
    Set CancelRequest = New PBResponse
    
    CancelRequest.code = result.Item("code")
    CancelRequest.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'역)발행요청 거부
Public Function Refuse(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional memo As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""memo"":""" + memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "REFUSE", postData, UserID)
    
    Set Refuse = New PBResponse
    
    Refuse.code = result.Item("code")
    Refuse.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function

'국세청 전송
Public Function sendToNTS(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "NTS", "", UserID)
    
    Set sendToNTS = New PBResponse
    
    sendToNTS.code = result.Item("code")
    sendToNTS.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'이메일 재전송
Public Function SendEmail(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Receiver As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    postData = "{""receiver"":""" + Receiver + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "EMAIL", postData, UserID)
    
    Set SendEmail = New PBResponse
    
    SendEmail.code = result.Item("code")
    SendEmail.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'문자재선송
Public Function SendSMS(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Sender As String, Receiver As String, Contents As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    Dim T As Variant
    
    Set T = CreateObject("Scripting.Dictionary")
    
    T.Add "receiver", Receiver
    T.Add "sender", Sender
    T.Add "contents", Contents
    
    postData = m_Linkhub.toString(T)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "SMS", postData, UserID)
    
    Set SendSMS = New PBResponse
    
    SendSMS.code = result.Item("code")
    SendSMS.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'팩스 재전송
Public Function SendFax(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Sender As String, Receiver As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postData As String
    
    Dim T As Variant
    
    Set T = CreateObject("Scripting.Dictionary")
    
    T.Add "receiver", Receiver
    T.Add "sender", Sender
    
    postData = m_Linkhub.toString(T)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), "FAX", postData, UserID)
    
    Set SendFax = New PBResponse
    
    SendFax.code = result.Item("code")
    SendFax.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'상태/요약정보 확인
Public Function GetInfo(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As PBTIInfo
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey, _
                                getSession_token(CorpNum), UserID)
    
    Set GetInfo = New PBTIInfo
    
    
    GetInfo.fromDictionary result
    
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetInfo = Nothing
End Function

'상태/요역정보 다량(최대1000건)확인
Public Function GetInfos(CorpNum As String, KeyType As MgtKeyType, MgtKeyList As Collection, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    If MgtKeyList Is Nothing Or MgtKeyList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    Dim postData As String
    
    postData = m_Linkhub.toString(MgtKeyList)
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST("/Taxinvoice/" + strKeyType, _
                                getSession_token(CorpNum), postData, UserID)
    
    Set GetInfos = New Collection
    
    Dim T As Variant
    Dim info As PBTIInfo
    
    For Each T In result
        Set info = New PBTIInfo
        
        info.fromDictionary T
        
        GetInfos.Add info
    
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetInfos = Nothing
End Function
'세금계산서 목록조회
Public Function Search(CorpNum As String, KeyType As MgtKeyType, DType As String, SDate As String, EDate As String, state As Collection, TType As Collection, taxType As Collection, LateOnly As String, Page As Integer, PerPage As Integer, Order As String, Optional TaxRegIDType As String, Optional TaxRegID As String, Optional TaxRegIDYN As String, Optional QString As String, Optional UserID As String = "", Optional interOPYN As String = "", Optional issueType As Collection = Null, Optional regType As Collection = Null, Optional closeDownState As Collection = Null, Optional mgtKey As String = "")
On Error GoTo ErrHandler
    Dim result As Variant
    Dim strKeyType As String
    Dim uri As String
    Dim i As Integer
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    uri = "/Taxinvoice/" + strKeyType
    
    If Len(DType) = 0 Or IsNull(DType) Then
        Err.Raise -99999999, "POPBILL", "검색일자 유형이 입력되지 않았습니다"
    End If
    
    If Len(SDate) = 0 Or IsNull(SDate) Then
        Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다"
    End If
    
    If Len(EDate) = 0 Or IsNull(EDate) Then
        Err.Raise -99999999, "POPBILL", "종료일자가 입력되지 않았습니다"
    End If
    
    uri = uri + "?DType=" + DType
    uri = uri + "&SDate=" + SDate
    uri = uri + "&EDate=" + EDate
    
    If state.Count > 0 Then
        uri = uri + "&State="
        For i = 1 To state.Count
            uri = uri + state(i) + ","
        Next
        
    End If
        
    If TType.Count > 0 Then
        uri = uri + "&TType="
        For i = 1 To TType.Count
            uri = uri + TType(i) + ","
        Next
        
    End If
    
    If taxType.Count > 0 Then
        uri = uri + "&TaxType="
        For i = 1 To taxType.Count
            uri = uri + taxType(i) + ","
        Next
        
    End If
    
    If Not issueType Is Nothing Then
        uri = uri + "&IssueType="
        For i = 1 To issueType.Count
            uri = uri + issueType(i) + ","
        Next
    End If
    
    If Not regType Is Nothing Then
        uri = uri + "&RegType="
        For i = 1 To regType.Count
            uri = uri + regType(i) + ","
        Next
    End If
    
    If Not closeDownState Is Nothing Then
        uri = uri + "&CloseDownState="
        For i = 1 To closeDownState.Count
            uri = uri + closeDownState(i) + ","
        Next
    End If
    
    If IsNull(LateOnly) = False Then
        uri = uri + "&LateOnly=" + LateOnly
    End If
    
    If IsNull(TaxRegIDType) = False Then
        uri = uri + "&TaxRegIDType=" + TaxRegIDType
    End If
    
    If IsNull(TaxRegID) = False Then
        uri = uri + "&TaxRegID=" + TaxRegID
    End If
    
    If IsNull(TaxRegIDYN) = False Then
        uri = uri + "&TaxRegIDYN=" + TaxRegIDYN
    End If
    
    If QString <> "" Then
        uri = uri + "&QString=" + URLEncode(QString)
    End If
    
    If mgtKey <> "" Then
        uri = uri + "&MgtKey=" + URLEncode(mgtKey)
    End If
    
    If interOPYN <> "" Then
        uri = uri + "&InterOPYN=" + interOPYN
    End If
    
    Page = IIf(Page < 1, 1, Page)
    PerPage = IIf(PerPage < 1, 500, PerPage)
    uri = uri + "&Page=" + CStr(Page)
    uri = uri + "&PerPage=" + CStr(PerPage)
    uri = uri + "&Order=" + Order
    
    
    Set result = httpGET(uri, getSession_token(CorpNum), UserID)
    
    Set Search = New PBTISearchList
    
    
    Search.fromDictionary result
    
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Search = Nothing
    
End Function
'상세정보 확인
Public Function GetDetailInfo(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As PBTaxinvoice
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?Detail", _
                                getSession_token(CorpNum), UserID)
    
    Set GetDetailInfo = New PBTaxinvoice
    
    
    GetDetailInfo.fromDictionary result
    
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetDetailInfo = Nothing
End Function
'상세정보 확인(XML)
Public Function GetXML(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As PBTaxinvoiceXML
On Error GoTo ErrHandler
    Dim result As Variant
    Dim strKeyType As String

    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?XML", getSession_token(CorpNum), UserID)
     
    Set GetXML = New PBTaxinvoiceXML
    
    GetXML.fromDictionary result
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetXML = Nothing
End Function
'팝업URL 확인
Public Function GetPopUpURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=POPUP", _
                                getSession_token(CorpNum), UserID)
    
    GetPopUpURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopUpURL = ""
End Function
'팝업URL (메뉴/버튼 제외)
Public Function GetViewURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=VIEW", _
                                getSession_token(CorpNum), UserID)
    
    GetViewURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetViewURL = ""
End Function

'인쇄 URL확인
Public Function GetPrintURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=PRINT", _
                                getSession_token(CorpNum), UserID)
    
    GetPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPrintURL = ""
End Function

Public Function GetOldPrintURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=PRINTOLD", _
                                getSession_token(CorpNum), UserID)
    
    GetOldPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetOldPrintURL = ""
End Function

Public Function GetPDFURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=PDF", _
                                getSession_token(CorpNum), UserID)
    
    GetPDFURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPDFURL = ""
End Function
'공급받는자 인쇄 URL확인
Public Function GetEPrintURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=EPRINT", _
                                getSession_token(CorpNum), UserID)
    
    GetEPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetEPrintURL = ""
End Function

'메일 URL확인
Public Function GetMailURL(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "?TG=MAIL", _
                                getSession_token(CorpNum), UserID)
    
    GetMailURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetMailURL = ""
End Function
'다량인쇄 URL확인
Public Function GetMassPrintURL(CorpNum As String, KeyType As MgtKeyType, MgtKeyList As Collection, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If MgtKeyList Is Nothing Or MgtKeyList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    Dim postData As String
    
    postData = m_Linkhub.toString(MgtKeyList)
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST("/Taxinvoice/" + strKeyType + "?Print", _
                                getSession_token(CorpNum), postData, UserID)
    
    GetMassPrintURL = result.Item("url")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetMassPrintURL = ""
End Function

'문서이력확인
Public Function GetLogs(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "/Logs", _
                                getSession_token(CorpNum), UserID)
    
    Set GetLogs = New Collection
    
    Dim T As Variant
    Dim info As PBTILog
    
    For Each T In result
        Set info = New PBTILog
        
        info.fromDictionary T
        
        GetLogs.Add info
    
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetLogs = Nothing
End Function
'파일 첨부
Public Function AttachFile(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, FilePath As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_File("/Taxinvoice/" + strKeyType + "/" + mgtKey + "/Files", _
                        getSession_token(CorpNum), "", FilePath, UserID)
         
    Set AttachFile = New PBResponse
    AttachFile.code = result.Item("code")
    AttachFile.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set AttachFile = Nothing
End Function
'첨부파일 목록 확인
Public Function GetFiles(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + mgtKey + "/Files", _
                                getSession_token(CorpNum), UserID)
    
    Set GetFiles = New Collection
    
    Dim T As Variant
    Dim info As PBAttachFile
    
    For Each T In result
        Set info = New PBAttachFile
        
        info.fromDictionary T
        
        GetFiles.Add info
    
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetFiles = Nothing
End Function
'첨부파일 삭제
Public Function DeleteFile(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, FileID As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
     If FileID = "" Then
        Err.Raise -99999999, "POPBILL", "파일아이디가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + mgtKey + "/Files/" + FileID, _
                                    getSession_token(CorpNum), "DELETE", "", UserID)
    
    Set DeleteFile = New PBResponse
    
    DeleteFile.code = result.Item("code")
    DeleteFile.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set DeleteFile = Nothing
End Function
'유통메일목록 확인
Public Function GetEmailPublicKeys(CorpNum As String) As Collection
On Error GoTo ErrHandler
    
    Dim result As Variant
    
    Set result = httpGET("/Taxinvoice/EmailPublicKeys", _
                                getSession_token(CorpNum), "")
    
    Set GetEmailPublicKeys = New Collection
    
    Dim T As Variant
    
    For Each T In result
        
        GetEmailPublicKeys.Add T.Item("email")
    
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetEmailPublicKeys = Nothing
End Function
'전자명세서 첨부
Public Function AttachStatement(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, SubItemCode As Integer, SubMgtKey As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    
    Dim postData As String
    Dim tmp As Variant
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "ItemCode", SubItemCode
    tmp.Add "MgtKey", SubMgtKey
    
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
        
    postData = m_Linkhub.toString(tmp)
    
    Set result = httpPOST("/Taxinvoice/" + strKeyType + "/" + mgtKey + "/AttachStmt", getSession_token(CorpNum), postData)
    Set AttachStatement = New PBResponse
    
    AttachStatement.code = result.Item("code")
    AttachStatement.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set AttachStatement = Nothing
End Function
'전자명세서 첨부해제
Public Function DetachStatement(CorpNum As String, KeyType As MgtKeyType, mgtKey As String, SubItemCode As Integer, SubMgtKey As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    
    Dim postData As String
    Dim tmp As Variant
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "ItemCode", SubItemCode
    tmp.Add "MgtKey", SubMgtKey
    
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
        
    postData = m_Linkhub.toString(tmp)
    
    Set result = httpPOST("/Taxinvoice/" + strKeyType + "/" + mgtKey + "/DetachStmt", getSession_token(CorpNum), postData)
    Set DetachStatement = New PBResponse
    
    DetachStatement.code = result.Item("code")
    DetachStatement.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set DetachStatement = Nothing
End Function

'문서번호 할당
Public Function AssignMgtKey(CorpNum As String, KeyType As MgtKeyType, itemKey As String, mgtKey As String, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler

    Dim postData As String
    
    If Not Len(itemKey) = 18 Then
        Err.Raise -99999999, "POPBILL", "아이템키(itemKey)가 올바르지 않습니다."
    End If
    
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    postData = "MgtKey=" + mgtKey
    
    
    Set result = httpPOST_override("/Taxinvoice/" + itemKey + "/" + strKeyType, _
                                    getSession_token(CorpNum), _
                                    "", _
                                    postData, _
                                    UserID, _
                                    "application/x-www-form-urlencoded; charset=utf-8")
                                    
    Set AssignMgtKey = New PBResponse
    
    AssignMgtKey.code = result.Item("code")
    AssignMgtKey.message = result.Item("message")
    
    Exit Function

ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set AssignMgtKey = Nothing
End Function

'알림메일 전송목록 조회
Public Function ListEmailConfig(CorpNum As String, Optional UserID As String = "") As Collection
    On Error GoTo ErrHandler
    
    Dim result As Variant

    Set result = httpGET("/Taxinvoice/EmailSendConfig", getSession_token(CorpNum), UserID)
   
    Set ListEmailConfig = New Collection
    
    Dim T As Variant
    Dim info As PBEmailConfig
    
    For Each T In result
        Set info = New PBEmailConfig
        
        info.fromDictionary T
        
        ListEmailConfig.Add info
    Next
    
    Exit Function

ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set ListEmailConfig = Nothing
End Function

'알림메일 전송설정 수정
Public Function UpdateEmailConfig(CorpNum As String, emailType As String, sendYN As Boolean, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler
    
    Dim result As Variant

    If emailType = "" Then
        Err.Raise -99999999, "POPBILL", "메일전송 타입이 입력되지 않았습니다."
    End If
    
    
    Set result = httpPOST("/Taxinvoice/EmailSendConfig?EmailType=" + emailType + "&SendYN=" + CStr(sendYN), getSession_token(CorpNum), "", UserID)
   
    Set UpdateEmailConfig = New PBResponse
    
    UpdateEmailConfig.code = result.Item("code")
    UpdateEmailConfig.message = result.Item("message")
    
    Exit Function

ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set UpdateEmailConfig = Nothing
End Function

'국세청 전송 옵션 설정 상태 확인
Public Function GetSendToNTSConfig(CorpNum As String, Optional UserID As String = "") As PBSendToNTSConfig
    On Error GoTo ErrHandler
    
    Dim result As Variant
    
    Set result = httpGET("/Taxinvoice/SendToNTSConfig", getSession_token(CorpNum), UserID)
    
    Set GetSendToNTSConfig = New PBSendToNTSConfig
    
    GetSendToNTSConfig.sendToNTS = result.Item("sendToNTS")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetSendToNTSConfig = Nothing
End Function

' 인증서 정보 확인
Public Function GetTaxCertInfo(CorpNum As String, Optional UserID As String = "") As PBTaxinvoiceCertificate
    On Error GoTo ErrHandler
    
    Dim result As Variant
    
    Set result = httpGET("/Taxinvoice/Certificate", getSession_token(CorpNum), UserID)
    
    Set GetTaxCertInfo = New PBTaxinvoiceCertificate
    GetTaxCertInfo.fromDictionary result
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetTaxCertInfo = Nothing
End Function

'공동인증서 유효성 확인
Public Function CheckCertValidation(CorpNum As String, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler
    
    Dim result As Variant

    If CorpNum = "" Then
        Err.Raise -99999999, "POPBILL", "팝빌회원의 사업자번호가 입력되지 않았습니다.."
    End If
    
    
    Set result = httpGET("/Taxinvoice/CertCheck", getSession_token(CorpNum), UserID)
   
    Set CheckCertValidation = New PBResponse
    
    CheckCertValidation.code = result.Item("code")
    CheckCertValidation.message = result.Item("message")
    
    Exit Function

ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckCertValidation = Nothing
End Function

