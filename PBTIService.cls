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
' Author : Kim Seongjun (pallet027@gmail.com)
' Written : 2014-04-22
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID_REAL = "POPBILL"
Private Const ServiceID_TEST = "POPBILL_TEST"
Private Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Private Const ServiceURL_TEST = "https://popbill_test.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private m_IsTest As Boolean

Private m_Token As LinkhubToken
Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String

'관리번호 형태
Public Enum MgtKeyType
    SELL '매출
    BUY '매입
    TRUSTEE '수탁
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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
'초기화
Public Sub Initialize(linkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.linkID = linkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "member"
    ''세금계산서
    m_scope.Add "110"
    
    m_IsTest = False
    
End Sub
Private Function getSession_token(CorpNum As String) As String
    Dim refresh As Boolean
    
    refresh = False
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = Format(m_Linkhub.UTCTime, "yyyy-MM-ddTHH:mm:ss.sssZ")
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope)
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "POPBILL", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
    End If
    
    getSession_token = m_Token.session_token

End Function

Private Function httpGET(url As String, BearerToken As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    winhttp1.Send
    winhttp1.WaitForResponse
    
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

Private Function httpPOST(url As String, BearerToken As String, postdata As String, Optional UserID As String) As Variant
    Set httpPOST = httpPOST_override(url, BearerToken, "", postdata, UserID)
    
End Function

Private Function httpPOST_override(url As String, BearerToken As String, override As String, postdata As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    
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

Private Function GetOnlyFileName(ByVal FilePath As String) As String

     Dim Temp() As String
     Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))

End Function

Private Function httpPOST_File(url As String, BearerToken As String, postdata As String, FilePath As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    Dim boundary As String
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    
    If postdata <> "" Then
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
    winhttp1.WaitForResponse
    
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
'회원잔액조회
Public Function GetBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
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
'회원가입 여부
Public Function CheckIsMember(CorpNum As String, linkID As String) As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + linkID, "")
    
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
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "LinkID", JoinInfo.linkID
    tmp.Add "CorpNum", JoinInfo.CorpNum
    tmp.Add "CorpName", JoinInfo.CorpName
    tmp.Add "CEOName", JoinInfo.CEOName
    tmp.Add "Addr", JoinInfo.Addr
    tmp.Add "ZipCode", JoinInfo.ZipCode
    tmp.Add "BizClass", JoinInfo.BizClass
    tmp.Add "BizType", JoinInfo.BizType
    tmp.Add "ContactName", JoinInfo.contactName
    tmp.Add "ContactEmail", JoinInfo.ContactEmail
    tmp.Add "ContactFAX", JoinInfo.ContactFAX
    tmp.Add "ContactHP", JoinInfo.ContactHP
    tmp.Add "ContactTEL", JoinInfo.ContactTEL
    tmp.Add "ID", JoinInfo.ID
    tmp.Add "PWD", JoinInfo.PWD
    
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/Join", "", postdata)
    
    JoinMember = New PBResponse
    
    JoinMember.code = result.Item("code")
    JoinMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
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
Public Function CheckMgtKeyInUse(CorpNum As String, KeyType As MgtKeyType, MgtKey As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey, getSession_token(CorpNum))
    
    Set CheckMgtKeyInUse = New PBResponse
    
    CheckMgtKeyInUse.code = IIf(IsNull(result.Item("itemKey")), 0, 1)
    CheckMgtKeyInUse.message = IIf(IsNull(result.Item("itemKey")), "미사용중", "사용중")
    
    Exit Function
ErrHandler:
    If Err.Number = -11000005 Then
        Set CheckMgtKeyInUse = New PBResponse
        CheckMgtKeyInUse.code = 0
        CheckMgtKeyInUse.message = "미사용중"
        Exit Function
    End If
    
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckMgtKeyInUse = Nothing
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
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST("/Taxinvoice", getSession_token(CorpNum), postdata, UserID)
    
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
Public Function Update(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Taxinvoice As PBTaxinvoice, Optional writeSpecification As Boolean = False, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If

    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim tmpDic As Variant
    
    Set tmpDic = Taxinvoice.toDictionary
    If writeSpecification Then
        tmpDic.Add "writeSpecification", True
    End If
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                    getSession_token(CorpNum), _
                                    "PATCH", _
                                    postdata, _
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
Public Function Delete(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, getSession_token(CorpNum), "DELETE", "", UserID)
    
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
Public Function Issue(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Memo As String, Optional EmailSubject As String, Optional ForceIssue As Boolean = False, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    Dim tmp As Variant
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "memo", Memo
    If EmailSubject <> "" Then
        tmp.Add "emailSubject", EmailSubject
    End If
    
    If ForceIssue Then
        tmp.Add "forceIssue", True
    End If
    
    postdata = m_Linkhub.toString(tmp)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, getSession_token(CorpNum), "ISSUE", postdata, UserID)
    
    Set Issue = New PBResponse
    
    Issue.code = result.Item("code")
    Issue.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행예정
Public Function Send(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, getSession_token(CorpNum), "SEND", postdata, UserID)
    
    Set Send = New PBResponse
    
    Send.code = result.Item("code")
    Send.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행예정 취소
Public Function CancelSend(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "CANCELSEND", postdata, UserID)
    
    Set CancelSend = New PBResponse
    
    CancelSend.code = result.Item("code")
    CancelSend.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발헁예정 승인
Public Function Accept(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "ACCEPT", postdata, UserID)
    
    Set Accept = New PBResponse
    
    Accept.code = result.Item("code")
    Accept.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행예정 거부
Public Function Deny(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "DENY", postdata, UserID)
    
    Set Deny = New PBResponse
    
    Deny.code = result.Item("code")
    Deny.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'발행 취소
Public Function CancelIssue(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "CANCELISSUE", postdata, UserID)
    
    Set CancelIssue = New PBResponse
    
    CancelIssue.code = result.Item("code")
    CancelIssue.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'역)발행요청
Public Function Request(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "REQUEST", postdata, UserID)
    
    Set Request = New PBResponse
    
    Request.code = result.Item("code")
    Request.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'역)발행요청 취소
Public Function CancelRequest(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "CANCELREQUEST", postdata, UserID)
    
    Set CancelRequest = New PBResponse
    
    CancelRequest.code = result.Item("code")
    CancelRequest.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'역)발행요청 거부
Public Function Refuse(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "REFUSE", postdata, UserID)
    
    Set Refuse = New PBResponse
    
    Refuse.code = result.Item("code")
    Refuse.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function

'국세청 전송
Public Function SendToNTS(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "NTS", "", UserID)
    
    Set SendToNTS = New PBResponse
    
    SendToNTS.code = result.Item("code")
    SendToNTS.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'이메일 재전송
Public Function SendEmail(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Receiver As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    postdata = "{""receiver"":""" + Receiver + """}"
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "EMAIL", postdata, UserID)
    
    Set SendEmail = New PBResponse
    
    SendEmail.code = result.Item("code")
    SendEmail.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'문자재선송
Public Function SendSMS(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Sender As String, Receiver As String, Contents As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    Dim T As Variant
    
    Set T = CreateObject("Scripting.Dictionary")
    
    T.Add "receiver", Receiver
    T.Add "sender", Sender
    T.Add "contents", Contents
    
    postdata = m_Linkhub.toString(T)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "SMS", postdata, UserID)
    
    Set SendSMS = New PBResponse
    
    SendSMS.code = result.Item("code")
    SendSMS.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'팩스 재전송
Public Function SendFax(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Sender As String, Receiver As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Dim postdata As String
    
    Dim T As Variant
    
    Set T = CreateObject("Scripting.Dictionary")
    
    T.Add "receiver", Receiver
    T.Add "sender", Sender
    
    postdata = m_Linkhub.toString(T)
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
                                getSession_token(CorpNum), "FAX", postdata, UserID)
    
    Set SendFax = New PBResponse
    
    SendFax.code = result.Item("code")
    SendFax.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'상태/요약정보 확인
Public Function GetInfo(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional UserID As String) As PBTIInfo
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey, _
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
Public Function GetInfos(CorpNum As String, KeyType As MgtKeyType, MgtKeyList As Collection, Optional UserID As String) As Collection
On Error GoTo ErrHandler
    If MgtKeyList Is Nothing Or MgtKeyList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(MgtKeyList)
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST("/Taxinvoice/" + strKeyType, _
                                getSession_token(CorpNum), postdata, UserID)
    
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
'상세정보 확인
Public Function GetDetailInfo(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional UserID As String) As PBTaxinvoice
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "?Detail", _
                                getSession_token(CorpNum), UserID)
    
    Set GetDetailInfo = New PBTaxinvoice
    
    
    GetDetailInfo.fromDictionary result
    
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetDetailInfo = Nothing
End Function
'팝업URL 확인
Public Function GetPopUpURL(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "?TG=POPUP", _
                                getSession_token(CorpNum), UserID)
    
    GetPopUpURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopUpURL = ""
End Function
'인쇄 URL확인
Public Function GetPrintURL(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "?TG=PRINT", _
                                getSession_token(CorpNum), UserID)
    
    GetPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPrintURL = ""
End Function
'공급받는자 인쇄 URL확인
Public Function GetEPrintURL(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "?TG=EPRINT", _
                                getSession_token(CorpNum), UserID)
    
    GetEPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetEPrintURL = ""
End Function

'메일 URL확인
Public Function GetMailURL(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "?TG=MAIL", _
                                getSession_token(CorpNum), UserID)
    
    GetMailURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetMailURL = ""
End Function
'다량인쇄 URL확인
Public Function GetMassPrintURL(CorpNum As String, KeyType As MgtKeyType, MgtKeyList As Collection, UserID As String) As String
On Error GoTo ErrHandler
    If MgtKeyList Is Nothing Or MgtKeyList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(MgtKeyList)
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST("/Taxinvoice/" + strKeyType + "?Print", _
                                getSession_token(CorpNum), postdata, UserID)
    
    GetMassPrintURL = result.Item("url")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetMassPrintURL = ""
End Function

'문서이력확인
Public Function GetLogs(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional UserID As String) As Collection
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "/Logs", _
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
'파일 첩무
Public Function AttachFile(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, FilePath As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_File("/Taxinvoice/" + strKeyType + "/" + MgtKey + "/Files", _
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
Public Function GetFiles(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, Optional UserID As String) As Collection
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpGET("/Taxinvoice/" + strKeyType + "/" + MgtKey + "/Files", _
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
Public Function DeleteFile(CorpNum As String, KeyType As MgtKeyType, MgtKey As String, FileID As String, Optional UserID As String) As Collection
On Error GoTo ErrHandler
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
     If FileID = "" Then
        Err.Raise -99999999, "POPBILL", "파일아이디가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    Dim strKeyType As String
    
    strKeyType = IIf(KeyType = TRUSTEE, "TRUSTEE", IIf(KeyType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST_override("/Taxinvoice/" + strKeyType + "/" + MgtKey + "/Files/" + FileID, _
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
