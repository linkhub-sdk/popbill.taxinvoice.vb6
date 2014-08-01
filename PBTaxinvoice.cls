VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTaxinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public writeDate            As String
Public chargeDirection      As String
Public issueType            As String
Public issueTiming          As String
Public taxType              As String

Public invoicerCorpNum      As String
Public invoicerMgtKey       As String
Public invoicerTaxRegID     As String
Public invoicerCorpName     As String
Public invoicerCEOName      As String
Public invoicerAddr         As String
Public invoicerBizClass     As String
Public invoicerBizType      As String
Public invoicerContactName  As String
Public invoicerDeptName     As String
Public invoicerTEL          As String
Public invoicerHP           As String
Public invoicerEmail        As String
Public invoicerSMSSendYN    As Boolean

Public invoiceeType         As String
Public invoiceeCorpNum      As String
Public invoiceeMgtKey       As String
Public invoiceeTaxRegID     As String
Public invoiceeCorpName     As String
Public invoiceeCEOName      As String
Public invoiceeAddr         As String
Public invoiceeBizClass     As String
Public invoiceeBizType      As String
Public invoiceeContactName1 As String
Public invoiceeDeptName1    As String
Public invoiceeTEL1         As String
Public invoiceeHP1          As String
Public invoiceeEmail1       As String
Public invoiceeContactName2 As String
Public invoiceeDeptName2    As String
Public invoiceeTEL2         As String
Public invoiceeHP2          As String
Public invoiceeEmail2       As String
Public invoiceeSMSSendYN    As Boolean

Public trusteeCorpNum       As String
Public trusteeMgtKey        As String
Public trusteeTaxRegID      As String
Public trusteeCorpName      As String
Public trusteeCEOName       As String
Public trusteeAddr          As String
Public trusteeBizClass      As String
Public trusteeBizType       As String
Public trusteeContactName   As String
Public trusteeDeptName      As String
Public trusteeTEL           As String
Public trusteeHP            As String
Public trusteeEmail         As String
Public trusteeSMSSendYN     As Boolean

Public taxTotal             As String
Public supplyCostTotal      As String
Public totalAmount          As String
Public modifyCode           As String
Public orgNTSConfirmNum     As String
Public purposeType          As String
Public serialNum            As String
Public cash                 As String
Public chkBill              As String
Public credit               As String
Public note                 As String
Public remark1              As String
Public remark2              As String
Public remark3              As String
Public kwon                 As String
Public ho                   As String
Public businessLicenseYN    As Boolean
Public bankBookYN           As Boolean
Public faxsendYN            As Boolean
Public faxreceiveNum        As String
Public ntsconfirmNum        As String
Public originalTaxinvoiceKey As String
Public detailList As Collection
Public addContactList As Collection

Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "writeDate", writeDate
    
    toDictionary.Add "chargeDirection", chargeDirection
    toDictionary.Add "issueType", issueType
    toDictionary.Add "issueTiming", issueTiming
    toDictionary.Add "taxType", taxType
    toDictionary.Add "invoicerCorpNum", invoicerCorpNum
    toDictionary.Add "invoicerMgtKey", invoicerMgtKey
    toDictionary.Add "invoicerTaxRegID", invoicerTaxRegID
    toDictionary.Add "invoicerCorpName", invoicerCorpName
    toDictionary.Add "invoicerCEOName", invoicerCEOName
    toDictionary.Add "invoicerAddr", invoicerAddr
    toDictionary.Add "invoicerBizClass", invoicerBizClass
    toDictionary.Add "invoicerBizType", invoicerBizType
    toDictionary.Add "invoicerContactName", invoicerContactName
    toDictionary.Add "invoicerDeptName", invoicerDeptName
    toDictionary.Add "invoicerTEL", invoicerTEL
    toDictionary.Add "invoicerHP", invoicerHP
    toDictionary.Add "invoicerEmail", invoicerEmail
    toDictionary.Add "invoicerSMSSendYN", invoicerSMSSendYN
    
    toDictionary.Add "invoiceeCorpNum", invoiceeCorpNum
    toDictionary.Add "invoiceeType", invoiceeType
    toDictionary.Add "invoiceeMgtKey", invoiceeMgtKey
    toDictionary.Add "invoiceeTaxRegID", invoiceeTaxRegID
    toDictionary.Add "invoiceeCorpName", invoiceeCorpName
    toDictionary.Add "invoiceeCEOName", invoiceeCEOName
    toDictionary.Add "invoiceeAddr", invoiceeAddr
    toDictionary.Add "invoiceeBizType", invoiceeBizType
    toDictionary.Add "invoiceeBizClass", invoiceeBizClass
    toDictionary.Add "invoiceeContactName1", invoiceeContactName1
    toDictionary.Add "invoiceeDeptName1", invoiceeDeptName1
    toDictionary.Add "invoiceeTEL1", invoiceeTEL1
    toDictionary.Add "invoiceeHP1", invoiceeHP1
    toDictionary.Add "invoiceeEmail1", invoiceeEmail1
    toDictionary.Add "invoiceeContactName2", invoiceeContactName2
    toDictionary.Add "invoiceeDeptName2", invoiceeDeptName2
    toDictionary.Add "invoiceeTEL2", invoiceeTEL2
    toDictionary.Add "invoiceeEmail2", invoiceeEmail2
    toDictionary.Add "invoiceeSMSSendYN", invoiceeSMSSendYN
    
    toDictionary.Add "trusteeCorpNum", trusteeCorpNum
    toDictionary.Add "trusteeMgtKey", trusteeMgtKey
    toDictionary.Add "trusteeTaxRegID", trusteeTaxRegID
    toDictionary.Add "trusteeCorpName", trusteeCorpName
    toDictionary.Add "trusteeCEOName", trusteeCEOName
    toDictionary.Add "trusteeAddr", trusteeAddr
    toDictionary.Add "trusteeBizType", trusteeBizType
    toDictionary.Add "trusteeBizClass", trusteeBizClass
    toDictionary.Add "trusteeContactName", trusteeContactName
    toDictionary.Add "trusteeDeptName", trusteeDeptName
    toDictionary.Add "trusteeTEL", trusteeTEL
    toDictionary.Add "trusteeHP", trusteeHP
    toDictionary.Add "trusteeEmail", trusteeEmail
    toDictionary.Add "trusteeSMSSendYN", trusteeSMSSendYN
    
    toDictionary.Add "taxTotal", taxTotal
    toDictionary.Add "supplyCostTotal", supplyCostTotal
    toDictionary.Add "totalAmount", totalAmount
    If modifyCode <> "" Then
        toDictionary.Add "modifyCode", CInt(modifyCode)
    End If
    
    toDictionary.Add "orgNTSConfirmNum", orgNTSConfirmNum
    toDictionary.Add "purposeType", purposeType
    toDictionary.Add "serialNum", serialNum
    toDictionary.Add "cash", cash
    toDictionary.Add "chkBill", chkBill
    toDictionary.Add "credit", credit
    toDictionary.Add "note", note
    
    toDictionary.Add "kwon", kwon
    toDictionary.Add "ho", ho
    
    toDictionary.Add "businessLicenseYN", businessLicenseYN
    toDictionary.Add "bankBookYN", bankBookYN
    toDictionary.Add "faxsendYN", faxsendYN
    
    toDictionary.Add "remark1", remark1
    toDictionary.Add "remark2", remark2
    toDictionary.Add "remark3", remark3
    
    toDictionary.Add "faxreceiveNum", faxreceiveNum
    toDictionary.Add "ntsconfirmNum", ntsconfirmNum
    toDictionary.Add "originalTaxinvoiceKey", originalTaxinvoiceKey
    
    If (detailList Is Nothing) = False Then
        Dim t_detailCollection As New Collection
        Dim row As PBTIDetail
        For Each row In detailList
            t_detailCollection.Add row.toDictionary()
        Next
        If t_detailCollection.Count > 0 Then
            toDictionary.Add "detailList", t_detailCollection
        End If
    End If
    
    If (addContactList Is Nothing) = False Then
        Dim t_contactList As New Collection
        Dim contact As PBTIContact
        For Each contact In addContactList
            t_contactList.Add contact.toDictionary()
        Next
        If t_contactList.Count > 0 Then
            toDictionary.Add "addContactList", t_contactList
        End If

    End If
    
    
End Function

Public Sub fromDictionary(dic As Variant)

    writeDate = dic.Item("writeDate")
    chargeDirection = dic.Item("chargeDirection")
    issueType = dic.Item("issueType")
    issueTiming = dic.Item("issueTiming")
    taxType = dic.Item("taxType")
    
    invoicerCorpNum = dic.Item("invoicerCorpNum")
    invoicerMgtKey = dic.Item("invoicerMgtKey")
    invoicerTaxRegID = dic.Item("invoicerTaxRegID")
    invoicerCorpName = dic.Item("invoicerCorpName")
    invoicerCEOName = dic.Item("invoicerCEOName")
    invoicerAddr = dic.Item("invoicerAddr")
    invoicerBizClass = dic.Item("invoicerBizClass")
    invoicerBizType = dic.Item("invoicerBizType")
    invoicerContactName = dic.Item("invoicerContactName")
    invoicerDeptName = dic.Item("invoicerDeptName")
    invoicerTEL = dic.Item("invoicerTEL")
    invoicerHP = dic.Item("invoicerHP")
    invoicerEmail = dic.Item("invoicerEmail")
    invoicerSMSSendYN = dic.Item("invoicerSMSSendYN")
    
    invoiceeType = dic.Item("invoiceeType")
    invoiceeCorpNum = dic.Item("invoiceeCorpNum")
    invoiceeMgtKey = dic.Item("invoiceeMgtKey")
    invoiceeTaxRegID = dic.Item("invoiceeTaxRegID")
    invoiceeCorpName = dic.Item("invoiceeCorpName")
    invoiceeCEOName = dic.Item("invoiceeCEOName")
    invoiceeAddr = dic.Item("invoiceeAddr")
    invoiceeBizClass = dic.Item("invoiceeBizClass")
    invoiceeBizType = dic.Item("invoiceeBizType")
    invoiceeContactName1 = dic.Item("invoiceeContactName1")
    invoiceeDeptName1 = dic.Item("invoiceeDeptName1")
    invoiceeTEL1 = dic.Item("invoiceeTEL1")
    invoiceeHP1 = dic.Item("invoiceeHP1")
    invoiceeEmail1 = dic.Item("invoiceeEmail1")
    invoiceeContactName2 = dic.Item("invoiceeContactName2")
    invoiceeDeptName2 = dic.Item("invoiceeDeptName2")
    invoiceeTEL2 = dic.Item("invoiceeTEL2")
    invoiceeHP2 = dic.Item("invoiceeHP2")
    invoiceeEmail2 = dic.Item("invoiceeEmail2")
    invoiceeSMSSendYN = dic.Item("invoiceeSMSSendYN")
    
    trusteeCorpNum = dic.Item("trusteeCorpNum")
    trusteeMgtKey = dic.Item("trusteeMgtKey")
    trusteeTaxRegID = dic.Item("trusteeTaxRegID")
    trusteeCorpName = dic.Item("trusteeCorpName")
    trusteeCEOName = dic.Item("trusteeCEOName")
    trusteeAddr = dic.Item("trusteeAddr")
    trusteeBizClass = dic.Item("trusteeBizClass")
    trusteeBizType = dic.Item("trusteeBizType")
    trusteeContactName = dic.Item("trusteeContactName")
    trusteeDeptName = dic.Item("trusteeDeptName")
    trusteeTEL = dic.Item("trusteeTEL")
    trusteeHP = dic.Item("trusteeHP")
    trusteeEmail = dic.Item("trusteeEmail")
    trusteeSMSSendYN = dic.Item("trusteeSMSSendYN")
    
    taxTotal = dic.Item("taxTotal")
    supplyCostTotal = dic.Item("supplyCostTotal")
    totalAmount = dic.Item("totalAmount")
    modifyCode = dic.Item("modifyCode")
    orgNTSConfirmNum = dic.Item("orgNTSConfirmNum")
    purposeType = dic.Item("purposeType")
    serialNum = dic.Item("serialNum")
    cash = dic.Item("cash")
    chkBill = dic.Item("chkBill")
    credit = dic.Item("credit")
    note = dic.Item("note")
    remark1 = dic.Item("remark1")
    remark2 = dic.Item("remark2")
    remark3 = dic.Item("remark3")
    kwon = dic.Item("kwon")
    ho = dic.Item("ho")
    businessLicenseYN = dic.Item("businessLicenseYN")
    bankBookYN = dic.Item("bankBookYN")
    faxsendYN = dic.Item("faxsendYN")
    faxreceiveNum = dic.Item("faxreceiveNum")
    ntsconfirmNum = dic.Item("ntsconfirmNum")
    originalTaxinvoiceKey = dic.Item("originalTaxinvoiceKey")
    
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("detailList")
    
    If (tCollection Is Nothing) = False Then
        Set detailList = New Collection
        
        Dim t As Variant
        Dim detail As PBTIDetail
            
            
        For Each t In tCollection
            Set detail = New PBTIDetail
            
            detail.fromDictionary t
            
            detailList.Add detail
            
        Next
    
    End If
    
    Set tCollection = dic.Item("addContactList")
    
    If (tCollection Is Nothing) = False Then
        Set addContactList = New Collection
        
        Dim c As Variant
        Dim contact As PBTIContact
            
        For Each c In tCollection
            Set contact = New PBTIContact
            
            contact.fromDictionary c
            
            addContactList.Add contact
            
        Next
    
    End If
    
  
End Sub
