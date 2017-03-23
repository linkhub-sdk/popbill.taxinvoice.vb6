VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTIInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public closeDownState          As Integer
Public closeDownStateDate      As String

Public itemKey                 As String
Public stateCode               As Integer
Public taxType                 As String
Public purposeType             As String
Public modifyCode              As String
Public issueType               As String

Public writeDate               As String

Public invoicerCorpName        As String
Public invoicerCorpNum         As String
Public invoicerMgtKey          As String
Public invoicerPrintYN         As Boolean

Public invoiceeCorpName        As String
Public invoiceeCorpNum         As String
Public invoiceeMgtKey          As String
Public invoiceePrintYN         As Boolean

Public trusteeCorpName         As String
Public trusteeCorpNum          As String
Public trusteeMgtKey           As String
Public trusteePrintYN          As Boolean

Public supplyCostTotal         As String
Public taxTotal                As String

Public lateIssueYN             As Boolean
Public interOPYN               As Boolean
Public issueDT                 As String
Public preIssueDT              As String
Public stateDT                 As String
Public openYN                  As Boolean
Public openDT                  As String

Public ntsresult               As String
Public ntsconfirmNum           As String
Public ntssendDT               As String
Public ntsresultDT             As String
Public ntssendErrCode          As String

Public stateMemo               As String

Public regDT                   As String

Public Sub fromDictionary(dic As Variant)
    closeDownState = dic.Item("closeDownState")
    closeDownStateDate = dic.Item("closeDownStateDate")
    
    itemKey = dic.Item("itemKey")
    stateCode = dic.Item("stateCode")
    taxType = dic.Item("taxType")
    purposeType = dic.Item("purposeType")
    modifyCode = dic.Item("modifyCode")
    issueType = dic.Item("issueType")
     
    writeDate = dic.Item("writeDate")
    
    invoicerCorpName = dic.Item("invoicerCorpName")
    invoicerCorpNum = dic.Item("invoicerCorpNum")
    invoicerMgtKey = dic.Item("invoicerMgtKey")
    invoicerPrintYN = dic.Item("invoicerPrintYN")
    
    invoiceeCorpName = dic.Item("invoiceeCorpName")
    invoiceeCorpNum = dic.Item("invoiceeCorpNum")
    invoiceeMgtKey = dic.Item("invoiceeMgtKey")
    invoiceePrintYN = dic.Item("invoiceePrintYN")
    
    trusteeCorpName = dic.Item("trusteeCorpName")
    trusteeCorpNum = dic.Item("trusteeCorpNum")
    trusteeMgtKey = dic.Item("trusteeMgtKey")
    trusteePrintYN = dic.Item("trusteePrintYN")
    
    supplyCostTotal = dic.Item("supplyCostTotal")
    taxTotal = dic.Item("taxTotal")
    
    issueDT = dic.Item("issueDT")
    preIssueDT = dic.Item("preIssueDT")
    stateDT = dic.Item("stateDT")
    lateIssueYN = dic.Item("lateIssueYN")
    interOPYN = dic.Item("interOPYN")
    openYN = dic.Item("openYN")
    openDT = dic.Item("openDT")
    
    ntsresult = dic.Item("ntsresult")
    ntsconfirmNum = dic.Item("ntsconfirmNum")
    ntssendDT = dic.Item("ntssendDT")
    ntsresultDT = dic.Item("ntsresultDT")
    ntssendErrCode = dic.Item("ntssendErrCode")
    
    stateMemo = dic.Item("stateMemo")
    
    regDT = dic.Item("regDT")
End Sub

