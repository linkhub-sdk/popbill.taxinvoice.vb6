VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBBulkTaxinvoiceIssueResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public invoicerMgtKey      As String
Public trusteeMgtKey       As String
Public code                As Long
Public message             As String
Public ntsConfirmNum       As String
Public issueDT             As String

Public Sub fromDictinary(dic As Variant)
    invoicerMgtKey = dic.Item("invoicerMgtKey")
    trusteeMgtKey = dic.Item("trusteeMgtKey")
    code = dic.Item("code")
    message = dic.Item("message")
    ntsConfirmNum = dic.Item("ntsconfirmNum")
    issueDT = dic.Item("issueDT")
End Sub
