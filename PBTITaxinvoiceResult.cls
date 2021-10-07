VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTIBulkTaxinvoiceResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public code                    As Integer
Public message                 As String
Public submitID                As String
Public submitCount             As Integer
Public successCount            As Integer
Public failCount               As Integer
Public txState                 As Integer
Public txResultCode            As Integer
Public txStartDT               As String
Public txEndDT                 As String
Public receiptDT               As String
Public receiptID               As String

Public issueResult             As Collection

Public Sub fromDictionary(dic As Variant)
    code = dic.Item("code")
    message = dic.Item("message")
    submitID = dic.Item("submitID")
    submitCount = dic.Item("submitCount")
    successCount = dic.Item("successCount")
    failCount = dic.Item("failCount")
    txState = dic.Item("txState")
    txResultCode = dic.Item("txResultCode")
    txStartDT = dic.Item("txStartDT")
    txEndDT = dic.Item("txEndDT")
    receiptDT = dic.Item("receiptDT")
    receiptID = dic.Item("receiptID")
    
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("issueResult")
    If (tCollection Is Nothing) = False Then
        Set issueResult = New Collection
        
        Dim T As Variant
        Dim bulkTIIssueResult As PBTIBulkTaxinvoiceIssueResult
        
        For Each T In tCollection
            Set bulkTIIssueResult = New PBTIBulkTaxinvoiceIssueResult
            
            bulkTIIssueResult.fromDictinary T
            issueResult.Add bulkTIIssueResult
        Next
        
    End If
End Sub
