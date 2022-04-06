VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTaxinvoiceCertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public regDT            As String
Public expireDT         As String
Public issuerDN         As String
Public subjectDN        As String
Public issuerName       As String
Public oid              As String
Public regContactName   As String
Public regContactID     As String

Public Sub fromDictionary(dic As Variant)

    regDT = dic.Item("regDT")
    expireDT = dic.Item("expireDT")
    issuerDN = dic.Item("issuerDN")
    subjectDN = dic.Item("subjectDN")
    issuerName = dic.Item("issuerName")
    oid = dic.Item("oid")
    regContactName = dic.Item("regContactName")
    regContactID = dic.Item("regContactID")
    
End Sub
