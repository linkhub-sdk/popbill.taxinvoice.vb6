VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBBulkSubmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public forceIssue As Boolean
Public invoices As New Collection


Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    If forceIssue Then toDictionary.Add "forceIssue", True
    
    If (invoices Is Nothing) = False Then
        Dim t_invoiceCollection As New Collection
        Dim row As PBTaxinvoice
        For Each row In invoices
            t_invoiceCollection.Add row.toDictionary()
        Next
        If t_invoiceCollection.Count > 0 Then
            toDictionary.Add "invoices", t_invoiceCollection
        End If
    End If
End Function

