VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTIContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public serialNum       As Integer
Public email           As String
Public contactName     As String

Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "serialNum", serialNum
    toDictionary.Add "email", email
    toDictionary.Add "contactName", contactName
    
End Function

Public Sub fromDictionary(dic As Variant)

    serialNum = dic.Item("serialNum")
    email = dic.Item("email")
    contactName = dic.Item("contactName")
    
End Sub
    
