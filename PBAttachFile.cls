VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBAttachFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public serialNum       As Integer
Public AttachedFile    As String
Public DisplayName     As String
Public regDT           As String


Public Sub fromDictionary(dic As Variant)

    serialNum = dic.Item("serialNum")
    AttachedFile = dic.Item("attachedFile")
    DisplayName = dic.Item("displayName")

    
    regDT = dic.Item("regDT")
  
End Sub
