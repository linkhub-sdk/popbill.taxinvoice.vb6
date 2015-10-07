VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCorpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ceoname            As String
Public corpName           As String
Public addr               As String
Public bizType            As String
Public bizClass           As String

Public Sub fromDictionary(dic As Variant)
    ceoname = dic.Item("ceoname")
    corpName = dic.Item("corpName")
    addr = dic.Item("addr")
    bizType = dic.Item("bizType")
    bizClass = dic.Item("bizClass")
End Sub



