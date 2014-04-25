VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTIDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public serialNum       As Integer
Public purchaseDT      As String
Public itemName        As String
Public spec            As String
Public qty             As String
Public unitCost        As String
Public supplyCost      As String
Public tax             As String
Public remark          As String



Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "serialNum", serialNum
    toDictionary.Add "purchaseDT", purchaseDT
    toDictionary.Add "itemName", itemName
    toDictionary.Add "spec", spec
    toDictionary.Add "qty", qty
    toDictionary.Add "unitCost", unitCost
    toDictionary.Add "supplyCost", supplyCost
    toDictionary.Add "tax", tax
    toDictionary.Add "remark", remark
    
End Function

Public Sub fromDictionary(dic As Variant)

    serialNum = dic.Item("serialNum")
    purchaseDT = dic.Item("purchaseDT")
    itemName = dic.Item("itemName")
    spec = dic.Item("spec")
    qty = dic.Item("qty")
    unitCost = dic.Item("unitCost")
    supplyCost = dic.Item("supplyCost")
    tax = dic.Item("tax")
    remark = dic.Item("remark")
    
End Sub
    
