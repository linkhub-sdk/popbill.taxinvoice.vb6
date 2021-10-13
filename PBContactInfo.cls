VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBContactInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id               As String
Public Password         As String
Public pwd              As String
Public personName       As String
Public tel              As String
Public hp               As String
Public fax              As String
Public email            As String
Public regDT            As String
Public searchAllAllowYN As Boolean
Public mgrYN            As Boolean
Public state            As Integer
Public searchRole       As Integer

Public Sub fromDictionary(dic As Variant)
    id = dic.Item("id")
    Password = dic.Item("Password")
    pwd = dic.Item("pwd")
    personName = dic.Item("personName")
    tel = dic.Item("tel")
    hp = dic.Item("hp")
    fax = dic.Item("fax")
    email = dic.Item("email")
    regDT = dic.Item("regDT")
    searchAllAllowYN = CStr(dic.Item("searchAllAllowYN"))
    mgrYN = CStr(dic.Item("mgrYN"))
    state = dic.Item("state")
    searchRole = dic.Item("searchRole")
End Sub
