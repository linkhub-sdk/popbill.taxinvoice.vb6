VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBTISearchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public code          As Integer
Public total         As Integer
Public PerPage       As Integer
Public pageNum       As Integer
Public pageCount     As Integer
Public message       As String


Public list As Collection


Public Sub fromDictionary(dic As Variant)

    code = dic.Item("code")
    total = dic.Item("total")
    PerPage = dic.Item("perPage")
    pageNum = dic.Item("pageNum")
    pageCount = dic.Item("pageCount")
    message = dic.Item("message")
        
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("list")
    
    If (tCollection Is Nothing) = False Then
        Set list = New Collection
        
        Dim T As Variant
        Dim searchList As PBTIInfo

        For Each T In tCollection
            Set searchList = New PBTIInfo
            
            searchList.fromDictionary T
            
            list.Add searchList
        Next
        
    End If
End Sub

