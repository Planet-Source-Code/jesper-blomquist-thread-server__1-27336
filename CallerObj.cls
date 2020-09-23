VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICallerObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==============================================================
' FileName:    CallerObj.cls
' Author:      Jesper Blomquist
' Date:        14 September 2001
'
' This is the interface that must be implemented by both the threadSrv obj
' and the the client object
' ==============================================================

Public Function done(data As Long)

End Function
