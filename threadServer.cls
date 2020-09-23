VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "threadSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==============================================================
' FileName:    threadSrv.cls
' Author:      Jesper Blomquist
' Date:        14 September 2001
'
' This is the actual object that will perform the work
' ==============================================================

Option Explicit

'implement the interface
Implements ICallerObj

' some variables to hold data handled to us
Private a As Long, b As Long

' a reference to the object that made the call for us
Private theCaller As ICallerObj

Public Function theCall(caller As ICallerObj, ByVal n1 As Long, ByVal n2 As Long) As Boolean

    ' save the args in private variables for later use
    a = n1: b = n2
    
    ' save a reference to the object that made the call
    Set theCaller = caller
    
    ' call Timer.bas - Start function
    Timer.Start Me
    
    ' return true if ok
    If Err.Number = 0 Then theCall = True
    
End Function

Private Sub processInfo()
    ' ==============================================================
    ' this is where you can put your own code
    ' ==============================================================
    
    Dim t As Long
    
    ' just a timer to sleep for 3 sec
    Sleep 3000
    t = a + b
    
    ' send the calculated answer back to the caller
    ' if you want to change the type of data that is returned (or the number of args) this
    ' must be done in the interface class (ICallerObj.cls)
    
    theCaller.done t
    
End Sub

Private Function ICallerObj_done(data As Long) As Variant
    ' call the sub where your code is
    processInfo
End Function
