Attribute VB_Name = "Timer"
' ==============================================================
' FileName:    Timer.bas
' Author:      Jesper Blomquist
' Date:        14 September 2001
'
' A module is needed for the setTimer API to be able to use a callback function
' ==============================================================

' To prevent object going out of scope whilst the timer fires:
Declare Function CoLockObjectExternal Lib "ole32" (ByVal _
      pUnk As IUnknown, ByVal fLock As Long, ByVal _
      fLastUnlockReleases As Long) As Long

' Timer API:
Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
      ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) _
      As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
      ByVal nIDEvent As Long) As Long

' Sleep API:
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private thisRunning As ICallerObj

' The ID of the API Timer:
Private m_lTimerID As Long

Private Sub TimerProc(ByVal lHwnd As Long, ByVal lMsg As Long, _
         ByVal lTimerID As Long, ByVal lTime As Long)

    'kill the timer
    KillTimer 0, m_lTimerID
    
    'tell the threadSrv object that it is time to go to work :)
    thisRunning.done 0
    
    'Unlock th object
    CoLockObjectExternal thisRunning, 0, 1
    
    Set thisRunning = Nothing
    m_lTimerID = 0
    
End Sub
Public Sub Start(this As ICallerObj)
    ' Ask the system to lock the object so that
    ' it will still perform its work even if it
    ' is released
    CoLockObjectExternal this, 1, 1

    'add this to thisRunning
    Set thisRunning = this
    
    ' Create a timer to start running the object:
    m_lTimerID = SetTimer(0, 0, 1, AddressOf TimerProc)
    
End Sub

