VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Height          =   525
      Left            =   1710
      TabIndex        =   0
      Top             =   1170
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' FileName:    frmMain.frm
' Author:      Jesper Blomquist
' Date:        14 September 2001
'
' Very simple demo of a client using a COM server to thread processes
'
' Remember to add a reference to ThreadServer (ThreadSrv.exe)
' ==============================================================

Implements threadServer.ICallerObj
Private Sub Command1_Click()
    ' create an array of objects
    Dim server(10) As New threadServer.threadSrv
    Dim a As Long, b As Long

    b = 10
    For a = 0 To UBound(server)
        ' make the call call to the treaded object
        Call server(a).theCall(Me, a, b)
        Debug.Print "Client ready. Time: " & Timer
    Next a
    
End Sub

Private Function ICallerObj_done(data As Long) As Variant
    ' this is the treaded objects "callback" function (not really a callback function)
    Debug.Print "Server ready. Data: " & data & " Time: " & Timer
End Function
