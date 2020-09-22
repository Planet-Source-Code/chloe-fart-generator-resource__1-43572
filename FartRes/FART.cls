VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum FartEnum
    fartLONG = 101
    fartTRUMPET = 102
    fartWETLONG = 103
    fartMUSTY = 104
End Enum
Public Sub PlayFart(f As FartEnum)
    On Error GoTo ErrorPlayFart
    PlayWaveRes f
    Exit Sub
ErrorPlayFart:
    Err.Raise Err.Number
    Exit Sub
End Sub
Public Sub RandomFart()
    On Error GoTo ErrorRandomFart
    Randomize
    Dim x As Long
    x = Int(Rnd * gMaxFarts) + 1
    x = x + 100
    PlayWaveRes x
    Exit Sub
ErrorRandomFart:
    Err.Raise Err.Number
    Exit Sub
End Sub

