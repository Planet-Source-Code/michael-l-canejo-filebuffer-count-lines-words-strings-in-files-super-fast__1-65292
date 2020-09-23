VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBenchMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStartTime As Long
Private mFinishTime As Long
Private mElapsedTime As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
    End Type
    Private LIPerfFrequency As LARGE_INTEGER
    Private LICounterStart As LARGE_INTEGER
    Private LICounterEnd As LARGE_INTEGER
    Private LIcrFrequency As Currency
    Private bEnabled As Boolean

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub Start()
    If bEnabled Then
        Call QueryPerformanceCounter(LICounterStart)
    Else
        mStartTime = timeGetTime
    End If
End Sub


Public Sub Finish()
    If bEnabled Then
        Call QueryPerformanceCounter(LICounterEnd)
    Else
        mFinishTime = timeGetTime
        mElapsedTime = mFinishTime - mStartTime
    End If
End Sub

Public Property Get ElapsedTime() As Double
    If bEnabled Then
        Dim crStart As Currency
        Dim crStop As Currency
        Dim crFrequency As Currency
        crStart = Large2Currency(LICounterStart)
        crStop = Large2Currency(LICounterEnd)
        ElapsedTime = ((crStop - crStart) / LIcrFrequency)
    Else
        ElapsedTime = mElapsedTime / 1000
    End If
End Property

Private Function Large2Currency(largeInt As LARGE_INTEGER) As Currency
    If (largeInt.lowpart) > 0& Then
        Large2Currency = largeInt.lowpart
    Else
        Large2Currency = CCur(2 ^ 31) + CCur(largeInt.lowpart And &H7FFFFFFF)
    End If
    Large2Currency = Large2Currency + largeInt.highpart * CCur(2 ^ 32)
End Function

Private Sub Class_Initialize()
    bEnabled = QueryPerformanceFrequency(LIPerfFrequency)
    If bEnabled Then
        LIcrFrequency = Large2Currency(LIPerfFrequency)
    End If
End Sub

Private Sub Class_Terminate()
    bEnabled = False
    mStartTime = 0
    mFinishTime = 0
    mElapsedTime = 0
End Sub
