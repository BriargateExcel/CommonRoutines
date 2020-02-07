VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Timer comes from https://www.jkp-ads.com/Articles/performanceclass.asp

'Which item of gvPerfResults "belongs" to this instance?
Private pIndex As Long

'When did we start
Private pStartTime As Double

Private Sub Class_Initialize()
    'New item to keep track of, increase #
    IncrementPerformanceIndex
    
    'store which one is in this class instance
    pIndex = PerformanceIndex
    
    'Increase the depth to create an illusion of a call stack
    IncrementCallDepth
    
    If PerformanceResultsBounded Then
        ReDimPreservePerformanceResults 3, PerformanceIndex
    Else
        ReDimPerformanceResults 3, PerformanceIndex
    End If
    
    'Record when this instance was started
    pStartTime = dMicroTimer
End Sub

Public Sub SetRoutine(ByVal sRoutineName As String)
    SetPerformanceResults 1, pIndex, String(CallDepth * 4, " ") & sRoutineName
End Sub

Private Sub Class_Terminate()
    'Called automatically when the variable pointing to this
    'class instance goes out of scope
    
    'Outdent the call stack depth
    DecrementCallDepth
    
    'Record starttime and run-time
    SetPerformanceResults 2, pIndex, pStartTime
    SetPerformanceResults 3, pIndex, dMicroTimer - pStartTime
End Sub

