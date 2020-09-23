VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "multithreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_threadhandle As Long
Public m_thread_created As Boolean
Enum threadpriority
th_normal = THREAD_PRIORITY_NORMAL
th_abovenormal = THREAD_PRIORITY_ABOVE_NORMAL
th_belownormal = THREAD_PRIORITY_BELOW_NORMAL
th_idle = THREAD_PRIORITY_IDLE
th_lowest = THREAD_PRIORITY_LOWEST
th_critical = THREAD_PRIORITY_TIME_CRITICAL
th_highest = THREAD_PRIORITY_HIGHEST
End Enum

Public Sub startthread(v_addressoffunction As Long, Optional v_immediately As Boolean = True, Optional v_setpriority As threadpriority = threadpriority.th_normal)
Attribute startthread.VB_Description = "This function takes the address of the  function that is to be executed in a seperate thread use address of operator in vb make sure the function does have  doevents statement in the loop"
 Dim l_tid As Long
If Not m_thread_created Then
 If v_addressoffunction > 0 Then
'm_thread_created = True
  If v_immediately Then
   m_threadhandle = CreateThread(ByVal 0&, ByVal 0&, v_addressoffunction, ByVal 0&, 0&, l_tid)
  Else
'use resume thread for execution
   m_threadhandle = CreateThread(ByVal 0&, ByVal 0&, v_addressoffunction, ByVal 0&, CREATE_SUSPENDED, l_tid)
  End If

     If m_threadhandle > 0 Then
       Call setpriority(v_setpriority)
       m_thread_created = True
     End If

  End If
End If
End Sub

Public Sub stopthread()
Dim l_ret As Long
If m_thread_created Then
l_ret = TerminateThread(m_threadhandle, 0&)
m_thread_created = False
End If
End Sub

Public Sub pausethread()
If m_thread_created Then
Dim l_ret As Long
l_ret = SuspendThread(m_threadhandle)
End If
End Sub

Public Sub ResumetheThread()
If m_thread_created Then
Dim l_ret As Long
l_ret = ResumeThread(m_threadhandle)
End If
End Sub
Public Sub setpriority(Optional ByVal v_priority As threadpriority = threadpriority.th_normal)
If m_thread_created Then
Dim l_ret As Long
l_ret = SetThreadPriority(m_threadhandle, v_priority)

End If
End Sub

Private Sub Class_Initialize()
m_thread_created = False
End Sub

Private Sub Class_Terminate()

If m_threadhandle > 0 Then
Dim l_ret As Long
l_ret = TerminateThread(m_threadhandle, 0&)
End If
End Sub

Public Sub setotherthreadpriority(ByVal v_threadid As Long, Optional ByVal v_priority As threadpriority = threadpriority.th_normal)
Dim l_ret As Long
l_ret = SetThreadPriority(v_threadid, v_priority)


End Sub


Public Sub about()
Attribute about.VB_Description = "shows about"
Attribute about.VB_UserMemId = -552
ShellAbout GetDesktopWindow(), "Multi threader dll", "Multi threader", 0&
End Sub
