VERSION 5.00
Begin VB.Form frmtest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Do not run in ide Or a gpf will occur"
   ClientHeight    =   4845
   ClientLeft      =   3435
   ClientTop       =   2535
   ClientWidth     =   6285
   Icon            =   "frmtest.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmsmsgbox 
      Caption         =   "Show msgbox"
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdstopnothread 
      Caption         =   "Stop"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4200
      TabIndex        =   14
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdnonthread 
      Caption         =   "not thread"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdpausethread3 
      Appearance      =   0  'Flat
      Caption         =   "Pause thread 3"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdpausethread2 
      Appearance      =   0  'Flat
      Caption         =   "Pause thread 2"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdpausethread1 
      Appearance      =   0  'Flat
      Caption         =   "Pause thread 1"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdstoppro3 
      Caption         =   "Stop process 3"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdpro3 
      Caption         =   "Start process 3"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdstoppro2 
      Caption         =   "Stop process 2"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdstoppro1 
      Caption         =   "Stop Process 1"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdpro2 
      Caption         =   "Start process 2"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdpro11 
      Caption         =   "Start process 1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As multiple_threader.multithreader
Dim y As multiple_threader.multithreader
Dim z As multiple_threader.multithreader
Dim a As New multiple_threader.multithreader
Dim m_start As Boolean
Private Sub cmdabout_Click()
x.about
End Sub

'Private Sub cmdnew_Click()
'Dim x As New frmtest
'x.Show
'End Sub

Private Sub cmdnonthread_Click()
Dim i As Double
While m_start = True
DoEvents
i = i + 1
Text4 = i
Wend

End Sub

Private Sub cmdpausethread1_Click()
x.pausethread
End Sub

Private Sub cmdpausethread2_Click()
y.pausethread
End Sub

Private Sub cmdpausethread3_Click()
z.pausethread
End Sub

Private Sub cmdpro11_Click()
MsgBox x.m_thread_created
If x.m_thread_created Then
x.ResumetheThread
Else
x.startthread AddressOf loopthousand, False, th_idle

End If
End Sub

Private Sub cmdpro2_Click()
MsgBox y.m_thread_created
If y.m_thread_created Then
y.ResumetheThread
Else
y.startthread AddressOf loop100, False, th_idle

End If
End Sub

Private Sub cmdpro3_Click()
MsgBox z.m_thread_created
If z.m_thread_created Then
z.ResumetheThread
Else
z.startthread AddressOf loop1000, False, th_idle

End If
End Sub

Private Sub cmdstopnothread_Click()
m_start = Not m_start
End Sub

Private Sub cmdstoppro1_Click()
x.stopthread
End Sub




Private Sub cmdstoppro2_Click()
y.stopthread
End Sub

Private Sub cmdstoppro3_Click()
z.stopthread
End Sub


Private Sub cmsmsgbox_Click()
MsgBox "This will block the main thread but not other"
End Sub

Private Sub Form_Load()
m_start = True
Set x = New multiple_threader.multithreader
x.startthread AddressOf loopthousand, False, th_idle
'x.setotherthreadpriority App.ThreadID, th_idle
Set y = New multiple_threader.multithreader
y.startthread AddressOf loop100, False, th_idle

Set z = New multiple_threader.multithreader

z.startthread AddressOf loop1000, False, th_idle

'a.startthread AddressOf startnew

End Sub

Private Sub Form_Unload(Cancel As Integer)
x.stopthread
y.stopthread
z.stopthread
End Sub
