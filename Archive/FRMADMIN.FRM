VERSION 4.00
Begin VB.Form frmAdmin 
   Caption         =   "Administration"
   ClientHeight    =   1080
   ClientLeft      =   2190
   ClientTop       =   2220
   ClientWidth     =   2415
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   1485
   Left            =   2130
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   2415
   Top             =   1875
   Visible         =   0   'False
   Width           =   2535
   Begin VBX.MAPIFormVBX MapifVbx1 
      Height          =   855
      Left            =   480
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit


Dim WindowCnt As Integer
Dim MasterFrm As Form
Dim fShutdown As Integer
Private Sub Form_Load()
    MapifVbx1.ClassID = kMsgGUID
    Set MasterFrm = Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gfMapifVbxUnload = True
    giUnloadMode = UnloadMode
    
    If UnloadMode = vbAppWindows Then
        fShutdown = True
    Else
        fShutdown = False
    End If
    
    Unload frmGame
    Unload frmNewGame
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gfMapifVbxUnload = True
    If 0 = WindowCnt Then
        Unload Me
    End If
    Set frmGame = Nothing
    Set frmNewGame = Nothing
    Set gmCurFolder = Nothing
    Set gmfObj = Nothing
    Set gmSession = Nothing
    Set gmLoadMsg = Nothing
    
    If fShutdown = False Then End
End Sub

Private Sub MapifVbx1_ClassFactoryFailure()
    MsgBox "Failed creating factory" ', vbStop
    Unload Me
End Sub

Private Sub MapifVbx1_Close()
    Terminate frmGame
    
    Unload frmGame
    Unload frmNewGame
End Sub


Private Sub MapifVbx1_DoVerb()
    '
    ' set the mapiform object
    '
    Call DoVerb(MapifVbx1.MapifObj, MapifVbx1.MapifObj.Verb)
End Sub

Private Sub MapifVbx1_OnActivateNext()
'
' check if this form can handle the next
' message.  If so, save the current message
' as necessary, and bring up the new message.
' Otherwise, let the form manager launch the
' message.
'
    'If MapifVbx1.MapifObj.NextMessageClass = FormMsgClass Then
    '    MapifVbx1.MapifObj.AcceptNext = True
    'Else
    '    MapifVbx1.MapifObj.AcceptNext = False
    'End If
    
End Sub




