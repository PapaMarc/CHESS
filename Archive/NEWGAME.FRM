VERSION 4.00
Begin VB.Form frmNewGame 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Chess Match"
   ClientHeight    =   2610
   ClientLeft      =   2070
   ClientTop       =   1725
   ClientWidth     =   4935
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   3015
   Left            =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4935
   Top             =   1380
   Width           =   5055
   Begin VB.Frame fraOpponent 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opponent:"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      Begin VB.TextBox txtSelectOpponent 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1540
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "Player2"
         Text            =   "<none selected>"
         Top             =   420
         Width           =   3030
      End
      Begin VB.CommandButton cmdAddress 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Select..."
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraMatch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Match &Name:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtSubject 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Tag             =   "Subject"
         Text            =   "Exchange Chess"
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1st &Move"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "NEWGAME.frx":0000
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddress_Click()
    Dim mRecip As Object
    Const kExchangeUser% = 0

    On Error GoTo ErrCmdAddress

    Set mRecip = gmSession.AddressBook(Nothing, kSelectOpponent, True, False, False, False, False, False, frmNewGame.hWnd)
    mRecip.Resolve
    If mRecip.Count > 0 Then
        If mRecip(1).address = "" Then
            MsgBox kErrCantFindRecipient, vbExclamation
        'ElseIf mRecip(1).DisplayType <> kExchangeUser Then
        '    MsgBox kErrNotExchangeUser, vbExclamation
        'ElseIf gmSession.CurrentUser.Name = gmNewGameRecips(1).Name Then
        '    MsgBox kErrCantPlayYourself, vbExclamation
        Else
            txtSelectOpponent = mRecip(1).Name
        End If
    End If
    gasPlayerName(kRed) = mRecip(1).Name
    gasPlayerAddress(kRed) = mRecip(1).address
    Set mRecip = Nothing
    cmdBegin.SetFocus
ErrCmdAddress:
End Sub

Private Sub cmdBegin_Click()
    Dim i%
    Dim mCurrentUser As Object
    
    If giCommand = kCompose Then
        If txtSelectOpponent <> "" Then
            Hide
            giCurTurn = kBlue
            
            Set mCurrentUser = gmSession.CurrentUser
            gasPlayerName(kBlue) = mCurrentUser.Name
            gasPlayerAddress(kBlue) = mCurrentUser.type & ":" & mCurrentUser.address
            '** red player name/address setup in cmdAddress
            Set mCurrentUser = Nothing
            
            frmGame!lblSubject.Caption = Chr(34) & txtSubject.Text & Chr(34)
            frmGame!lstMoves.Clear
            frmGame!txtComments = ""
            gsMoveData = ""
            gfByOpponent = True
            gfPlayed = False
            gsSubject = txtSubject
            giMoveNum = 0
            giEnPassant = -1
            giOldEnPassant = -1
            For i = 0 To 3
                gfRookMoved(i) = False
            Next i
            gfKingMoved(kBlue) = False
            gfKingMoved(kRed) = False

            Visible = False
            InitFrmGame
        End If
    Else
        Dim iErr%
    
        'go to first message of this conversation id
        iErr = iGetNewestMsg(gmLoadMsg)
        If (iErr) Then
            MsgBox kErrNoFirstMessage, vbExclamation
        Else
            gfByOpponent = gmLoadMsg.Fields.Item("ByOpponent").Value
            gfPlayed = gmLoadMsg.Fields.Item("Played").Value
            giMoveNum = 0
            Visible = False
            InitFrmGame
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    giUnloadMode = vbFormCode
    Terminate Me
End Sub










Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    giUnloadMode = UnloadMode
    If Not gfMapifVbxUnload Then
        Cancel = True
        Terminate Me
    End If
End Sub

Private Sub txtSelectOpponent_Change()
    If txtSelectOpponent.Text = "" Then
        cmdBegin.Enabled = False
    Else
        cmdBegin.Enabled = True
    End If
End Sub

Private Sub txtSelectOpponent_Click()
    If txtSelectOpponent.FontUnderline Then
        txtSelectOpponent.SelStart = 0
        txtSelectOpponent.SelLength = Len(txtSelectOpponent.Text)
    End If
End Sub

Private Sub txtSelectOpponent_KeyPress(KeyAscii As Integer)
    If txtSelectOpponent.FontUnderline Then
        txtSelectOpponent.Text = ""
        txtSelectOpponent.FontUnderline = False
    End If
End Sub



