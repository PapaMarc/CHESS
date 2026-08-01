VERSION 4.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microsoft Exchange Chess"
   ClientHeight    =   4905
   ClientLeft      =   5760
   ClientTop       =   2745
   ClientWidth     =   5520
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   5310
   Icon            =   "GAME.frx":0000
   KeyPreview      =   -1  'True
   Left            =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   Top             =   2400
   Width           =   5640
   Begin VB.PictureBox picPlayers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   5295
   End
   Begin VB.Timer tmrCheck 
      Interval        =   600
      Left            =   4800
      Top             =   4200
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Surrender"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ListBox lstMoves 
      Height          =   1620
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgKingChmate 
      Height          =   480
      Index           =   2
      Left            =   1440
      Picture         =   "GAME.frx":030A
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKingChmate 
      Height          =   480
      Index           =   1
      Left            =   1440
      Picture         =   "GAME.frx":0614
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKingCh 
      Height          =   480
      Index           =   2
      Left            =   2520
      Picture         =   "GAME.frx":091E
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKingCh 
      Height          =   480
      Index           =   1
      Left            =   2520
      Picture         =   "GAME.frx":0C28
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Comments:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label lblMoves 
      BackStyle       =   0  'Transparent
      Caption         =   "Move &History:"
      Height          =   492
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   852
   End
   Begin VB.Label lblSubject 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subject"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.Image imgRook 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "GAME.frx":0F32
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgQueen 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "GAME.frx":123C
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKnight 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "GAME.frx":1546
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPawn 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "GAME.frx":1850
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKing 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "GAME.frx":1B5A
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBishop 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "GAME.frx":1E64
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTemp 
      Height          =   372
      Left            =   600
      Top             =   840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Image imgBishop 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "GAME.frx":216E
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKing 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "GAME.frx":2478
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPawn 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "GAME.frx":2782
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKnight 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "GAME.frx":2A8C
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgQueen 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "GAME.frx":2D96
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRook 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "GAME.frx":30A0
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Height          =   384
      Index           =   0
      Left            =   120
      Top             =   840
      Width           =   384
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdSend_Click()
    SendMessage
End Sub

Private Sub cmdUndo_Click()
    Dim fSurrender%

    If cmdUndo.Caption = kUndo Then
        If Len(gsMoveData) <= 6 Then
            gsMoveData = ""
        Else
            gsMoveData = Left(gsMoveData, Len(gsMoveData) - 1) 'strip off ;
            While Right(gsMoveData, 1) <> ";"
                gsMoveData = Left(gsMoveData, Len(gsMoveData) - 1) 'other chars
            Wend
        End If
        lstMoves.RemoveItem lstMoves.ListCount - 1
        cmdSend.Enabled = False
        gfMovedPiece = False
        gsMoveNotation = ""
        If gfGameOver = True Then
            gfGameOver = False
            imgSquare(giKingIndex(giCurTurn)) = imgKing(giCurTurn)
        End If
        giEnPassant = giOldEnPassant
        ReplayMoves lstMoves.ListCount
        ResetKings
        If lstMoves.ListCount > 0 Then
            lstMoves.Tag = "x"
            lstMoves.Selected(lstMoves.ListCount - 1) = True
            lstMoves.Tag = ""
        End If
        '** if undo'ing a surrender action
        If gfGameOver = True Then gfGameOver = False
    Else
        fSurrender = MsgBox(kSurrenderPrompt, vbYesNo)
        If fSurrender = vbYes Then
            gfGameOver = True
            gfMovedPiece = True
            'cmdUndo.Caption = kUndo
            gsMoveData = gsMoveData & "!,S," & giCurTurn & ";"
            gsMoveNotation = kSurrenderNotation
            lstMoves.AddItem gsMoveNotation
            cmdSend.Enabled = True
            'cmdUndo.Enabled = True
            lstMoves.Tag = "x" 'don't activate click event
            lstMoves.Selected(lstMoves.ListCount - 1) = True
            lstMoves.Tag = "" 'allow click event
            imgSquare(giKingIndex(giCurTurn)) = imgKingChmate(giCurTurn)
        End If
    End If
    AdjustUndoButton
End Sub

Private Sub DrawBoard()
    Dim i%, j%, color%
    Dim iWidth, iHeight
    
    iWidth = 32 'scalewidth / 8
    iHeight = 32 ' scaleheight / 8

    For i = 0 To 7
        For j = 0 To 7
            imgSquare((i * 8) + j).Left = j * iWidth + giHorizBoardOffset
            imgSquare((i * 8) + j).Top = i * iHeight + giVertBoardOffset
            
            color = Not color
            If color = 0 Then
                Line (j * iWidth + giHorizBoardOffset, i * iHeight + giVertBoardOffset)-((j + 1) * iWidth + giHorizBoardOffset, (i + 1) * iHeight + giVertBoardOffset), &H808080, BF
            Else
                Line (j * iWidth + giHorizBoardOffset, i * iHeight + giVertBoardOffset)-((j + 1) * iWidth + giHorizBoardOffset, (i + 1) * iHeight + giVertBoardOffset), &HFFFFFF, BF
            End If
        Next j
        color = Not color
    Next i
    Line (giHorizBoardOffset - 1, giVertBoardOffset - 1)-(giHorizBoardOffset + 32 * 8 + 1, 32 * 8 + giVertBoardOffset), , B
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then 'ctrl key
        PrintCoordinates False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then 'escape key
        giUnloadMode = vbFormCode
        Terminate Me
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then 'ctrl key
        PrintCoordinates True
    End If
End Sub


Private Sub Form_Load()
    Dim iBorder%, i%
    
    For i = 1 To 63
        Load imgSquare(i)
        imgSquare(i).Visible = True
    Next i
    
    giVertBoardOffset = picPlayers.Top + picPlayers.Height + 6
    giHorizBoardOffset = 15

    'set form width
    iBorder = Width - (ScaleWidth * Screen.TwipsPerPixelX)
    Width = (8 * 32 + (giHorizBoardOffset * 3) + lstMoves.Width) * Screen.TwipsPerPixelX + iBorder
   
    SetFormHeight False

    lblSubject.Left = (ScaleWidth - lblSubject.Width) / 2
    picPlayers.Left = (ScaleWidth - picPlayers.Width) / 2
    
    cmdSend.Top = ScaleHeight - cmdSend.Height - giHorizBoardOffset
    cmdUndo.Top = cmdSend.Top - (cmdUndo.Height + 4)
    lblMoves.Top = giVertBoardOffset
    lblMoves.Left = (2 * giHorizBoardOffset + 32 * 8)
    lstMoves.Top = lblMoves.Top + lblMoves.Height
    lstMoves.Left = (2 * giHorizBoardOffset + 32 * 8)
    lstMoves.Height = (cmdUndo.Top) - (lblMoves.Top + lblMoves.Height) - 5
    cmdSend.Width = lstMoves.Width
    cmdSend.Left = lstMoves.Left
    cmdUndo.Left = cmdSend.Left
    cmdUndo.Width = cmdSend.Width
    
    txtComments.Top = cmdSend.Top + cmdSend.Height + giHorizBoardOffset + 1
    lblComments.Top = txtComments.Top + giHorizBoardOffset
    lblComments.Left = giHorizBoardOffset
    txtComments.Left = lblComments.Left + lblComments.Width + giHorizBoardOffset / 2
    txtComments.Width = ScaleWidth - txtComments.Left - giHorizBoardOffset
    
    SetFormHeight True
    CenterFormOnScreen frmGame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    giUnloadMode = UnloadMode
    If Not gfMapifVbxUnload Then
        Terminate Me
        Cancel = True
    End If
End Sub

Private Sub imgSquare_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim fValidMove%, iMoveType%, iOpponentColor%, iSrcColor%
    
    imgSquare(giLastIndex).Drag vbEndDrag
    iSrcColor = Val(Left(gasSqr(giLastIndex), 1))
    
    If Source Is imgSquare(giLastIndex) And giLastIndex <> Index Then
        fValidMove = fValidChessMove(Index, iMoveType)
    
        If fValidMove = True And gasSqr(Index) <> iOpponentColor & kKingNotation Then
            If iMoveType = kStandard Then
                fValidMove = DoStandardMove(Index, Source)
                If fValidMove = False Then Exit Sub
            ElseIf iMoveType = kCastling Then
                DoCastling Index, Source
            ElseIf iMoveType = kEnPassant Then
                fValidMove = DoEnPassant(Index, Source)
                If fValidMove = False Then Exit Sub
            End If
            
            iOpponentColor = (1 / iSrcColor + 1 / iSrcColor)
            If fCheckForCheck(iOpponentColor) Then
                gfInCheck(iOpponentColor) = True
                If fCheckForCheckmate(iOpponentColor) Then
                    gsMoveNotation = gsMoveNotation & "+"
                    imgSquare(giKingIndex(iOpponentColor)) = imgKingChmate(iOpponentColor)
                    CheckmateAlert
                End If
                gsMoveNotation = gsMoveNotation & "+"
            Else
                ' does this code ever get called??
                If gfInCheck(iOpponentColor) = True Then 'avoid unnecessary redraws
                    gfInCheck(iOpponentColor) = False
                    imgSquare(giKingIndex(iOpponentColor)).Picture = imgKing(iOpponentColor).Picture
                End If
            End If
               
            gfMovedPiece = True
            
            lstMoves.AddItem gsMoveNotation
            lstMoves.Tag = "x" 'don't activate click event
            lstMoves.Selected(lstMoves.ListCount - 1) = True
            lstMoves.Tag = "" 'allow click event
            
            AdjustUndoButton
            cmdSend.Enabled = True
        Else
            Beep
        End If
    End If
End Sub

Private Sub imgSquare_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PrintCoordinates False
    Else
        If gfGameOver = False And lstMoves.ListCount = lstMoves.ListIndex + 1 And gfPlayed = False And Val(Left(gasSqr(Index), 1)) = giCurTurn And (cmdUndo.Caption = kSurrender Or (cmdUndo.Caption = kUndo And cmdUndo.Enabled = False)) Then
             If Right(gasSqr(Index), 1) = kKingNotation Then
                imgSquare(Index).DragIcon = imgKing(Left(gasSqr(Index), 1))
             Else
                imgSquare(Index).DragIcon = imgSquare(Index).Picture
             End If
             imgSquare(Index).Drag vbBeginDrag
             giLastIndex = Index
        End If
    End If
End Sub

Private Sub imgSquare_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub lblComments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates False
End Sub

Private Sub lblComments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub lblMoves_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates False
End Sub

Private Sub lblMoves_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub lblSubject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates False
End Sub

Private Sub lblSubject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub lstMoves_Click()
    Dim sUndoCaption$

    If lstMoves.Tag = "" Then
        'if clicked bottom list item
        If lstMoves.ListIndex + 1 = lstMoves.ListCount Then
            'SetFormHeight True
             'cmdUndo.Enabled = True
            If gfMovedPiece = True Then
                cmdSend.Enabled = True
                sUndoCaption = kUndo
            Else
                cmdSend.Enabled = False
                sUndoCaption = kSurrender
            End If
            If gfPlayed = False Then
                PrintPlayerNames giCurTurn
            Else
                If Int(lstMoves.ListIndex / 2) * 2 = lstMoves.ListIndex Then
                    PrintPlayerNames kBlue
                Else
                    PrintPlayerNames kRed
                End If
            End If
            '** the following lines might be able to go:
            If gfGameOver = True Then
                'cmdUndo.Enabled = False
                sUndoCaption = kSurrender
            End If
            If gsMoveNotation = kSurrenderNotation Then
                'cmdUndo.Enabled = True
                sUndoCaption = kUndo
            End If
        Else
            If Int(lstMoves.ListIndex / 2) * 2 = lstMoves.ListIndex Then
                PrintPlayerNames kBlue
            Else
                PrintPlayerNames kRed
            End If
            'SetFormHeight False
            cmdSend.Enabled = False
            'cmdUndo.Enabled = False
        End If
        If lstMoves.ListCount <> 0 Then
            ReplayMoves lstMoves.ListIndex + 1
            ResetKings
        End If
        FindComment lstMoves.ListIndex + 1
    End If
    'If sUndoCaption <> "" Then cmdUndo.Caption = sUndoCaption
    AdjustUndoButton
End Sub


Public Sub SetFormHeight(fShowComments%)
    Dim iBorder%, iCommentHeight

    If fShowComments = True Then
        iCommentHeight = Screen.TwipsPerPixelY * (giHorizBoardOffset + txtComments.Height)
    End If

    iBorder = Height - (ScaleHeight * Screen.TwipsPerPixelY)
    Height = iCommentHeight + (8 * 32 + (giVertBoardOffset + giHorizBoardOffset)) * Screen.TwipsPerPixelY + iBorder

    DrawBoard
End Sub

Private Sub lstMoves_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates False
End Sub

Private Sub lstMoves_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub picPlayers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates False
End Sub

Private Sub picPlayers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PrintCoordinates True
End Sub

Private Sub tmrCheck_Timer()
    Dim iColor%
    
    For iColor = kBlue To kRed
        If gfInCheck(iColor) And Not (lstMoves.ListIndex = lstMoves.ListCount - 1 And gfGameOver = True) Then
            If tmrCheck.Tag = "" Then
                imgSquare(giKingIndex(iColor)).Picture = imgKingCh(iColor)
                tmrCheck.Tag = "x"
            Else
                imgSquare(giKingIndex(iColor)).Picture = imgKing(iColor)
                tmrCheck.Tag = ""
            End If
        End If
    Next
End Sub

Private Sub txtComments_Change()
    txtComments.Tag = "x" 'been modified
End Sub




Public Sub FlipBoard()
    Dim i%
    
    'frmGame!imgSquare(0).Index = 15
    'For i = 0 To 31 '** 64/2 - 1
    '    frmGame(i).index =
    'Next i
End Sub
