VERSION 4.00
Begin VB.Form frmPawnExchange 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1605
   ClientLeft      =   2415
   ClientTop       =   7170
   ClientWidth     =   4695
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2010
   Left            =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   107
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   Top             =   6825
   Width           =   4815
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblExchange 
      Caption         =   "Click piece to exchange pawn for:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image imgPiece 
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmPawnExchange"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Const kNumPieces% = 4 'number of pieces to choose from
Dim giCurSelPiece%

Private Sub cmdok_click()
    Dim sNewTag$, iCurColor%, iIndex%
    
    iCurColor = Val(Left(gasSqr(giPawnIndex), 1))
    gasSqr(giPawnIndex) = iCurColor & imgPiece(giCurSelPiece).Tag
    DrawSquare giPawnIndex
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i%, iBorder%, iSrcColor%

    CenterFormOnScreen Me
    iBorder = (frmPawnExchange.ScaleWidth - kNumPieces * 32) / (kNumPieces + 1)
    For i = 1 To kNumPieces%
        Load imgPiece(i)
        imgPiece(i).Visible = True
    Next i
    For i = 0 To kNumPieces%
        imgPiece(i).Left = iBorder * (i + 1) + 32 * i
        imgPiece(i).Top = lblExchange.Top + lblExchange.Height + 4
    Next i
    
    iSrcColor = Val(Left(gasSqr(giPawnIndex), 1))
    
    imgPiece(0) = frmGame!imgRook(iSrcColor)
    imgPiece(0).Tag = kRookNotation
    imgPiece(1) = frmGame!imgQueen(iSrcColor)
    imgPiece(1).Tag = kQueenNotation
    imgPiece(2) = frmGame!imgKnight(iSrcColor)
    imgPiece(2).Tag = kKnightNotation
    imgPiece(3) = frmGame!imgBishop(iSrcColor)
    imgPiece(3).Tag = kBishopNotation

    giCurSelPiece = 1
    imgPiece(giCurSelPiece).BorderStyle = 1
End Sub

Private Sub imgPiece_Click(Index As Integer)
    Dim i%
    
    ' deselect all pieces
    For i = 0 To kNumPieces - 1
        imgPiece(i).BorderStyle = 0
    Next i
   
    ' select the clicked piece
    imgPiece(Index).BorderStyle = 1
    giCurSelPiece = Index
End Sub

Private Sub imgPiece_DblClick(Index As Integer)
    cmdok_click
End Sub

