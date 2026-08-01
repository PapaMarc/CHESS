Attribute VB_Name = "GameUtilProcs"
'** GAMEUTIL.BAS contains all code related to chess play

Option Explicit

'** vars: play related
Global gasSqr(0 To 63) As String    '** cur board data, each square is 2-char string
                                    '** 1st char: color, 2nd char: piece type
Global gfInCheck%(1 To 2) '** is either blue or red in check?
Global giKingIndex%(1 To 2) '** the index of blue/red king
Global giCurTurn% '** who's turn? blue? red?
Global gsSubject$ '** subject of current message
Global gasPlayerName$(1 To 2) '** blue/red player names
Global gasPlayerAddress$(1 To 2) '** blue/red player addresses
Global gsMoveData$ '** history of all moves played during game
Global giPawnIndex% '** for use w/ pawn exchange, gasSqr index of pawn
Global gsMoveNotation$ '** stores current move's notation
Global giMoveNum% '** number of move opened by user
Global gfPlayed% '** has gmLoadMsg been played?
Global gfMovedPiece% '** can the user undo?
Global gfByOpponent% '** is gmLoadMsg a move by the opponent?
Global gfRookMoved%(0 To 3) '** castling illegal because rook moved?
Global gfKingMoved%(1 To 2) '** castling illegal because king moved?
Global giEnPassant% '** index of pawn, if one just jumped two ranks
Global giOldEnPassant% '** backs-up giEnPassant, in case of undo
Global giVertBoardOffset% '** pixels to top of board
Global giHorizBoardOffset% '** pixel to left of board
Global gfGameOver% '** is the game over?
Global gfOpponentSurrender% '** did the opponent surrender (set by ReplayMoves)
Global giLastIndex% '** last square player MouseDown'ed in

'constants: move types
Global Const kStandard = 0
Global Const kCastling = 1
Global Const kEnPassant = 2

'constants: piece colors
Global Const kBlue = 1
Global Const kRed = 2

'constants: piece types
Global Const kPawn = 0
Global Const kKing = 1
Global Const kQueen = 2
Global Const kRook = 3
Global Const kKnight = 4
Global Const kBishop = 5
Global Const kEmptySquare = ""

Sub PrintPlayerNames(iCurTurn%)
    '** Print the names of the players (in color) in picPlayers of frmGame.
    '** Current player (iCurTurn) is underlined.
    '** Ex: Aaron Bell vs. Colonol Sanders
    
    Dim sName1$, sName2$

    sName1 = gasPlayerName(kRed)
    sName2 = gasPlayerName(kBlue)

    If frmGame!picPlayers.TextWidth(sName1) > (2 / 5) * frmGame!picPlayers.ScaleWidth Then
        sName1 = Left(sName1, 16) & "..."
    End If
    
    If frmGame!picPlayers.TextWidth(sName2) > (2 / 5) * frmGame!picPlayers.ScaleWidth Then
        sName2 = Left(sName2, 16) & "..."
    End If

    frmGame!picPlayers.Cls
    frmGame!picPlayers.CurrentY = 0
    frmGame!picPlayers.CurrentX = (frmGame!picPlayers.ScaleWidth - frmGame!picPlayers.TextWidth(sName1 & kVersus & sName2)) / 2
    frmGame!picPlayers.ForeColor = RGB(255, 0, 0) '** red
    
    If Val(iCurTurn) = kRed Then frmGame!picPlayers.FontUnderline = True
    frmGame!picPlayers.Print sName1;
    
    frmGame!picPlayers.ForeColor = RGB(0, 0, 0) '** black
    frmGame!picPlayers.FontUnderline = False
    frmGame!picPlayers.Print kVersus;

    If Val(iCurTurn) = kBlue Then frmGame!picPlayers.FontUnderline = True
    frmGame!picPlayers.ForeColor = RGB(0, 0, 255) '** blue
    frmGame!picPlayers.Print sName2
    frmGame!picPlayers.FontUnderline = False
End Sub




Sub CalcCoord(iIndex As Integer, iX As Integer, iY As Integer)
    '** the gasSqr array has 64 items (0-63), each one representing
    '** a square of the board. CalcCoord converts this index number to
    '** x, y coordinates.
    '** Top-right corner = 0,0  Top-left corner = 7,7
    iY = Int(iIndex / 8)
    iX = iIndex - iY * 8
End Sub

Function iCalcIndex(iX%, iY%)
    '** convert x, y coordinates to a gasSqr index
    iCalcIndex = (iY * 8) + iX
End Function

Sub CenterFormOnScreen(frmCenter As Form)
    '** center frmCenter on the screen
    frmCenter.Move (Screen.Width - frmCenter.Width) / 2, (Screen.Height - frmCenter.Height) / 2
End Sub

Sub DrawSquare(iIndex%)
    '** set the icon for a board square
    '**     index = board square (0 - 63)
    Dim iColor%, sType$

    sType$ = Right(gasSqr(iIndex), 1)
    iColor% = Val(Left(gasSqr(iIndex), 1))
    Select Case sType
        Case kRookNotation
            frmGame!imgSquare(iIndex) = frmGame!imgRook(iColor%)
        Case kQueenNotation
            frmGame!imgSquare(iIndex) = frmGame!imgQueen(iColor%)
        Case kKnightNotation
            frmGame!imgSquare(iIndex) = frmGame!imgKnight(iColor%)
        Case kPawnNotation
            frmGame!imgSquare(iIndex) = frmGame!imgPawn(iColor%)
        Case kKingNotation
            frmGame!imgSquare(iIndex) = frmGame!imgKing(iColor%)
            giKingIndex(iColor) = iIndex
        Case kBishopNotation
            frmGame!imgSquare(iIndex) = frmGame!imgBishop(iColor%)
        End Select
End Sub



Public Sub ReplayMoves(iNumMoves%)
    '** replays the history of the board up to certain move (iNumMoves)
    '** uses data from gsMoveData (comma/semi-delimited string)
    '**  initializes board using saPieces (array like gasSqr)
    '**  parses gsMoveData, Ex: 34,23;!P,7,1,Q
    '**  "34,23" would replace data in saPieces(23) with data from
    '**  saPieces(34), then set saPieces(34) = kEmptySquare
    '**  "!P,7,1,1Q" ! means special move, P means pawn exchange
    '**  set saPieces(7) to kEmptySquare and saPieces(1) to 1Q
    '**  !C means castling, !E means EnPassant, S! means surrender

    Dim iLoc%, sDest$, i%, sMove$, iLocMove%
    Dim sMovingPiece$, saPieces$(0 To 63)
    Dim sBoardData$, iDestSqr%, iKingDead%
    Dim fCheckmate%
    
    If Right(frmGame.lstMoves, 2) = "++" Then fCheckmate = True

    '** start board initialization
    saPieces(0) = kRed & kRookNotation
    saPieces(1) = kRed & kKnightNotation
    saPieces(2) = kRed & kBishopNotation
    saPieces(3) = kRed & kQueenNotation
    saPieces(4) = kRed & kKingNotation
    saPieces(5) = kRed & kBishopNotation
    saPieces(6) = kRed & kKnightNotation
    saPieces(7) = kRed & kRookNotation
    For i = 8 To 15
        saPieces(i) = kRed & kPawnNotation
    Next i
    For i = 48 To 55
        saPieces(i) = kBlue & kPawnNotation
    Next
    saPieces(56) = kBlue & kRookNotation
    saPieces(57) = kBlue & kKnightNotation
    saPieces(58) = kBlue & kBishopNotation
    saPieces(59) = kBlue & kQueenNotation
    saPieces(60) = kBlue & kKingNotation
    saPieces(61) = kBlue & kBishopNotation
    saPieces(62) = kBlue & kKnightNotation
    saPieces(63) = kBlue & kRookNotation
    '** end board initialization

    iLoc% = 1
    i = 0
    While i < iNumMoves
        ParseSemicolonString gsMoveData, sMove, iLoc
        sMove = sMove & ","
        iLocMove = 1
        ParseCommaString sMove, sDest, iLocMove
        If sDest <> "!" Then
            sMovingPiece = saPieces(CInt(sDest))
            saPieces(CInt(sDest)) = kEmptySquare
            ParseCommaString sMove, sDest, iLocMove
            saPieces(CInt(sDest)) = sMovingPiece
        Else
            ParseCommaString sMove, sDest, iLocMove
            If sDest = "C" Then
                Dim iOrigKing%, iOrigRook%
            
                ParseCommaString sMove, sDest, iLocMove
                iOrigKing = CInt(sDest)
                ParseCommaString sMove, sDest, iLocMove
                iOrigRook = CInt(sDest)
                ParseCommaString sMove, sDest, iLocMove
                saPieces(CInt(sDest)) = saPieces(iOrigKing)
                saPieces(iOrigKing) = ""
                ParseCommaString sMove, sDest, iLocMove
                saPieces(CInt(sDest)) = saPieces(iOrigRook)
                saPieces(iOrigRook) = ""
            ElseIf sDest = "P" Then
                ParseCommaString sMove, sDest, iLocMove
                sMovingPiece = saPieces(CInt(sDest))
                saPieces(CInt(sDest)) = kEmptySquare
                ParseCommaString sMove, sDest, iLocMove
                iDestSqr% = CInt(sDest)
                ParseCommaString sMove, sDest, iLocMove
                saPieces(CInt(iDestSqr)) = Left(sMovingPiece, 1) & sDest
            ElseIf sDest = "E" Then
                ParseCommaString sMove, sDest, iLocMove
                sMovingPiece = saPieces(CInt(sDest))
                saPieces(CInt(sDest)) = kEmptySquare
                ParseCommaString sMove, sDest, iLocMove
                saPieces(sDest) = sMovingPiece
                ParseCommaString sMove, sDest, iLocMove
                saPieces(CInt(sDest)) = kEmptySquare
            ElseIf sDest = "S" Then
                ParseCommaString sMove, sDest, iLocMove
                iKingDead = Val(sDest)
                gfGameOver = True
                If iKingDead <> giCurTurn Then
                    gfOpponentSurrender = True
                End If
            End If

        End If
        i = i + 1
    Wend
    
    For i = 0 To 63
        sBoardData = sBoardData & saPieces(i) & ","
    Next i
    UnstuffBoardData sBoardData
    
    If gfGameOver Or fCheckmate Then
        If fCheckmate Then
            If Int(frmGame.lstMoves.ListCount / 2) * 2 = frmGame.lstMoves.ListCount Then
                iKingDead = kBlue
            Else
                iKingDead = kRed
            End If
        End If
        If iKingDead Then
            frmGame!imgSquare(giKingIndex(iKingDead)) = frmGame!imgKingChmate(iKingDead)
        Else
            frmGame!imgSquare(giKingIndex(kRed)) = frmGame!imgKing(kRed)
            frmGame!imgSquare(giKingIndex(kBlue)) = frmGame!imgKing(kBlue)
        End If
    End If
End Sub

Public Function fCheckForCheck(iKingColor%)
    '** checks if the iKingColor king is in check from opponent
    '** king is assumed to be in gasSqr(giKingIndex(iKingColor))
    '** loops through each opponent piece checking if can attack king
    '** fCheckForCheck = true if in check
    
    Dim i%, iKingIndex%, sPieceType$, iOffenseColor%
    
    If iKingColor = kRed Then
        iOffenseColor = kBlue
    Else
        iOffenseColor = kRed
    End If
    
    iKingIndex = giKingIndex(iKingColor)

    fCheckForCheck = False
    For i = 0 To 63
        If Val(Left$(gasSqr(i), 1)) = iOffenseColor Then
            sPieceType = Right(gasSqr(i), 1)
            Select Case sPieceType
                Case kRookNotation
                    If fValidRookMove(i, iKingIndex, iOffenseColor, iKingColor) Then fCheckForCheck = True
                Case kQueenNotation
                    If fValidQueenMove(i, iKingIndex, iOffenseColor, iKingColor) Then fCheckForCheck = True
                Case kKnightNotation
                    If fValidKnightMove(i, iKingIndex, iOffenseColor, iKingColor) Then fCheckForCheck = True
                Case kPawnNotation
                    If fValidPawnMove(i, iKingIndex, iOffenseColor, iKingColor) Then fCheckForCheck = True
                Case kBishopNotation
                    If fValidBishopMove(i, iKingIndex, iOffenseColor, iKingColor) Then fCheckForCheck = True
                Case kKingNotation
                    If fValidKingMove(i, iKingIndex, iOffenseColor, iKingColor) Then fCheckForCheck = True
            End Select
        End If
        If fCheckForCheck = True Then Exit For
    Next i
End Function
Public Function fValidBishopMove(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** Can the bishop at iSrcIndex move to iDestIndex?
    Dim iSrcX%, iSrcY%, iDestX%, iDestY%

    fValidBishopMove = False

    CalcCoord iSrcIndex, iSrcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY

    If Abs(iSrcX - iDestX) = Abs(iDestY - iSrcY) Then
        If CheckPathSect(iSrcX, iSrcY, iDestX, iDestY, iSrcColor, iDestColor) = False Then
            fValidBishopMove = True
        End If
    End If
End Function

Public Function fValidKingMove(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** Can the King at iSrcIndex move to iDestIndex?
    Dim iSrcX%, iSrcY%, iDestX%, iDestY%
    
    CalcCoord iSrcIndex, iSrcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY

    If (iSrcColor <> iDestColor) And Abs(iSrcX - iDestX) < 2 And Abs(iSrcY - iDestY) < 2 Then
        fValidKingMove = True
    End If
    
End Function

Public Function fValidKnightMove(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** Can the knight at iSrcIndex move to iDestIndex?
    Dim iSrcX%, iSrcY%, iDestX%, iDestY%

    CalcCoord iSrcIndex, iSrcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY

    If (iSrcColor <> iDestColor) And ((Abs(iSrcX - iDestX) = 1 And Abs(iSrcY - iDestY) = 2) Or (Abs(iSrcX - iDestX) = 2 And Abs(iSrcY - iDestY) = 1)) Then
        fValidKnightMove = True
    End If
End Function

Public Function fValidPawnMove(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** Can the pawn at iSrcIndex move to iDestIndex?
    Dim iSrcX%, iSrcY%, iDestX%, iDestY%, iYOffset%
    
    fValidPawnMove = False
    
    If Abs(iSrcIndex - iDestIndex) > 16 Then Exit Function 'little shortcut

    CalcCoord iSrcIndex, iSrcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY
    If iSrcColor = kRed Then
        iYOffset = 1
    Else
        iYOffset = -1
    End If

    If iDestColor = 0 Then 'empty square
        If iSrcX = iDestX And iDestY - iSrcY = iYOffset And iDestColor = 0 Then
            fValidPawnMove = True
        ElseIf iSrcColor = kRed And iSrcX = iDestX And iSrcY = 1 And iDestY = 3 And gasSqr(iSrcIndex + 8) = kEmptySquare Then
            fValidPawnMove = True
        ElseIf iSrcColor = kBlue And iSrcX = iDestX And iSrcY = 6 And iDestY = 4 And gasSqr(iSrcIndex - 8) = kEmptySquare Then
            fValidPawnMove = True
        End If
    ElseIf iSrcColor <> iDestColor Then 'opponent square
        If Abs(iSrcX - iDestX) = 1 And iDestY - iSrcY = iYOffset Then
            fValidPawnMove = True
        End If
    End If
End Function
Public Function fValidChessMove(Index%, iMoveType%)
    '** can the the piece at giLastIndex legally move to Index?
    
    Dim iSrcColor%, iDestColor%, sPieceType$, i%

    iSrcColor = Val(Left(gasSqr(giLastIndex), 1))
    iDestColor = Val(Left(gasSqr(Index), 1))
    
    iMoveType = kStandard
    sPieceType = Right(gasSqr(giLastIndex), 1)
    If sPieceType = kPawnNotation Then
        fValidChessMove = fValidEnPassant(giLastIndex, Index, iSrcColor, iDestColor)
        If fValidChessMove = True Then
            iMoveType = kEnPassant
        Else
            fValidChessMove = fValidPawnMove(giLastIndex, Index, iSrcColor, iDestColor)
        End If
    ElseIf sPieceType = kBishopNotation Then
        fValidChessMove = fValidBishopMove(giLastIndex, Index, iSrcColor, iDestColor)
    ElseIf sPieceType = kRookNotation Then
        If (gasSqr(Index) = iSrcColor & kKingNotation) Then
            fValidChessMove = fValidCastling(Index, giLastIndex, iSrcColor)
            iMoveType = kCastling
        Else
            fValidChessMove = fValidRookMove(giLastIndex, Index, iSrcColor, iDestColor)
        End If
    ElseIf sPieceType = kKnightNotation Then
        fValidChessMove = fValidKnightMove(giLastIndex, Index, iSrcColor, iDestColor)
    ElseIf sPieceType = kKingNotation Then
        iMoveType = kCastling
        If (gasSqr(Index) = iSrcColor & kRookNotation) Then
            i = Index
            fValidChessMove = fValidCastling(giLastIndex, i, iSrcColor)
        ElseIf (iSrcColor = kRed And Index = 2) Then
            i = 0
            fValidChessMove = fValidCastling(giLastIndex, i, iSrcColor)
        ElseIf (iSrcColor = kRed And Index = 6) Then
            i = 7
            fValidChessMove = fValidCastling(giLastIndex, i, iSrcColor)
        ElseIf (iSrcColor = kBlue And Index = 58) Then
            i = 56
            fValidChessMove = fValidCastling(giLastIndex, i, iSrcColor)
        ElseIf (iSrcColor = kBlue And Index = 62) Then
            i = 63
            fValidChessMove = fValidCastling(giLastIndex, i, iSrcColor)
        End If
        If fValidChessMove = False Then
            iMoveType = kStandard
            fValidChessMove = fValidKingMove(giLastIndex, Index, iSrcColor, iDestColor)
        Else
            Index = i
        End If
    ElseIf sPieceType = kQueenNotation Then
        fValidChessMove = fValidQueenMove(giLastIndex, Index, iSrcColor, iDestColor)
    End If
End Function

Public Function fValidQueenMove(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** Can the queen at iSrcIndex move to iDestIndex?
    Dim iSrcX%, iSrcY%, iDestX%, iDestY%, i%

    fValidQueenMove = False

    CalcCoord iSrcIndex, iSrcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY

    If Abs(iSrcX - iDestX) = Abs(iDestY - iSrcY) Or iDestY = iSrcY Or iDestX = iSrcX Then
        If CheckPathSect(iSrcX, iSrcY, iDestX, iDestY, iSrcColor, iDestColor) = False Then
            fValidQueenMove = True
        End If
    End If
End Function

Public Function fValidRookMove(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** Can the rook at iSrcIndex move to iDestIndex?
    Dim iSrcX%, iSrcY%, iDestX%, iDestY%, i%

    CalcCoord iSrcIndex, iSrcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY

    If iSrcX = iDestX Or iDestY = iSrcY Then
        If CheckPathSect(iSrcX, iSrcY, iDestX, iDestY, iSrcColor, iDestColor) = False Then
            fValidRookMove = True
        End If
    End If
End Function

Public Function CheckPathSect(iSrcX%, iSrcY%, iDestX%, iDestY%, ByVal iSrcColor%, ByVal iDestColor%)
    '** checks for a piece placed between a queens/bishops/rooks source and dest
    '** returns true if piece is in the way
    Dim iDeltaX%, iDeltaY%, iCurX%, iCurY%

    If iDestX > iSrcX Then
        iDeltaX = 1
    ElseIf iDestX = iSrcX Then
        iDeltaX = 0
    Else
        iDeltaX = -1
    End If

    If iDestY > iSrcY Then
        iDeltaY = 1
    ElseIf iDestY = iSrcY Then
        iDeltaY = 0
    Else
        iDeltaY = -1
    End If
    
    iCurX = iSrcX
    iCurY = iSrcY
    While iCurX <> iDestX Or iCurY <> iDestY
        iCurX = iCurX + iDeltaX
        iCurY = iCurY + iDeltaY
        If gasSqr(iCalcIndex(iCurX, iCurY)) <> kEmptySquare Then
            'if it's not the last square, where the colors are different
            If Not (iCurX = iDestX And iCurY = iDestY And iSrcColor <> iDestColor) Then
                CheckPathSect = True
                Exit Function
            End If
        End If
    Wend
End Function

Public Sub ResetKings()
    '** checks blue/red king to see if he is in check
    '** if so, gfInCheck is set = true
    '** else, gfInCheck is set = false and kings icon is set to normal
    
    Dim iColor%
        
    For iColor = kBlue To kRed
        If fCheckForCheck(iColor) Then
            gfInCheck(iColor) = True
        Else
            If gfInCheck(iColor) = True Then
                gfInCheck(iColor) = False
                frmGame!imgSquare(giKingIndex(iColor)).Picture = frmGame!imgKing(iColor).Picture
            End If
        End If
    Next iColor
End Sub

Public Function fValidCastling(ByVal iKingIndex%, ByVal iRookIndex%, ByVal iColor%)
    '** can rook at iRookIndex and king and iKingIndex of color iColor
    '** legally perfom castling?
    Dim iKingX%, iRookX%, iY%, i%
    Dim iOffset%, ioKingIndex%

    CalcCoord iKingIndex, iKingX, iY
    CalcCoord iRookIndex, iRookX, iY
    
    fValidCastling = False
    If gfKingMoved(iColor) = False And gfInCheck(iColor) = False And gasSqr(iRookIndex) = iColor & kRookNotation Then
        If (iRookIndex = 0 And gfRookMoved(0) = False) Or (iRookIndex = 7 And gfRookMoved(1) = False) Or (iRookIndex = 56 And gfRookMoved(2) = False) Or (iRookIndex = 63 And gfRookMoved(3) = False) Then
        'ok, none have of the pieces have illegally moved
            If CheckPathSect(iKingX, iY, iRookX, iY, 0, 1) = False Then
                'nothing is in the way
                ioKingIndex = iKingIndex
                If iKingIndex > iRookIndex Then
                    iOffset = -1
                Else
                    iOffset = 1
                End If
                fValidCastling = True
                For i = 1 To 2
                    gasSqr(iKingIndex) = kEmptySquare
                    iKingIndex = iKingIndex + iOffset
                    gasSqr(iKingIndex) = iColor & kKingNotation
                    giKingIndex(iColor) = iKingIndex
                    If fCheckForCheck(iColor) = True Then
                        fValidCastling = False
                    End If
                Next i
                giKingIndex(iColor) = ioKingIndex
                gasSqr(ioKingIndex + iOffset) = kEmptySquare
                gasSqr(ioKingIndex + 2 * iOffset) = kEmptySquare
                gasSqr(ioKingIndex) = iColor & kKingNotation
            End If
        End If
    End If
End Function

Public Function fValidEnPassant(iSrcIndex%, iDestIndex%, ByVal iSrcColor%, ByVal iDestColor%)
    '** can pawn at iSrcIndex legally move to iDestIndex?
    Dim srcX%, iSrcY%, iDestX%, iDestY%

    If giEnPassant <> iDestIndex Or iDestColor <> 0 Then Exit Function
    
    CalcCoord iSrcIndex, srcX, iSrcY
    CalcCoord iDestIndex, iDestX, iDestY
    

    If iSrcColor = kRed Then
        If Abs(srcX - iDestX) = 1 And iSrcY - iDestY = -1 Then fValidEnPassant = True
    Else
        If Abs(srcX - iDestX) = 1 And iSrcY - iDestY = 1 Then fValidEnPassant = True
    End If
End Function

Public Sub PrintCoordinates(fErase%)
    '** if fErase = false
    '**   prints board coordinates along left and bottom sides of board
    '** if fErase = true
    '**   erases board coordinates along left and bottom sides of board
    
    Dim i%, fReverse%
    
    If (giCurTurn = kRed And gfByOpponent = True) Or (giCurTurn = kBlue And gfByOpponent = False) Then
        fReverse = True
    Else
        fReverse = False
    End If
    
    frmGame.AutoRedraw = False
    If fErase = False Then
        frmGame.CurrentX = -2
        
        frmGame.CurrentY = giVertBoardOffset + 10
        For i = 8 To 1 Step -1
            If fReverse Then
                frmGame.Print 9 - i;
            Else
                frmGame.Print i;
            End If
            frmGame.CurrentX = -2
            frmGame.CurrentY = frmGame.CurrentY + 32
        Next i
        frmGame.CurrentY = giVertBoardOffset + 8 * 32
        frmGame.CurrentX = giHorizBoardOffset + 12
        For i = 1 To 8
            If fReverse Then
                frmGame.Print Chr((9 - i) + 96);
            Else
                frmGame.Print Chr(i + 96);
            End If
            frmGame.CurrentX = giHorizBoardOffset + (i * 32) + 12
        Next i
    Else
        frmGame.Line (0, giVertBoardOffset)-(giHorizBoardOffset - 2, giVertBoardOffset + (8 * 32)), frmGame.BackColor, BF
        frmGame.Line (0, giVertBoardOffset + (8 * 32) + 1)-(giHorizBoardOffset + (8 * 32), frmGame!txtComments.Top - 2), frmGame.BackColor, BF
    End If
    frmGame.AutoRedraw = True
End Sub

Public Function fCheckForCheckmate%(iKingColor%)
     '** checks if the iKingColor king is in checkmate from opponent

    Dim i%, iKingIndex%, sPieceType$, iOffenseColor%, fStillInCheck%
    Dim sTmp$, iaKingDelta%(0 To 7), iLoop%, fValidMove%, iAttacker%, iDefender%
    Dim iDelta%, iBlock%, iDest%, iStep%
    Dim fEnPassant%, iEPDest%, iEPOpponent%

    fCheckForCheckmate = True
    
    If iKingColor = kRed Then
        iOffenseColor = kBlue
    Else
        iOffenseColor = kRed
    End If
    
    iKingIndex = giKingIndex(iKingColor)
    iaKingDelta(0) = -7: iaKingDelta(1) = -8: iaKingDelta(2) = -9
    iaKingDelta(3) = -1: iaKingDelta(4) = 1
    iaKingDelta(5) = 7: iaKingDelta(6) = 8: iaKingDelta(7) = 9

    For iLoop = 0 To 7
        i = iaKingDelta(iLoop)
        If iKingIndex + i >= 0 And iKingIndex + i <= 63 Then
            If fValidKingMove(iKingIndex, iKingIndex + i, iKingColor, Val(Left(gasSqr(iKingIndex + i), 1))) Then
                sTmp = gasSqr(iKingIndex + i)
                gasSqr(iKingIndex + i) = gasSqr(iKingIndex)
                gasSqr(iKingIndex) = ""
                giKingIndex(iKingColor) = iKingIndex + i
                If fCheckForCheck(iKingColor) = False Then
                    fCheckForCheckmate = False
                    giKingIndex(iKingColor) = iKingIndex
                    gasSqr(iKingIndex) = gasSqr(iKingIndex + i)
                    gasSqr(iKingIndex + i) = sTmp
                    Exit Function
                End If
                giKingIndex(iKingColor) = iKingIndex
                gasSqr(iKingIndex) = gasSqr(iKingIndex + i)
                gasSqr(iKingIndex + i) = sTmp
            End If
        End If
    Next iLoop
   
    For iAttacker = 0 To 63
           If Val(Left(gasSqr(iAttacker), 1)) = iOffenseColor Then
                fValidMove = False
                sPieceType = Right(gasSqr(iAttacker), 1)
                Select Case sPieceType
                    Case kRookNotation
                        If fValidRookMove(iAttacker, iKingIndex, iOffenseColor, iKingColor) Then fValidMove = True
                    Case kQueenNotation
                        If fValidQueenMove(iAttacker, iKingIndex, iOffenseColor, iKingColor) Then fValidMove = True
                    Case kKnightNotation
                        If fValidKnightMove(iAttacker, iKingIndex, iOffenseColor, iKingColor) Then fValidMove = True
                    Case kPawnNotation
                        If fValidPawnMove(iAttacker, iKingIndex, iOffenseColor, iKingColor) Then fValidMove = True
                    Case kBishopNotation
                        If fValidBishopMove(iAttacker, iKingIndex, iOffenseColor, iKingColor) Then fValidMove = True
                End Select
                
                If fValidMove = True Then
                    If sPieceType = kQueenNotation Then
                        Dim h1%, h2%, v1%, v2%
                        CalcCoord iAttacker, h1, v1
                        CalcCoord iKingIndex, h2, v2
                        If h1 <> h2 And v1 <> v2 Then
                            sPieceType = kBishopNotation
                        Else
                            sPieceType = kRookNotation
                        End If
                    End If

                    If sPieceType = kRookNotation Then
                        iDelta = iAttacker - iKingIndex
                        If iDelta >= 8 Then 'rook under king
                            iDest = iKingIndex + 8
                            iStep = -8
                        ElseIf iDelta > 0 Then ' rook to right of king
                            iDest = iKingIndex + 1
                            iStep = -1
                        ElseIf iDelta >= -8 Then 'rook to left of king
                            iDest = iKingIndex - 1
                            iStep = 1
                        Else 'king under rook
                            iDest = iKingIndex - 8
                            iStep = 8
                        End If
                    ElseIf sPieceType = kBishopNotation Then
                        iDelta = iAttacker - iKingIndex
                        If Int(iDelta / 9) * 9 = iDelta Then
                            If iDelta > 0 Then
                                iDest = iKingIndex + 9
                                iStep = -9
                            Else
                                iDest = iKingIndex - 9
                                iStep = 9
                            End If
                        Else
                            If iDelta > 0 Then
                                iDest = iKingIndex - 7
                                iStep = -7
                            Else
                                iDest = iKingIndex + 7
                                iStep = 7
                            End If
                        End If
                    ElseIf sPieceType = kPawnNotation Or sPieceType = kKnightNotation Then
                        iDest = iAttacker
                        iStep = 1
                    End If
                        
                    For iBlock = iAttacker To iDest Step iStep
                        If iBlock >= 0 And iBlock <= 63 Then
                            For iDefender = 0 To 63
                                If Val(Left(gasSqr(iDefender), 1)) = iKingColor Then
                                    fValidMove = False
                                    fEnPassant = False
                                    sPieceType = Right(gasSqr(iDefender), 1)
                                    Select Case sPieceType
                                        Case kRookNotation
                                            If fValidRookMove(iDefender, iBlock, iKingColor, Val(Left(gasSqr(iBlock), 1))) Then fValidMove = True
                                        Case kQueenNotation
                                            If fValidQueenMove(iDefender, iBlock, iKingColor, Val(Left(gasSqr(iBlock), 1))) Then fValidMove = True
                                        Case kKnightNotation
                                            If fValidKnightMove(iDefender, iBlock, iKingColor, Val(Left(gasSqr(iBlock), 1))) Then fValidMove = True
                                        Case kPawnNotation
                                            If fValidPawnMove(iDefender, iBlock, iKingColor, Val(Left(gasSqr(iBlock), 1))) Then
                                                fValidMove = True
                                            Else
                                                If iBlock = iAttacker Then
                                                    If iKingColor = kBlue Then
                                                        iEPDest = iBlock - 8
                                                    Else
                                                        iEPDest = iBlock + 8
                                                    End If
                                                    iEPOpponent = iBlock
                                                Else
                                                    If iKingColor = kRed Then
                                                        iEPOpponent = iBlock - 8
                                                    Else
                                                        iEPOpponent = iBlock + 8
                                                    End If
                                                    iEPDest = iBlock
                                                End If
                                                If iEPDest >= 0 And iEPDest <= 63 Then
                                                    If fValidEnPassant(iDefender, iEPDest, iKingColor, Val(Left(gasSqr(iEPDest), 1))) Then
                                                        fEnPassant = True
                                                        fValidMove = True
                                                    End If
                                                End If
                                            End If
                                        Case kBishopNotation
                                            If fValidBishopMove(iDefender, iBlock, iKingColor, Val(Left(gasSqr(iBlock), 1))) Then fValidMove = True
                                    End Select
                                    If fValidMove = True Then
                                        '** set up board to after move
                                        If fEnPassant Then
                                            sTmp = gasSqr(iEPOpponent)
                                            gasSqr(iEPDest) = gasSqr(iDefender)
                                            gasSqr(iDefender) = kEmptySquare
                                            gasSqr(iEPOpponent) = kEmptySquare
                                        Else
                                            sTmp = gasSqr(iBlock)
                                            gasSqr(iBlock) = gasSqr(iDefender)
                                            gasSqr(iDefender) = kEmptySquare
                                        End If
                                        
                                        '** are we still in check?
                                        fStillInCheck = False
                                        fStillInCheck = fCheckForCheck(iKingColor)
                                        
                                        '** reset board
                                        If fEnPassant Then
                                            gasSqr(iDefender) = gasSqr(iEPDest)
                                            gasSqr(iEPDest) = kEmptySquare
                                            gasSqr(iEPOpponent) = sTmp
                                        Else
                                            gasSqr(iDefender) = gasSqr(iBlock)
                                            gasSqr(iBlock) = sTmp
                                        End If
                                        
                                        '** if still in check, then we have a checkmate
                                        If fStillInCheck = False Then
                                            fCheckForCheckmate = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Next iDefender
                        End If
                    Next iBlock
                End If
            End If
        Next iAttacker
End Function

Public Function sSpecificNotation(ByVal sPieceInfo$, ByVal iSrcIndex%, ByVal iDestIndex%)
    '** the notation for a standard chess move can be ambigous
    '** ex: two knights of the same color may be able to move to the same
    '** square, if one moves there, the notation must specify which knight
    '** moved, either by rank or file or both if necessary.
    '** piece info is in the format color-piece type (2K, 1Q, 2P, etc..)
    '** iSrcIndex is the pieces origin, iDestIndex is where the piece moved to
    Dim sDestInfo$, i%, fSpecifyRank%, fSpecifyFile%, iColor%, fValid%
    Dim iSrcFile%, iSrcRank%, iAmbigFile%, iAmbigRank%, fClearOther%

    sDestInfo = gasSqr(iDestIndex)
    iColor = Left(sDestInfo, 1)
    gasSqr(iDestIndex) = kEmptySquare
    gasSqr(iSrcIndex) = sPieceInfo
    CalcCoord iSrcIndex, iSrcFile, iSrcRank

    For i = 0 To 63
        fValid = False
        If gasSqr(i) = sPieceInfo And i <> iSrcIndex Then
            Select Case Right(sPieceInfo, 1)
                Case kRookNotation
                    If fValidRookMove(i, iDestIndex, iColor, 0) Then fValid = True
                Case kQueenNotation
                    If fValidQueenMove(i, iDestIndex, iColor, 0) Then fValid = True
                Case kKnightNotation
                    If fValidKnightMove(i, iDestIndex, iColor, 0) Then fValid = True
                Case kPawnNotation
                    If fValidPawnMove(i, iDestIndex, iColor, 0) Then fValid = True
                Case kBishopNotation
                    If fValidBishopMove(i, iDestIndex, iColor, 0) Then fValid = True
            End Select
        End If

        If fValid = True Then
            CalcCoord i, iAmbigFile, iAmbigRank
            If iAmbigRank = iSrcRank Then
                fSpecifyFile = True
                If fClearOther Then fSpecifyRank = False
            ElseIf iAmbigFile = iSrcFile Then
                fSpecifyRank = True
                If fClearOther Then fSpecifyFile = False
            Else
                If fSpecifyRank = False And fSpecifyFile = False Then
                    fClearOther = True
                    fSpecifyFile = True
                End If
            End If
        End If
    Next i
    gasSqr(iDestIndex) = sDestInfo
    gasSqr(iSrcIndex) = kEmptySquare
    If fSpecifyFile Then sSpecificNotation = Chr(iSrcFile + 97)
    If fSpecifyRank Then sSpecificNotation = sSpecificNotation & (7 - iSrcRank + 1)
End Function

Public Sub CheckmateAlert()
    '** Alerts the user of a checkmate
    '** Occurs after executing a game-ending move or upon
    '**     opening a chess game were you user was defeated
    Dim sGameOver$

    gfGameOver = True
    If gfInCheck(kBlue) Then
        sGameOver = kCheckmate & gasPlayerName(kRed) & kDefeats & gasPlayerName(kBlue) & "..."
    Else
        sGameOver = kCheckmate & gasPlayerName(kBlue) & kDefeats & gasPlayerName(kRed) & "..."
    End If
    MsgBox sGameOver, vbOKOnly, kGameOver
End Sub

Public Sub SurrenderAlert()
    '** similar to checkmate alert
    '** alerts user upon opening a chess game in which the opponent surrendered.
    Dim sGameOver$
    
    gfGameOver = True
    If (gfOpponentSurrender And giCurTurn = kBlue) Or (Not gfOpponentSurrender And giCurTurn = kRed) Then
        sGameOver = gasPlayerName(kRed) & kSurrendered
    Else
        sGameOver = gasPlayerName(kBlue) & kSurrendered
    End If
    MsgBox sGameOver, vbOKOnly, kGameOver
End Sub

Public Sub AdjustUndoButton()
    '** the Undo button has tree states:
    '** - "disabled undo" on the first move of the game
    '** - "enabled undo" after move made or decision to surrender
    '** - "enabled surrender" before a move is made
    '** AdjustUndoButton sets the undo button to the correct state
    If frmGame!lstMoves.ListCount = 0 Or (frmGame!lstMoves.ListCount = 1 And giCurTurn = kRed) Or frmGame!lstMoves.ListCount - 1 <> frmGame!lstMoves.ListIndex Or gfGameOver Or gfOpponentSurrender Or gfPlayed Then
        frmGame!cmdUndo.Caption = kUndo
        '** if user just surrendered, allow them to undo
        If gfMovedPiece And gfGameOver And gfOpponentSurrender = False And gfPlayed = False Then
            frmGame!cmdUndo.Enabled = True
        Else
            frmGame!cmdUndo.Enabled = False
        End If
    ElseIf gfMovedPiece Or gsMoveNotation = kSurrenderNotation Then
        frmGame!cmdUndo.Caption = kUndo
        frmGame!cmdUndo.Enabled = True
    Else
        frmGame!cmdUndo.Caption = kSurrender
        frmGame!cmdUndo.Enabled = True
    End If
End Sub

Public Function DoStandardMove%(Index%, Source As Object)
    '** perform standard move
    
    Dim iFile%, iRank%, sPieceType$
    Dim sDestTag$, iSrcColor%, iDestColor%, iOpponentColor%

    iSrcColor = Val(Left(gasSqr(giLastIndex), 1))
    iOpponentColor = (1 / iSrcColor + 1 / iSrcColor)
    iDestColor = Val(Left(gasSqr(Index), 1))
    sPieceType = Right(gasSqr(giLastIndex), 1)
    
    sDestTag = gasSqr(Index)

    If sPieceType = kKingNotation Then giKingIndex(iSrcColor) = Index

    gasSqr(Index) = gasSqr(giLastIndex)
    gasSqr(giLastIndex) = kEmptySquare

    'if we just moved ourselves into check, get out function!
    If fCheckForCheck(iSrcColor) Then
        'quick! revert while you can!
        gasSqr(giLastIndex) = gasSqr(Index)
        gasSqr(Index) = sDestTag$
         
        If sPieceType = kKingNotation Then giKingIndex(iSrcColor) = giLastIndex
        
        Beep
        Exit Function
    Else
        If gfInCheck(iSrcColor) = True Then
            gfInCheck(iSrcColor) = False
            If sPieceType = kKingNotation Then
                frmGame!imgSquare(giLastIndex).Picture = frmGame!imgKing(iSrcColor).Picture
            Else
                frmGame!imgSquare(giKingIndex(iSrcColor)).Picture = frmGame!imgKing(iSrcColor).Picture
            End If
        End If

        If sPieceType = kPawnNotation Then
            gsMoveNotation = ""
        Else
            gsMoveNotation = sPieceType
        End If

        gsMoveNotation = gsMoveNotation & sSpecificNotation(gasSqr(Index), giLastIndex, Index)
        If sDestTag <> "" Then
            gsMoveNotation = gsMoveNotation & "x"
        End If
        
        If sPieceType = kPawnNotation And gsMoveNotation = "x" Then
            CalcCoord giLastIndex, iFile, iRank
            gsMoveNotation = Chr(iFile + 97) & "x"
        End If

        CalcCoord Index, iFile, iRank
        frmGame!imgSquare(Index).Picture = Source.Picture
        Source.Picture = LoadPicture("")
        Source.Width = 32
        Source.Height = 32
      
        gsMoveNotation = gsMoveNotation & Chr(iFile + 97) & (7 - iRank + 1)
        
        '** is this a pawn promotion move?
        If sPieceType = kPawnNotation And ((iSrcColor = kRed And iRank = 7) Or (iSrcColor = kBlue And iRank = 0)) Then
            giPawnIndex = Index
            frmPawnExchange.Show 1
            gsMoveNotation = gsMoveNotation & "=" & Right(gasSqr(Index), 1)
            gsMoveData = gsMoveData & "!,P," & giLastIndex & "," & Index & "," & Right(gasSqr(Index), 1) & ";"
        Else
            gsMoveData = gsMoveData & giLastIndex & "," & Index & ";"
        End If
        
        If sPieceType = kPawnNotation And Abs(giLastIndex - Index) = 16 Then
            If iSrcColor = kRed Then
                giEnPassant = Index - 8
            Else
                giEnPassant = Index + 8
            End If
        Else
            giEnPassant = -1
        End If
    End If
    DoStandardMove = True
End Function

Public Sub DoCastling(Index%, Source As Object)
    '** performs switch move between King and Rook
    
    Dim iSrcColor%, sPieceType$, i%

    
    
    iSrcColor = Val(Left(gasSqr(giLastIndex), 1))
    sPieceType = Right(gasSqr(giLastIndex), 1)
     
    '** arrange so king is always giLastIndex
    If Right(frmGame!imgSquare(Index), 1) = kKingNotation Then
        i = giLastIndex
        giLastIndex = Index
        Index = i
    End If
    
    If giLastIndex > Index Then
        i = -1
    Else
        i = 1
    End If

    If Index = 0 Or Index = 56 Then 'if rook is a1 or a8
        gsMoveNotation = kQueensideCastling
    Else
        gsMoveNotation = kKingsideCastling
    End If
    
    gsMoveData = gsMoveData & "!,C," & giLastIndex & "," & Index & ","

    gasSqr(Index) = kEmptySquare
    gasSqr(giLastIndex) = kEmptySquare

    frmGame!imgSquare(Index).Picture = LoadPicture("")
    frmGame!imgSquare(Index).Width = 32
    frmGame!imgSquare(Index).Height = 32

    frmGame!imgSquare(giLastIndex).Picture = LoadPicture("")
    frmGame!imgSquare(giLastIndex).Width = 32
    frmGame!imgSquare(giLastIndex).Height = 32

    DrawSquare giLastIndex
    giLastIndex = giLastIndex + i * 2 'move king 2 over
    Index = giLastIndex - i 'move rook to opposite side of king
    gasSqr(giLastIndex) = iSrcColor & kKingNotation
    frmGame!imgSquare(giLastIndex).Picture = frmGame!imgKing(iSrcColor)
    gasSqr(Index) = iSrcColor & kRookNotation
    frmGame!imgSquare(Index).Picture = frmGame!imgRook(iSrcColor)
    giKingIndex(iSrcColor) = giLastIndex
    gsMoveData = gsMoveData & giLastIndex & "," & Index & ";"
    giEnPassant = -1
End Sub

Public Function DoEnPassant(Index%, Source As Object)
    '** perform en passant move
    
    Dim iSrcColor%, iDestColor%, iOpponentColor%
    Dim iFile%, iRank%
    Dim iOpponentIndex%, sTmp$
        
    iSrcColor = Val(Left(gasSqr(giLastIndex), 1))
    iOpponentColor = (1 / iSrcColor + 1 / iSrcColor)
    iDestColor = Val(Left(gasSqr(Index), 1))
     
    If iSrcColor = kRed Then
        iOpponentIndex = Index - 8
    Else
        iOpponentIndex = Index + 8
    End If
    
    CalcCoord giLastIndex, iFile, iRank
    CalcCoord Index, iFile, iRank
    
    gasSqr(Index) = gasSqr(giLastIndex)
    gasSqr(giLastIndex) = kEmptySquare
    sTmp = gasSqr(iOpponentIndex)
    gasSqr(iOpponentIndex) = kEmptySquare
    
    '** check to see if after move, player is in check!
    If fCheckForCheck(iSrcColor) Then
        gasSqr(giLastIndex) = gasSqr(Index)
        gasSqr(Index) = kEmptySquare
        gasSqr(iOpponentIndex) = sTmp
        
        DoEnPassant = False
        Beep
        Exit Function
    End If
    
    gsMoveNotation = Chr(iFile + 97) & "x" & Chr(iFile + 97) & (7 - iRank + 1) & "ep"
    gsMoveData = gsMoveData & "!,E," & giLastIndex & "," & Index & "," & iOpponentIndex & ";"

    frmGame!imgSquare(Index).Picture = frmGame!imgSquare(giLastIndex).Picture
    frmGame!imgSquare(giLastIndex).Picture = LoadPicture("")
    frmGame!imgSquare(Index).Width = 32
    frmGame!imgSquare(Index).Height = 32
    frmGame!imgSquare(iOpponentIndex).Picture = LoadPicture("")
    frmGame!imgSquare(Index).Width = 32
    frmGame!imgSquare(Index).Height = 32
    giEnPassant = -1
        
    If gfInCheck(iSrcColor) = True Then
        gfInCheck(iSrcColor) = False
        frmGame!imgSquare(giKingIndex(iSrcColor)).Picture = frmGame!imgKing(iSrcColor).Picture
    End If

    gfInCheck(iSrcColor) = False
    DoEnPassant = True
End Function

Public Sub FindComment(iMove%)
    '** this function is FOOBAR
    
    Dim sSearch$, iPos%, sMoveStr$, iNumCRs%

    Exit Sub
    
    sMoveStr = "------ " & iMove & " ------"
    If iMove = -1 Then
        frmGame!txtComments.SelStart = 1
    Else
        'frmGame!txtComments.SelStart = 1
        
        iPos = InStr(frmGame!txtComments, sMoveStr) + Len(sMoveStr) - 1
        Do
            iPos = iPos + 1
            If Asc(Mid(frmGame!txtComments, iPos, 1)) = 13 Then '** if we hit a carriage return
                iNumCRs = iNumCRs + 1
            End If
        Loop While (iNumCRs < 2 And iPos < Len(frmGame!txtComments))
        frmGame!txtComments.SelStart = iPos + 1
    End If
    frmGame!txtComments.SelLength = 0
End Sub

Public Sub FlipBoard(iColor%)
    '** set board so iColor side is at bottom
    Dim i%, j%, iHeight%, iWidth%
    
    iWidth = 32 'scalewidth / 8
    iHeight = 32 ' scaleheight / 8
    
    For i = 0 To 7
        For j = 0 To 7
            If iColor = kBlue Then
                frmGame!imgSquare((i * 8) + j).Left = j * iWidth + giHorizBoardOffset
                frmGame!imgSquare((i * 8) + j).Top = i * iHeight + giVertBoardOffset
            Else
                frmGame!imgSquare(((7 - i) * 8) + (7 - j)).Left = j * iWidth + giHorizBoardOffset
                frmGame!imgSquare(((7 - i) * 8) + (7 - j)).Top = i * iHeight + giVertBoardOffset
            End If
        Next j
    Next i
End Sub


