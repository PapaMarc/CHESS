Attribute VB_Name = "ChessMainProcs"
'** CHSMAIN.BAS module contains all code relating to:
'**   load and unloading the chess form,
'**   and opening, sending, creating chess form messages

Option Explicit

'** vars: mapiform and ole/messaging related
Global gmfObj As Object             '** mapiform object set by mapiform.vbx
Global gmSession As Object          '** current mapi session (from gmfObj)
Global gmLoadMsg As Object          '** message user opened in Exchange
Global gmCurFolder As Object        '** folder where chess is being played
Global giCommand As Integer         '** stores mail verb (open, compose, reply)
Global gfMapifVbxUnload As Integer  '** true when mapifvbx is unloading
Global gfGameInProgress As Integer  '** is there a game in progress? (else, chess is idle)
                                    '** keeps from loading multiple instances
Global giUnloadMode As Integer      '** parameter for gmfObj.OnClose, how are we closing?

'** constants: messages
Global Const kMsgGUID = "{AC9A9507-C10B-101B-BBFE-5A5DD22B1ED3}"
Global Const kMsgClass = "IPM.Microsoft.Chess"
Global Const kConvIndexLen% = 44                 '** length of Exchange conversation index string

'** constants: mail verbs
Global Const kCompose = 0
Global Const kOpen = 1
Global Const kReply = 2

'** constants: strings to be internationalized
Global Const kNewChessPrefix$ = "New Chess Match" '** prefixes new game message subject
Global Const kChessPrefix$ = "Chess Match"
Global Const kErrOneGame$ = "You must first finish playing the current chess move."
Global Const kErrStrangeCommand$ = "Strange command error!"
Global Const kErrSending$ = "Error sending Chess Form. "
Global Const kChessErrorOccured$ = "An error occured in the Chess Form. "
Global Const kDefaultMatchname$ = "Exchange Chess"
Global Const kNoOpponent$ = "<none selected>"
Global Const kSurrenderPrompt$ = "Are you sure you want to surrender? This will forfeit the game."
Global Const kSelectOpponent$ = "Select Opponent"
Global Const kErrNotExchangeUser$ = "Opponent must be an Exchange user!"
Global Const kErrCantFindRecipient$ = "Recipient's address cannot be found. Please select a different opponent."
Global Const kErrCantPlayYourself$ = "Please select someone other than yourself as an opponent."
Global Const kErrNoFirstMessage$ = "First message does not exist in current folder."
Global Const kErrCommentsTrunc$ = "<Exchange Chess comments have exceeded the allowable  limit -- earlier text has been truncated.>"
Global Const kSurrender = "S&urrender" '** button caption
Global Const kUndo = "&Undo" '** button caption
Global Const kSurrenderNotation = "Surrender"
Global Const kVersus = " vs. "
Global Const kCheckmate = "CHECKMATE!!! "
Global Const kDefeats = " defeats "
Global Const kGameOver = "Game Over"
Global Const kSurrendered = " surrendered!"

Global Const kQueensideCastling = "O-O-O"
Global Const kKingsideCastling = "O-O"
Global Const kKingNotation = "K"
Global Const kQueenNotation = "Q"
Global Const kRookNotation = "R"
Global Const kBishopNotation = "B"
Global Const kKnightNotation = "N"
Global Const kPawnNotation = "P"

Global Const kCommentsTruncSize = 32500     '** game comments get truncated when 32500 characters is reached
Sub Main()
    '** load frmAdmin to establish link between viewer & chess

    Load frmAdmin
End Sub

Sub DoVerb(mF As Object, Verb As Integer)
    '** setup mapiform, ole/messaging objects

    Dim i%
    Dim sFolderID$, sStoreID$, iTmpCommand%
    
    On Error GoTo ErrDoVerb
    
    If mF.Message.subject <> "" Then
        iTmpCommand = kOpen
    Else
        iTmpCommand = kCompose
    End If
    
    '** handle multiple games open
    If gfGameInProgress Then
        If iTmpCommand = kOpen And giCommand = kOpen Then
            If mF.Message.ID <> gmLoadMsg.ID Then '** already working this msg
                MsgBox kErrOneGame, vbExclamation
                mF.OnClose vbFormCode
                Set mF = Nothing
            End If
        Else
            MsgBox kErrOneGame, vbExclamation
            mF.OnClose vbFormCode
            Set mF = Nothing
        End If
        Screen.ActiveForm.Show '** bring chess to front
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    '** initialize some game variables
    gfGameInProgress = True
    gsMoveNotation = ""
    gfMovedPiece = 0
    gfGameOver = False
    gfOpponentSurrender = False
    gfInCheck(kBlue) = False
    gfInCheck(kRed) = False
    
    Set gmfObj = mF '** make a reference to the mapiform object
    gmfObj.OwnerChildWindow = 1 '** tell the mf object we have a child window
    Set gmSession = mF.Session
    
    giCommand = iTmpCommand
    
    
    Select Case giCommand
        Case kCompose   '** starting new chess match
            InitFrmNewGame
        Case kOpen, kReply  '** opened or replied to a move from the viewer
            Set gmLoadMsg = mF.Message
            sFolderID = sGetFolderID(gmLoadMsg)
            sStoreID = gmLoadMsg.storeID
            If Len(gmLoadMsg.ConversationIndex) = kConvIndexLen Then '** opened new game message
                Set gmCurFolder = gmSession.GetFolder(sFolderID, sStoreID)
                InitFrmNewGame
                UpdateTasklist kAddTask, frmNewGame
                frmNewGame.Show
            Else '** opened chess board message
                gfPlayed = gmLoadMsg.Fields.Item("Played").Value
                giMoveNum = gmLoadMsg.Fields.Item("CountMove").Value
                If gfPlayed = True Then '** if this move is played on, find the newest
                                        '** chess move of this game in the current folder
                    Set gmCurFolder = gmSession.GetFolder(sFolderID, sStoreID)
                    i = iGetNewestMsg(gmLoadMsg)
                End If
                
                On Error GoTo ErrDoVerb
                
                gfByOpponent = gmLoadMsg.Fields.Item("ByOpponent").Value
                gfPlayed = gmLoadMsg.Fields.Item("Played").Value
                InitFrmGame
                UpdateTasklist kAddTask, frmGame
                frmGame.Show
            End If
        Case Else   '** strange mail command
            MsgBox kErrStrangeCommand, vbExclamation
            Error 51 '** internal error
    End Select
    Exit Sub
ErrDoVerb:
    MsgBox kChessErrorOccured & Err.Description & ": " & Err.Number, vbExclamation
    giUnloadMode = vbFormCode
    Terminate frmGame
End Sub
Public Function iGetNewestMsg(mDestMsg As Object)
    '** starts with newest chess message in folder
    '** cycles down through chess messages until one is found
    '** which has the same game conversation index and
    '** was created after the mDestMsg

    Dim sGameConvIndex$, sMoveConvIndex
    Dim mSrcMsg As Object
    Dim mMessages As Object
    Dim fDone%
    Dim dTimeRecieved
    
    On Error GoTo errGetNewestMsg
    iGetNewestMsg = 1
    
    dTimeRecieved = mDestMsg.TimeReceived
    sGameConvIndex = Left(mDestMsg.ConversationIndex, kConvIndexLen)
    
    On Error GoTo errGetNewestMsg '** error will be created when there are no more messages to cycle through
    Set mMessages = gmCurFolder.Messages
    Set mSrcMsg = mMessages.GetLast(kMsgClass)
    
    While fDone = False
        If mSrcMsg.TimeReceived >= dTimeRecieved Then
            sMoveConvIndex = mSrcMsg.ConversationIndex
            If Len(sMoveConvIndex) <> kConvIndexLen And Left(sMoveConvIndex, kConvIndexLen) = sGameConvIndex Then
                Set mDestMsg = mSrcMsg
                iGetNewestMsg = 0
                fDone = True
            End If
        End If
        Set mSrcMsg = mMessages.GetPrevious
    Wend
    
errGetNewestMsg:
    Set mMessages = Nothing
    Set mSrcMsg = Nothing
End Function

Sub ParseCommaString(ByVal sData$, sDest$, iLoc%)
    '** Go to iLoc offset of sData. Read chars into sDest until comma.
    '** Adjust iLoc to after comma.

    sDest = ""
    While Mid(sData$, iLoc%, 1) <> "," Or iLoc% > Len(sData)
        sDest = sDest & Mid(sData$, iLoc%, 1)
        iLoc% = iLoc% + 1
    Wend
    iLoc% = iLoc% + 1
End Sub

Sub ParseSemicolonString(ByVal sData$, sDest, iLoc%)
    '** Go to iLoc offset of sData. Read chars into sDest until semicolon.
    '** Adjust iLoc to after semicolon.

    sDest = ""
    While Mid(sData$, iLoc%, 1) <> ";" Or iLoc% > Len(sData)
        sDest = sDest & Mid(sData$, iLoc%, 1)
        iLoc% = iLoc% + 1
    Wend
    iLoc% = iLoc% + 1
End Sub






Function sGetBigPropToken$(sBigProp$)
    '** Return all chars in sBigProp until carriage return.
    '** Strip these chars off of sBigProp.
    Dim iLoc%
    
    iLoc = InStr(sBigProp, Chr(13))
    If iLoc <> 0 Then
        sGetBigPropToken$ = Left$(sBigProp$, iLoc - 1)
        sBigProp$ = Right(sBigProp$, Len(sBigProp$) - iLoc)
    End If
End Function
Function sStuffMainProperty()
    '** stuffs game data into "GameData" field of a message
    '** "GameData" is a return-char (13) delimited string
    Dim sProp$

    sProp$ = gsSubject$ & Chr(13)
    sProp$ = sProp$ & gasPlayerName(kRed) & Chr(13) & gasPlayerName(kBlue) & Chr(13)
    sProp$ = sProp$ & gasPlayerAddress(kRed) & Chr(13) & gasPlayerAddress(kBlue) & Chr(13)
    
    sProp$ = sProp$ & giCurTurn & Chr(13) & giEnPassant & Chr(13)

    sProp$ = sProp$ & gfRookMoved(0) & Chr(13) & gfRookMoved(1) & Chr(13) & gfRookMoved(2) & Chr(13) & gfRookMoved(3) & Chr(13)
    sProp$ = sProp$ & gfKingMoved(kBlue) & Chr(13) & gfKingMoved(kRed) & Chr(13)
    
    sProp$ = sProp$ & sStuffListToString(frmGame!lstMoves) & Chr(13)
    sProp$ = sProp$ & gsMoveData & Chr(13)
    sStuffMainProperty = sProp$
End Function
Sub UnstuffStringToList(sSrc$, lstDest As ListBox)
    '** fills out items of a listbox with data from comma-delimited string
    Dim i%, iLoc%, iNumItems%, sParsed$
    
    lstDest.Clear
    
    iLoc% = 1
    ParseCommaString sSrc, sParsed, iLoc%
    iNumItems = Val(sParsed)
    
    For i = 1 To iNumItems%
        ParseCommaString sSrc, sParsed, iLoc%
        lstDest.AddItem sParsed
    Next i
End Sub
Sub UnstuffBoardData(ByVal sSrcString$)
    '** parse through a comma-delimited string holding board info
    '** Ex: 2R,2N,2B... (color-piece-comma)
    '** if matches gasSqr, do nothing
    '** else, set appropriate icon or none

    Dim i%, iLoc%, sDestString$
    Dim iColor%, sType$

    iLoc = 1
    For i = 0 To 63
        ParseCommaString sSrcString, sDestString, iLoc%
        If sDestString <> gasSqr(i) Then
            gasSqr(i) = sDestString
            If sDestString <> "" Then
                iColor% = Val(Left(sDestString, 1))
                sType = Right(sDestString, 1)
                Select Case sType
                    Case kRookNotation
                        frmGame!imgSquare(i) = frmGame!imgRook(iColor)
                    Case kKnightNotation
                        frmGame!imgSquare(i) = frmGame!imgKnight(iColor)
                    Case kBishopNotation
                        frmGame!imgSquare(i) = frmGame!imgBishop(iColor)
                    Case kKingNotation
                        frmGame!imgSquare(i) = frmGame!imgKing(iColor)
                        giKingIndex(iColor) = i
                    Case kQueenNotation
                        frmGame!imgSquare(i) = frmGame!imgQueen(iColor)
                    Case kPawnNotation
                        frmGame!imgSquare(i) = frmGame!imgPawn(iColor)
                End Select
            Else
                frmGame!imgSquare(i).Picture = Nothing
                frmGame!imgSquare(i).Width = 32
                frmGame!imgSquare(i).Height = 32
            End If
        End If
    Next i
End Sub






Function sStuffListToString$(lstSrc As ListBox)
    '** stuff the items of listbox (lstSrc) into a comma-delimited string
    '** Ex: the move history listbox would convert to "e4,Bxf1,Kb2"
    Dim i%, sRet$

    sRet$ = lstSrc.ListCount & ","
    For i = 0 To lstSrc.ListCount - 1
        sRet$ = sRet$ & lstSrc.List(i) & ","
    Next i
    sStuffListToString$ = sRet$
End Function


Sub UnstuffMainProperty(mMessage As Object)
    '** grabs game data from "GameData" field of a message
    '** "GameData" is a return char (13) delimited string
    
    Dim sBigProp$, i%
    Dim sFiller$
    
    sBigProp = mMessage.Fields.Item("GameData").Value

    gsSubject$ = sGetBigPropToken(sBigProp)
    frmGame!lblSubject.Caption = Chr(34) & gsSubject & Chr(34) 'asc 34 = '"'
    gasPlayerName(kRed) = sGetBigPropToken(sBigProp)
    gasPlayerName(kBlue) = sGetBigPropToken(sBigProp)
    sFiller = sGetBigPropToken(sBigProp)
    sFiller = sGetBigPropToken(sBigProp)

    giCurTurn = CInt(sGetBigPropToken(sBigProp))
    giEnPassant = CInt(sGetBigPropToken(sBigProp))
    giOldEnPassant = giEnPassant
    For i = 0 To 3
        gfRookMoved(i) = CInt(sGetBigPropToken(sBigProp))
    Next i
    gfKingMoved(kBlue) = CInt(sGetBigPropToken(sBigProp))
    gfKingMoved(kRed) = CInt(sGetBigPropToken(sBigProp))

    UnstuffStringToList sGetBigPropToken(sBigProp), frmGame!lstMoves
    gsMoveData$ = sGetBigPropToken(sBigProp)
    frmGame!txtComments.Text = mMessage.Text
End Sub

Sub Terminate(frmClose As Form)
    '** hides chess form
    '**     and clears up ole/messaging and mapiform objects related to previous game
    '** frmClose is the form calling Terminate()
    
    frmClose.Visible = False
    Set gmCurFolder = Nothing
    Set gmSession = Nothing
    Set gmLoadMsg = Nothing
    If Not gmfObj Is Nothing Then gmfObj.OnClose giUnloadMode
    Set gmfObj = Nothing
    gfGameInProgress = False
    UpdateTasklist kRemoveTask, frmClose
End Sub











Public Sub SendMessage()
    '** send chess message to yourself and opponent
    
    Dim iOpponentColor As Integer
    Dim iOldCurTurn As Integer
    Dim mSendMsg As Object
    Dim mSaveMsg As Object
    Dim mSender As Object
    Dim mSendRecip As Object
    Dim mSaveRecip As Object
    Dim mSendFields As Object
    Dim mSaveFields As Object
    Dim mTmpField
    
    On Error GoTo ErrSendMessage
    
    Screen.MousePointer = vbHourglass
    
    iOpponentColor = 1 / giCurTurn + 1 / giCurTurn '** changes kBlue (1) to kRed (2) and vice versa
    iOldCurTurn = giCurTurn
    
    Set mSendMsg = gmSession.Outbox.Messages.Add
    Set mSaveMsg = gmSession.Outbox.Messages.Add
    giMoveNum = frmGame!lstMoves.ListCount
    
    Set mSendRecip = mSendMsg.Recipients.Add(gasPlayerName(iOpponentColor), gasPlayerAddress(iOpponentColor))
    mSendRecip.Resolve
    Set mSaveRecip = mSaveMsg.Recipients.Add(gasPlayerName(iOldCurTurn), gasPlayerAddress(iOldCurTurn))
    mSaveRecip.Resolve
    
    Set mSendFields = mSendMsg.Fields
    Set mSaveFields = mSaveMsg.Fields
    
    Set mTmpField = mSendFields.Add("CountMove", vbInteger, giMoveNum)
    Set mTmpField = Nothing
    
    Set mTmpField = mSendFields.Add("Move", vbString, gsMoveNotation)
    Set mTmpField = Nothing
    mSendMsg.type = kMsgClass
    
    Set mTmpField = mSaveFields.Add("CountMove", vbInteger, giMoveNum)
    Set mTmpField = Nothing
    
    Set mTmpField = mSaveFields.Add("Move", vbString, gsMoveNotation)
    Set mTmpField = Nothing
    mSaveMsg.type = kMsgClass
    
    If giCurTurn = kRed Then
        Set mTmpField = mSendFields.Add("Player", vbString, gasPlayerName(kRed))
        Set mTmpField = Nothing
        Set mTmpField = mSaveFields.Add("Player", vbString, gasPlayerName(kRed))
        Set mTmpField = Nothing
        giCurTurn = kBlue
    Else
        Set mTmpField = mSendMsg.Fields.Add("Player", vbString, gasPlayerName(kBlue))
        Set mTmpField = Nothing
        Set mTmpField = mSaveMsg.Fields.Add("Player", vbString, gasPlayerName(kBlue))
        Set mTmpField = Nothing
        giCurTurn = kRed
    End If
    
    If gasSqr(0) <> kRed & kRookNotation Then gfRookMoved(0) = True
    If gasSqr(7) <> kRed & kRookNotation Then gfRookMoved(1) = True
    If gasSqr(56) <> kBlue & kRookNotation Then gfRookMoved(2) = True
    If gasSqr(63) <> kBlue & kRookNotation Then gfRookMoved(3) = True
    
    If gasSqr(4) <> kRed & kKingNotation Then gfKingMoved(kRed) = True
    If gasSqr(60) <> kBlue & kKingNotation Then gfKingMoved(kBlue) = True

    '** grab data from glob. variables and stuff them into message object
    Set mTmpField = mSendFields.Add("GameData", vbString, sStuffMainProperty())
    Set mTmpField = Nothing
    
    Set mTmpField = mSaveFields.Add("GameData", vbString, mSendFields.Item("GameData").Value)
    Set mTmpField = Nothing
    
    mSendMsg.Text = frmGame!txtComments.Text
    mSaveMsg.Text = frmGame!txtComments.Text

    mSendMsg.ConversationTopic = gsSubject
    mSaveMsg.ConversationTopic = gsSubject
    
    If giCommand = kCompose Then
        '** if this is the first move a match then send out New Chess Match messages
        '** NG = New Game
        Dim mNGSendMsg As Object
        Dim mNGSaveMsg As Object
        Dim mNGSendRecip As Object
        Dim mNGSaveRecip As Object
        Dim mNGSendFields As Object
        Dim mNGSaveFields As Object
        Dim sConvIndex$

        Set mNGSendMsg = gmSession.Outbox.Messages.Add
        Set mNGSaveMsg = gmSession.Outbox.Messages.Add

        mNGSendMsg.ConversationIndex = sNewConversationIndex()
        mNGSendMsg.ConversationTopic = gsSubject
        mNGSendMsg.subject = kNewChessPrefix
        mNGSendMsg.Text = gasPlayerName(kRed)
        mNGSendMsg.type = kMsgClass

        mNGSaveMsg.ConversationIndex = mNGSendMsg.ConversationIndex
        mNGSaveMsg.ConversationTopic = gsSubject
        mNGSaveMsg.subject = mNGSendMsg.subject
        mNGSaveMsg.Text = mNGSendMsg.Text
        mNGSaveMsg.type = mNGSendMsg.type

        Set mNGSendFields = mNGSendMsg.Fields
        Set mNGSaveFields = mNGSaveMsg.Fields
        
        If iOldCurTurn = kBlue Then
            Set mTmpField = mNGSendFields.Add("Player", vbString, gasPlayerName(kBlue))
            Set mTmpField = Nothing
            Set mTmpField = mNGSaveFields.Add("Player", vbString, gasPlayerName(kBlue))
            Set mTmpField = Nothing
        Else
            Set mTmpField = mNGSendFields.Add("Player", vbString, gasPlayerName(kRed))
            Set mTmpField = Nothing
            Set mTmpField = mNGSaveFields.Add("Player", vbString, gasPlayerName(kRed))
            Set mTmpField = Nothing
        End If

        Set mTmpField = mNGSendFields.Add("Move", vbString, kNewChessPrefix)
        Set mTmpField = Nothing
        Set mTmpField = mNGSaveFields.Add("Move", vbString, kNewChessPrefix)
        Set mTmpField = Nothing

        Set mNGSendRecip = mNGSendMsg.Recipients.Add(gasPlayerName(iOpponentColor), gasPlayerAddress(iOpponentColor))
        mNGSendRecip.Resolve
        Set mNGSaveRecip = mNGSaveMsg.Recipients.Add(gasPlayerName(iOldCurTurn), gasPlayerAddress(iOldCurTurn))
        mNGSaveRecip.Resolve
        
        mNGSendMsg.Update
        mNGSaveMsg.Update
        sConvIndex = mNGSendMsg.ConversationIndex

        mNGSaveMsg.Send False, False
        mNGSendMsg.Send False, False
        
        mSendMsg.ConversationIndex = sAddConversationIndex(Left$(sConvIndex, kConvIndexLen)) ' & LongToHex(giMoveNum + 1) & Hex(Int(Rnd * 250) + 2)
        mSaveMsg.ConversationIndex = mSendMsg.ConversationIndex

        Set mNGSendMsg = Nothing
        Set mNGSaveMsg = Nothing
        Set mNGSendRecip = Nothing
        Set mNGSaveRecip = Nothing
        Set mNGSendFields = Nothing
        Set mNGSaveFields = Nothing
    Else
        mSendMsg.ConversationIndex = sAddConversationIndex(Left$(gmLoadMsg.ConversationIndex, kConvIndexLen)) ' & LongToHex(giMoveNum + 1) & Hex(Int(Rnd * 250) + 2)
        mSaveMsg.ConversationIndex = mSendMsg.ConversationIndex
    End If
    
    mSendMsg.subject = gsMoveNotation
    mSaveMsg.subject = gsMoveNotation

    Set mTmpField = mSendFields.Add("ByOpponent", vbBoolean, True)
    Set mTmpField = Nothing
    Set mTmpField = mSendFields.Add("Played", vbBoolean, False)
    Set mTmpField = Nothing
    
    Set mTmpField = mSaveFields.Add("ByOpponent", vbBoolean, False)
    Set mTmpField = Nothing
    Set mTmpField = mSaveFields.Add("Played", vbBoolean, True)
    Set mTmpField = Nothing

    mSaveMsg.Update False
    mSendMsg.Update False
    mSaveMsg.Send False, False  '** send message to yourself
    mSendMsg.Send False, False  '** send message to opponent
        
    Set mSendMsg = Nothing
    Set mSaveMsg = Nothing
    Set mSender = Nothing
    Set mSaveRecip = Nothing
    Set mSendRecip = Nothing
    Set mSendFields = Nothing
    Set mSaveFields = Nothing
    
    If giCommand = kOpen Then
        '** update played flag on loaded message
        gmLoadMsg.Fields.Item("Played").Value = True
        gmLoadMsg.Update
    End If
        
    Screen.MousePointer = vbArrow
    giUnloadMode = vbFormCode
    Terminate frmGame
    Exit Sub
ErrSendMessage:
    Set mSendMsg = Nothing
    Set mSaveMsg = Nothing
    Set mSender = Nothing
    Set mSaveRecip = Nothing
    Set mSendRecip = Nothing
    Set mSendFields = Nothing
    Set mSaveFields = Nothing
    
    Screen.MousePointer = vbArrow
    MsgBox kErrSending & Err.Description & ": " & Err.Number, vbExclamation
    giUnloadMode = vbFormCode
    Terminate frmGame
    Exit Sub
End Sub

Public Sub InitFrmNewGame()
    '** initialize the frmNewGame when user opens/creates a New Chess Match message
    
    Const iLogoOffset% = 7
    Const kWhite& = &HFFFFFF
    
    Load frmNewGame
    CenterFormOnScreen frmNewGame
    frmNewGame!imgLogo.Left = iLogoOffset% * Screen.TwipsPerPixelX
    frmNewGame!imgLogo.Top = frmNewGame.ScaleHeight - (frmNewGame!imgLogo.Height + (iLogoOffset% * Screen.TwipsPerPixelY))
    
    If giCommand = kCompose Then    '** if user is creating a new New Chess Match message
        frmNewGame!fraMatch.Enabled = True
        frmNewGame!fraOpponent.Enabled = True
        frmNewGame!cmdAddress.Enabled = True
        frmNewGame!txtSelectOpponent.Text = kNoOpponent
        frmNewGame!txtSubject.Locked = False
        frmNewGame!txtSubject = kDefaultMatchname
        frmNewGame!txtSubject.SelStart = 0
        frmNewGame!txtSubject.SelLength = Len(frmNewGame!txtSubject.Text)
        frmNewGame!txtSubject.BackColor = kWhite
        frmNewGame!txtSubject.TabStop = True
        frmNewGame!cmdBegin.Enabled = False
    Else    '** if user is opening a New Chess Match message
        frmNewGame!fraMatch.Enabled = False
        frmNewGame!fraOpponent.Enabled = False
        frmNewGame!cmdAddress.Enabled = False
        frmNewGame!txtSelectOpponent.Text = gmLoadMsg.Text
        frmNewGame!txtSubject.Locked = True
        frmNewGame!txtSubject.BackColor = frmNewGame!txtSelectOpponent.BackColor
        frmNewGame!txtSubject.TabStop = False
        frmNewGame!txtSubject.Text = gmLoadMsg.ConversationTopic
    End If
    frmNewGame.Visible = True
    Screen.MousePointer = vbArrow
End Sub

Public Sub InitFrmGame()
    '** initialize frmGame after chess move is opened
    
    Dim i%, iOpponent%
    Dim mRecip As Object

    frmGame.WindowState = 0
    frmGame!cmdUndo.Enabled = True
    frmGame!cmdUndo.Caption = kSurrender
    For i = 0 To 63
        gasSqr(i) = kEmptySquare
        frmGame!imgSquare(i).Picture = Nothing
        frmGame!imgSquare(i).Width = 32
        frmGame!imgSquare(i).Height = 32
    Next i
    
    '** load game data from gmLoadMsg
    If giCommand <> kCompose Then
        UnstuffMainProperty gmLoadMsg
        
        If gfPlayed = True Then
            frmGame!cmdSend.Enabled = False
        End If
        
        '** setup player's name/addresses
        Set mRecip = gmSession.CurrentUser
        gasPlayerName(giCurTurn) = mRecip.Name
        gasPlayerAddress(giCurTurn) = mRecip.type & ":" & mRecip.address
        Set mRecip = Nothing
            
        If gfPlayed = False Then
            iOpponent = 1 / giCurTurn + 1 / giCurTurn '** changes kBlue (1) to kRed (2) and vice versa
            Set mRecip = gmLoadMsg.Sender
            gasPlayerName(iOpponent) = mRecip.Name
            gasPlayerAddress(iOpponent) = mRecip.type & ":" & mRecip.address
            Set mRecip = Nothing
        End If
    End If
    
    '** flip the board if necessary
    If frmGame!imgSquare(0).Top < frmGame!imgSquare(63).Top Then 'setup for blue
        If (giCurTurn = kRed And gfByOpponent = True) Or (giCurTurn = kBlue And gfByOpponent = False) Then
            FlipBoard kRed
        End If
    Else 'setup for red
        If (giCurTurn = kRed And gfByOpponent = False) Or (giCurTurn = kBlue And gfByOpponent = True) Then
            FlipBoard kBlue
        End If
    End If
    
    If giMoveNum > frmGame!lstMoves.ListCount Then
        frmGame!lstMoves.Selected(frmGame!lstMoves.ListCount - 1) = True
    ElseIf giMoveNum <> 0 Then
        frmGame!lstMoves.Selected(giMoveNum - 1) = True
    Else
        ReplayMoves 0
        PrintPlayerNames giCurTurn
    End If

    frmGame!txtComments.Locked = False
    If gfPlayed = True Then
        frmGame!txtComments.Locked = True
    ElseIf giMoveNum > 0 Then
        If Len(frmGame!txtComments) > kCommentsTruncSize Then '** are comments to large?
            TruncComments
        End If
        frmGame!txtComments = Chr(13) & Chr(10) & "------ " & frmGame!lstMoves.ListCount & " ------" & Chr(13) & Chr(10) & frmGame!txtComments
    End If
    
    '** make kings in check flash
    frmGame.Visible = True
    For i = kBlue To kRed
        If gfInCheck(i) Then
            If fCheckForCheckmate(i) Then   '** if player is in checkmate, pop up alert
                gfGameOver = True
                CheckmateAlert
            End If
        End If
    Next i

    If gfOpponentSurrender Then
        gfGameOver = True
        If gfByOpponent = True Then
            SurrenderAlert '** if opponent surreneder, pop up alert
        End If
    End If
    
    AdjustUndoButton
    frmGame.lstMoves.SetFocus
    Screen.MousePointer = vbArrow
    Exit Sub
End Sub

Public Function sGetFolderID(mMsg As Object)
    '** if message is inserted, there is no folderID & error occurs
    '**     function created to localize error
    
    On Error GoTo errGetFolderID
    
    sGetFolderID = mMsg.folderID
    Exit Function
errGetFolderID:
    sGetFolderID = ""
End Function

Public Sub TruncComments()
    '** if frmGame!txtComments is too large to prepend text to it, truncate
    '** it to kTruncSize. Never truncate in the middle of a divider.
    
    Dim iTruncLoc, iOldTruncLoc
    Const sDivRight$ = " ------"
    
    Do
        iOldTruncLoc = iTruncLoc
        iTruncLoc = InStr(iTruncLoc + 1, frmGame!txtComments, sDivRight)
    Loop While (iTruncLoc > 0 And iTruncLoc < kCommentsTruncSize - 250)
    
    frmGame!txtComments = Left(frmGame!txtComments, iOldTruncLoc + Len(sDivRight) + 1) & kErrCommentsTrunc
     
    Exit Sub
End Sub
