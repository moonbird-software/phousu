Attribute VB_Name = "AI"
Option Explicit

Function GetNextPlrMode(Player As CardDeck, Dealer As CardDeck, Trick As CardDeck, fKill As Boolean) As Integer

End Function
Function IsCardFaceUp(Source As CardDeck, Dest As CardDeck) As Boolean
    IsCardFaceUp = IsCardFaceUpBasic(Source, Dest)
    Select Case Source.Index
    Case IDD_PLAYER1_ALT1, IDD_PLAYER1_ALT2, IDD_PLAYER1_ALT3, IDD_PLAYER2_ALT1, IDD_PLAYER2_ALT2, IDD_PLAYER2_ALT3, IDD_PLAYER3_ALT1, IDD_PLAYER3_ALT2, IDD_PLAYER3_ALT3, IDD_PLAYER4_ALT1, IDD_PLAYER4_ALT2, IDD_PLAYER4_ALT3
        IsCardFaceUp = True
    End Select
    Select Case Dest.Index
    Case IDD_USER
        IsCardFaceUp = True
    Case IDD_PLAYER1_ALT1, IDD_PLAYER1_ALT2, IDD_PLAYER1_ALT3, IDD_PLAYER2_ALT1, IDD_PLAYER2_ALT2, IDD_PLAYER2_ALT3, IDD_PLAYER3_ALT1, IDD_PLAYER3_ALT2, IDD_PLAYER3_ALT3, IDD_PLAYER4_ALT1, IDD_PLAYER4_ALT2, IDD_PLAYER4_ALT3
        If Source.Index <> IDD_DEALER Then
            IsCardFaceUp = True
        End If
    End Select
End Function

Sub ActionClick()
Dim iCard As Integer, iDeck As Integer
    
    frmMain.cmdAction.Visible = False
    Game.FirstTurn = False
    Game.Mode = IDM_NORMAL
    
    Select Case frmMain.cmdAction.Caption
    Case IDS_ACTION_PLAY_CARD, IDS_ACTION_PLAY_CARDS
        AnimMoveSelCards Deck(IDD_USER), Deck(IDD_TRICK), 2
        
    Case IDS_ACTION_TRY_DECK
        AnimTryFromDeck Deck(IDD_USER), Deck(IDD_DEALER), Deck(IDD_TRICK)
    
    Case IDS_ACTION_TAKE_CARDS
        AnimTakeCards Deck(IDD_TRICK), Deck(IDD_USER)
        
    Case IDS_ACTION_CALL
        ' ep‰ile kortteja
        
    Case IDS_ACTION_PLACE_CARDS
        
        With frmMain
            .mnuGameNew.Enabled = True
            .mnuGameDemo.Enabled = True
        End With
        
        AnimMoveCards Deck(IDD_USER), Deck(IDD_USER_ALT1), cmSelected, , 3, True
        
        GetFirstPlayer
    
    End Select
    
    RotateTurn
    
End Sub
Sub AI_SelectCardsRandomSpanish(Player As CardDeck, Trick As CardDeck)
Dim iDeck As Integer, iDeckDest(2) As Integer, iRank As Integer
    GetDestDeckSpanish Player, iDeckDest
    Do
        iDeck = Int(Rnd * 3)
    Loop Until CountCards(Deck(iDeckDest(iDeck))) = 1
    iRank = GetRank(GetTopCard(Trick))
    AnimPopCards Deck(iDeckDest(iDeck)), Trick, 1, False
    If CheckRules(iRank, GetRank(GetTopCard(Trick))) Then
    Else
        AnimTakeCards Trick, Player
    End If
End Sub
Sub AI_SelectTableCardsSpanish(Player As CardDeck)
Dim iRank As Integer, iCard As Integer
Dim nCards As Integer
Dim iDeckDest(2) As Integer

    ' 3-of-a-kind
    For iRank = K To 3
        If SelCards(Player, iRank, 3, True) = 3 Then
            nCards = nCards + SelCards(Player, iRank, 3 - nCards)
        End If
    Next iRank
    
    ' deuces
    If CardInDeck(Player, 2) Then
        nCards = nCards + SelCards(Player, 2, 3 - nCards)
    End If
    
    ' tens
    If CardInDeck(Player, 10) Then
        nCards = nCards + SelCards(Player, 10, 3 - nCards)
    End If
    
    ' aces
    If CardInDeck(Player, A) Then
        nCards = nCards + SelCards(Player, A, 3 - nCards)
    End If
   
    ' pairs
    If nCards = 1 Then
        For iRank = K To 3
            If SelCards(Player, iRank, 2, True) = 2 Then
                nCards = nCards + SelCards(Player, iRank, 3 - nCards)
            End If
        Next iRank
    End If
    
    ' highest cards
    iRank = K
    Do Until nCards = 3 Or iRank = 3
        iRank = iRank - 1
        If CardInDeck(Player, iRank) Then
            nCards = nCards + SelCards(Player, iRank, 3 - nCards)
        End If
    Loop
    
    ' move cards to table
    GetDestDeckSpanish Player, iDeckDest
    AnimMoveCards Player, Deck(iDeckDest(0)), cmSelected, , 3, True
    
End Sub
Function CombineDeckSpanish(Player As CardDeck) As CardDeck
Dim iDeck As Integer, iDeckDest(2) As Integer, iCard As Integer
    GetDestDeckSpanish Player, iDeckDest
    ClearDeck CombineDeckSpanish
    For iDeck = 0 To 2
        If Deck(iDeckDest(iDeck)).Card(1) <> -1 Then
            CombineDeckSpanish.Card(iCard) = Deck(iDeckDest(iDeck)).Card(1)
            iCard = iCard + 1
        End If
    Next iDeck
End Function
Function CountPlayers() As Integer
Dim iPlr As Integer
Dim nPlr As Integer
    nPlr = MAX_PLAYERS
    For iPlr = 0 To MAX_PLAYERS - 1
        If CountCards(Deck(iPlr)) = 0 And GetCardLayerSpanish(Deck(iPlr)) = 0 Then
            nPlr = nPlr - 1
        End If
    Next iPlr
    CountPlayers = nPlr
End Function
Sub DealCardsDoneHook()
Dim iPlr As Integer
    Select Case Rules.GameType
    Case IDG_NORMAL
        Game.Mode = IDM_NORMAL
        
    Case IDG_SPANISH
        If Game.Demo Then
            Game.Mode = IDM_NORMAL
            For iPlr = 0 To MAX_PLAYERS - 1
                AI_SelectTableCardsSpanish Deck(iPlr)
            Next iPlr
        Else
            Game.Mode = IDM_CHOOSE_3
            With frmMain
                .mnuGameNew.Enabled = False
                .mnuGameDemo.Enabled = False
            End With
            For iPlr = 0 To MAX_PLAYERS - 1
                If iPlr <> IDD_USER Then
                    AI_SelectTableCardsSpanish Deck(iPlr)
                End If
            Next iPlr
        End If
    End Select
End Sub
Sub DeckClick(Deck As CardDeck, Dealer As CardDeck, Trick As CardDeck, User As CardDeck, Optional ByVal iCardClk As Integer)
Dim nCards As Integer, nSel As Integer
Dim iCard As Integer, iRankSel As Integer, iRankClk As Integer, iRankTop As Integer, iDeckDest As Integer, iDeck As Integer
Dim fCanSel As Boolean, fDeuce As Boolean, fAction As Boolean
Dim sTableClick As String
Dim PlayerAlt As CardDeck
    
    ' count cards in deck
    If iDeck <> -1 Then
        nCards = CountCards(Deck)
    End If
    
    Select Case Game.Mode
    Case IDM_NORMAL
        Select Case Deck.Index
        Case IDD_USER              ' human player's deck
            
            ' get ranks
            iRankClk = GetRank(Deck.Card(iCardClk))
            iRankSel = GetSelRank(Deck, fDeuce)
            
            ' check rules
            If Game.FirstTurn Then
                ' only lowest cards can be selected on first turn
                If iRankClk = Game.FirstCard Then
                    fCanSel = True
                Else
                    SetStatus Replace(IDS_STATUS_CHOOSE_LOWEST_RANK, "%s", GetCardName(Game.FirstCard))
                End If
            Else
                If iRankSel = 0 Then
                    ' no cards selected
                    If Rules.DeucesWild = True And fDeuce = True Then
                        ' if deuces wild and a two is selected, anything goes
                        fCanSel = True
                    Else
                        If CountCards(Trick) = 0 Then
                            ' check against empty trick
                            fCanSel = CheckRules(0, iRankClk, True)
                        Else
                            ' check against top card on trick
                            iRankTop = GetRank(GetTopCard(Trick))
                            fCanSel = CheckRules(iRankTop, iRankClk, True)
                        End If
                    End If
                Else
                    ' cards selected
                    If iRankSel = iRankClk Then
                        ' same rank selected as earlier
                        fCanSel = CheckRules(iRankSel, iRankClk, True)
                    Else
                        ' different rank selected than earlier
                        Select Case iRankSel
                        Case 10
                            SetStatus IDS_STATUS_ONLY_ONE_KILL
                        Case 1
                            If Rules.Only10s = False Then
                                SetStatus IDS_STATUS_ONLY_ONE_KILL
                            Else
                                SetStatus Replace(IDS_STATUS_CHOOSE_RANK_OR_SUIT, "%s", GetCardName(iRankSel))
                            End If
                        Case Else
                            SetStatus Replace(IDS_STATUS_CHOOSE_RANK_OR_SUIT, "%s", GetCardName(iRankSel))
                        End Select
                    End If
                    
                    ' can always select deuces if they are wild
                    If iRankClk = 2 And Rules.DeucesWild = True Then
                        fCanSel = True
                    End If
                End If
            End If
            
            ' select, unselect or flash card
            If Deck.Mode(iCardClk) = cmSelected Then
                
                ' play click sound
                PlaySound IDSND_CARDDROP
                
                ' unselect card
                If fDeuce = True And GetRank(Deck.Card(iCardClk)) = 2 Then
                    ' if unselecting a deuce, unselect all cards
                    UnSelCards Deck
                Else
                    ' unselect all cards of same rank
                    UnSelCards Deck, iRankClk
                    Deck.Mode(iCardClk) = cmNormal
                End If
                SetStatus IDS_STATUS_CHOOSE_CARDS
            Else
                
                ' play click sound
                PlaySound IDSND_CARDCLICK
                
                If fCanSel = True Then
                    ' select card
                    If Deck.Mode(iCardClk) = cmNormal Then
                        Deck.Mode(iCardClk) = cmSelected
                    End If
                    SetStatus IDS_STATUS_CHOOSE_CARDS
                    If CountCards(Deck) = 1 Then
                        ActionClick
                    End If
                Else
                    ' flash card since it can't be selected
                    AnimFlashCard Deck, iCardClk
                End If
            End If
            
            ' redraw deck
            DrawDeck Deck
    
            ' check which rank selected, if any
            iRankSel = 0
            For iCard = 0 To nCards - 1
                If Deck.Mode(iCard) = cmSelected Then
                    iRankSel = GetRank(Deck.Card(iCard))
                    nSel = nSel + 1
                End If
            Next iCard
            
            ' enable/disable action button
            With frmMain
                If iRankSel = 0 Then
                    .cmdAction.Visible = False
                Else
                    If nSel = 1 Then
                        .cmdAction.Caption = IDS_ACTION_PLAY_CARD
                    Else
                        .cmdAction.Caption = IDS_ACTION_PLAY_CARDS
                    End If
                    .cmdAction.Visible = True
                End If
            End With
    
        Case IDD_DEALER
            If CountCards(Dealer) > 0 And Game.FirstTurn = False Then
                
                ' play click sound
                PlaySound IDSND_CARDCLICK
                
                With frmMain
                    If .cmdAction.Visible Then
                        If .cmdAction.Caption <> IDS_ACTION_TRY_DECK Then
                            .cmdAction.Caption = IDS_ACTION_TRY_DECK
                            UnSelCards User
                            DrawDeck User
                        Else
                            .cmdAction.Visible = False
                        End If
                    Else
                        .cmdAction.Caption = IDS_ACTION_TRY_DECK
                        .cmdAction.Visible = True
                    End If
                End With
            End If
        
        Case IDD_TRICK
            If Rules.GameType = IDG_FAKE Then
                sTableClick = IDS_ACTION_CALL
            Else
                sTableClick = IDS_ACTION_TAKE_CARDS
            End If
            If CountCards(Trick) > 0 Then
                
                ' play click sound
                PlaySound IDSND_CARDCLICK
                
                With frmMain
                    If .cmdAction.Visible Then
                        If .cmdAction.Caption <> sTableClick Then
                            .cmdAction.Caption = sTableClick
                            UnSelCards User
                            DrawDeck User
                        Else
                            .cmdAction.Visible = False
                        End If
                    Else
                        .cmdAction.Caption = sTableClick
                        .cmdAction.Visible = True
                    End If
                End With
            End If
        
        Case IDD_USER_ALT1, IDD_USER_ALT2, IDD_USER_ALT3
            ' play click sound
            PlaySound IDSND_CARDCLICK
            If CountCards(User) = 0 And CountCards(Deck) = GetCardLayerSpanish(User) Then
                Select Case CountCards(Deck)
                Case 2
                    If Deck.Mode(iCardClk) = cmSelected Then
                        Deck.Mode(iCardClk) = cmNormal
                    Else
                        If CheckRules(GetRank(GetTopCard(Trick)), GetRank(GetTopCard(Deck)), True) Then
                            AnimPopCards Deck, Trick, 1, False
                            With frmMain
                                .cmdAction.Caption = IDS_ACTION_PLAY_CARDS
                                .cmdAction.Visible = True
                            End With
                            
                            ' check if can continue to play cards or not
                            iRankSel = GetRank(GetTopCard(Trick))
                            If Rules.DeucesWild Then
                                Select Case iRankSel
                                Case 10
                                    ActionClick
                                Case 2
                                    If Not CardInDeck(CombineDeckSpanish(User), 2) Then
                                        ActionClick
                                    End If
                                Case Else
                                    If Not CardInDeck(CombineDeckSpanish(User), iRankSel) Then
                                        ActionClick
                                    End If
                                End Select
                            Else
                                If Not CardInDeck(CombineDeckSpanish(User), iRankSel) Then
                                    ActionClick
                                End If
                            End If
                            
                            ' actionclick if no more faceup cards on table
                            'If GetCardLayerSpanish(User) = 1 Then
                            '    ActionClick
                           ' End If
                        Else
                            AnimFlashCard Deck, 1
                        End If
                    End If
                Case 1
                    If CheckRules(GetRank(GetTopCard(Trick)), GetRank(GetTopCard(Deck)), True) Then
                        AnimPopCards Deck, Trick, 1, False
                        With frmMain
                            .cmdAction.Caption = IDS_ACTION_PLAY_CARDS
                            .cmdAction.Visible = True
                        End With
                        ActionClick
                    Else
                        AnimPopCards Deck, Trick, 1, False
                        PlaySound IDSND_KOSH
                        Delay IDT_MOVE_CARD
                        Delay IDT_MOVE_CARD
                        With frmMain
                            .cmdAction.Caption = IDS_ACTION_TAKE_CARDS
                            .cmdAction.Visible = True
                        End With
                        ActionClick
                    End If
                End Select
            Else
                SetStatus IDS_STATUS_CHOOSE_FROM_HAND
                AnimFlashCard Deck, 1
            End If
            
        End Select
    
    Case IDM_CHOOSE_3
        ' choose 3 cards to place on table in Spanish Paskahousu
        Select Case Deck.Index
        Case IDD_USER
            
            ' check what can be done to the card
            If Deck.Mode(iCardClk) = cmSelected Then
                ' play click sound
                PlaySound IDSND_CARDDROP
                ' unselect card
                Deck.Mode(iCardClk) = cmNormal
            Else
                ' play click sound
                PlaySound IDSND_CARDCLICK
                If CountSelCards(Deck) < 3 Then
                    ' select card
                    Deck.Mode(iCardClk) = cmSelected
                Else
                    ' flash card
                    AnimFlashCard Deck, iCardClk
                    SetStatus IDS_STATUS_CHOOSE_ONLY_3
                End If
            End If
            
            ' redraw deck
            DrawDeck Deck
            
            ' if 3 cards selected, allow to start game
            With frmMain
                If CountSelCards(Deck) = 3 Then
                    .cmdAction.Caption = IDS_ACTION_PLACE_CARDS
                    .cmdAction.Visible = True
                Else
                    .cmdAction.Visible = False
                End If
            End With
        End Select
'
' alternative method for selecting the 3 cards
'
'        Select Case iDeck
'        Case IDD_USER_ALT1, IDD_USER_ALT2, IDD_USER_ALT3
'            If CountCards(Deck) = 2 Then
'                PlayerInput False
'                AnimPopCards Deck, User, 1, True
'                PlayerInput True
'            End If
'        Case IDD_USER
'            If CountCards(Deck(IDD_USER_ALT1)) = 2 And CountCards(Deck(IDD_USER_ALT2)) = 2 And CountCards(Deck(IDD_USER_ALT3)) = 2 Then
'                AnimFlashCard Deck, iCardClk
'                SetStatus IDS_STATUS_CHOOSE_ONLY_3
'            Else
'                Deck.Mode(iCardClk) = cmSelected
'                PlayerInput False
'                AnimMoveSelCards Deck, Deck(GetEmptyDeckSpanish(Deck, 2))
'                PlayerInput True
'            End If
'        End Select
'
'        With frmMain
'            If CountCards(Deck(IDD_USER_ALT1)) = 2 And CountCards(Deck(IDD_USER_ALT2)) = 2 And CountCards(Deck(IDD_USER_ALT3)) = 2 Then
'                .cmdAction.Caption = IDS_ACTION_PLACE_CARDS
'                .cmdAction.Visible = True
'            Else
'                .cmdAction.Visible = False
'            End If
'        End With
    End Select
    If fAction Then
        ActionClick
    End If
End Sub
Sub DeckKeyPress(DeckIndex As Integer, Dealer As CardDeck, Trick As CardDeck, User As CardDeck, KeyAscii As Integer)
Dim iCard As Integer, iDeck As Integer
Dim nCards As Integer
    Select Case DeckIndex
    Case IDD_USER
        nCards = CountCards(Deck(DeckIndex))
        Select Case KeyAscii
        Case IDK_TRYFROMDECK
            DeckClick Dealer, Dealer, Trick, User, iCard - 1
        Case IDK_PICKUP
            DeckClick Trick, Dealer, Trick, User, iCard - 1
        Case Else
            iCard = GetKeyValue(KeyAscii)
            If iCard > 0 And iCard <= nCards Then
                DeckClick Deck(DeckIndex), Dealer, Trick, User, iCard - 1
            End If
        End Select
    Case IDD_USER_ALT1, IDD_USER_ALT2, IDD_USER_ALT3
        iDeck = GetKeyValue(KeyAscii) - 1 + IDD_USER_ALT1
        Select Case GetKeyValue(KeyAscii)
        Case 1, 2, 3
            DeckClick Deck(iDeck), Deck(IDD_DEALER), Deck(IDD_TRICK), Deck(IDD_USER), CountCards(Deck(iDeck))
        End Select
    End Select
End Sub
Sub DealCardsHook()
Dim iPlr As Integer, iCard As Integer
    If Rules.GameType = IDG_SPANISH Then
        iPlr = GetNextPlayer(Game.Dealer)
        
        For iCard = 1 To (MAX_CARDS_IN_HAND_SPANISH * MAX_PLAYERS)
            DoEvents
            AnimPopCards Deck(IDD_DEALER), Deck(GetEmptyDeckSpanish(Deck(iPlr), 1)), 1, False
            iPlr = GetNextPlayer(iPlr)
        Next iCard
    End If
End Sub
Function GetCardLayerSpanish(Player As CardDeck) As Integer
Dim iDeck As Integer, iDeckDest(2) As Integer
Dim nCards As Integer
    GetDestDeckSpanish Player, iDeckDest
    For iDeck = 0 To 2
        nCards = CountCards(Deck(iDeckDest(iDeck)))
        If nCards > GetCardLayerSpanish Then
            GetCardLayerSpanish = nCards
        End If
    Next iDeck
End Function
Function GetEmptyDeckSpanish(Player As CardDeck, ByVal Cards As Integer) As Integer
Dim iDeck As Integer, iDeckDest(2) As Integer
    GetDestDeckSpanish Player, iDeckDest
    For iDeck = 0 To 2
        If CountCards(Deck(iDeckDest(iDeck))) < Cards Then
            GetEmptyDeckSpanish = iDeckDest(iDeck)
            Exit For
        End If
    Next iDeck
End Function
Sub GetDestDeckSpanish(Player As CardDeck, iDeckDest() As Integer)
    Select Case Player.Index
    Case IDD_PLAYER1
        iDeckDest(0) = IDD_PLAYER1_ALT1
        iDeckDest(1) = IDD_PLAYER1_ALT2
        iDeckDest(2) = IDD_PLAYER1_ALT3
    Case IDD_PLAYER2
        iDeckDest(0) = IDD_PLAYER2_ALT1
        iDeckDest(1) = IDD_PLAYER2_ALT2
        iDeckDest(2) = IDD_PLAYER2_ALT3
    Case IDD_PLAYER3
        iDeckDest(0) = IDD_PLAYER3_ALT1
        iDeckDest(1) = IDD_PLAYER3_ALT2
        iDeckDest(2) = IDD_PLAYER3_ALT3
    Case IDD_PLAYER4
        iDeckDest(0) = IDD_PLAYER4_ALT1
        iDeckDest(1) = IDD_PLAYER4_ALT2
        iDeckDest(2) = IDD_PLAYER4_ALT3
    End Select
End Sub
Function GetSelRank(Deck As CardDeck, fDeucesSelected As Boolean)
Dim nCards As Integer
Dim iCard As Integer
    nCards = CountCards(Deck)
    For iCard = 0 To nCards - 1
        If Deck.Mode(iCard) = cmSelected Then
            If Rules.DeucesWild Then
                If GetRank(Deck.Card(iCard)) = 2 Then
                    fDeucesSelected = True
                Else
                    GetSelRank = GetRank(Deck.Card(iCard))
                End If
            Else
                GetSelRank = GetRank(Deck.Card(iCard))
            End If
        End If
    Next iCard
End Function

Sub PlayerInput(Enabled As Boolean)
    PlayerInputBasic Enabled
    With frmMain
        
        .cmdAction.Visible = False
        .picDeck(IDD_USER_ALT1).Enabled = Enabled
        .picDeck(IDD_USER_ALT2).Enabled = Enabled
        .picDeck(IDD_USER_ALT3).Enabled = Enabled
        
        If Enabled Then
            
            ' set focus to a deck with cards
            If CountCards(Deck(IDD_USER)) > 0 Then
                .picDeck(IDD_USER).SetFocus
            Else
                If CountCards(Deck(IDD_USER_ALT1)) > 0 Then
                    .picDeck(IDD_USER_ALT1).SetFocus
                Else
                    If CountCards(Deck(IDD_USER_ALT2)) > 0 Then
                        .picDeck(IDD_USER_ALT2).SetFocus
                    Else
                        If CountCards(Deck(IDD_USER_ALT3)) > 0 Then
                            .picDeck(IDD_USER_ALT3).SetFocus
                        End If
                    End If
                End If
            End If
            
            ' set status string
            Select Case Game.Mode
            Case IDM_NORMAL
                SetStatus IDS_STATUS_CHOOSE_CARDS
            Case IDM_CHOOSE_3
                SetStatus IDS_STATUS_CHOOSE_THREE_CARDS_FOR_TABLE
            End Select
            
        End If
    End With
End Sub
Sub ClearGameData()
Dim iPlr As Integer
    With Game
        .New = False
        .Over = False
        
        .FirstTurn = False
        .FirstCard = 3
        .NextPos = 1
        
        For iPlr = 0 To MAX_PLAYERS - 1
            .Pos(iPlr) = 0
        Next iPlr
        
        If .Demo Then
            SetStatus IDS_STATUS_DEMO, True
        Else
            SetStatus
        End If
    End With
End Sub
Sub InitGameData()
Dim iPlr As Integer
    With Game
        .Inited = True
        
        .Turn = IDD_USER
        .Dealer = GetPrevPlayer(IDD_USER)
        .FirstTurn = False
        .FirstCard = 3
        .NextPos = 1
        .RoundNbr = 1
        
        For iPlr = 0 To MAX_PLAYERS - 1
            .Score(iPlr) = 0
            .Pos(iPlr) = 0
        Next iPlr
    End With
End Sub
Function GetCardsInHand() As Integer
    Select Case Rules.GameType
    Case IDG_SPANISH
        If Game.Dealing Then
            GetCardsInHand = MAX_CARDS_IN_HAND_SPANISH * 2
        Else
            GetCardsInHand = MAX_CARDS_IN_HAND_SPANISH
        End If
    Case Else
        GetCardsInHand = MAX_CARDS_IN_HAND
    End Select
End Function
Function CheckHand(Player As CardDeck, ByVal iRank As Integer) As Boolean
Dim nCards As Integer
Dim iCard As Integer
    nCards = CountCards(Player)
    If nCards = 0 Then
        Exit Function
    End If
    For iCard = 0 To nCards - 1
        If CheckRules(iRank, GetRank(Player.Card(iCard))) Then
            CheckHand = True
            Exit For
        End If
    Next iCard
    If Not CheckHand And Game.Demo Then
        AnimHiliteDeck Player
    End If
End Function
Sub AI_PlayCards(Player As CardDeck, Dealer As CardDeck, Trick As CardDeck)
Dim iRank As Integer
Dim nCards As Integer
Dim PlayerAlt As CardDeck
Dim iDeckAlt As Integer
    
    ' end first turn
    If Game.FirstTurn Then
        Game.FirstTurn = False
    End If
    
    iRank = GetRank(GetTopCard(Trick))
    
    Select Case Rules.GameType
    Case IDG_NORMAL
        If CheckHand(Player, iRank) Then
            AI_SelectCards Player, Dealer, Trick, iRank
            AnimMoveSelCards Player, Trick, 2
            Exit Sub
        End If
    
    Case IDG_SPANISH
        Select Case Game.Mode
        Case IDM_NORMAL
            If CountCards(Player) = 0 Then
                If GetCardLayerSpanish(Player) = 1 Then
                    AI_SelectCardsRandomSpanish Player, Trick
                    Exit Sub
                Else
                    PlayerAlt = CombineDeckSpanish(Player)
                    PlayerAlt.Index = Player.Index
                    If CheckHand(PlayerAlt, iRank) Then
                        AI_SelectCards PlayerAlt, Dealer, Trick, iRank
                        AnimMoveSelCardsSpanish PlayerAlt, Trick, 2
                        Exit Sub
                    End If
                End If
            Else
                If CheckHand(Player, iRank) Then
                    AI_SelectCards Player, Dealer, Trick, iRank
                    AnimMoveSelCards Player, Trick, 2
                    Exit Sub
                End If
            End If
        
        Case IDM_CHOOSE_3
            'AI_SelectTableCardsSpanish Player
            'Exit Sub
        End Select
    End Select

    ' try from deck or pick up cards as a last resort
    If CountCards(Dealer) > 0 Then
        If CountCards(Trick) >= 4 Then
            If IsCardArray(Trick, 2, 4) Then
                ' take cards if four deuces on trick
                AnimTakeCards Trick, Player
            Else
                AnimTryFromDeck Player, Dealer, Trick
            End If
        Else
            Select Case iRank
            Case A
                If Rules.Only10s Then
                    AnimTryFromDeck Player, Dealer, Trick
                Else
                    AnimTakeCards Trick, Player
                End If
            Case 10
                ' always take 10 fed by other players
                AnimTakeCards Trick, Player
            Case Else
                AnimTryFromDeck Player, Dealer, Trick
            End Select
        End If
    Else
        AnimTakeCards Trick, Player
    End If

End Sub
Function GetNextValidPlayer(ByVal iPlr As Integer) As Integer
    Do
        iPlr = GetNextPlayer(iPlr)
    Loop Until CountCards(Deck(iPlr)) > 0 Or GetCardLayerSpanish(Deck(iPlr)) > 0
    GetNextValidPlayer = iPlr
End Function
Function GetPrevValidPlayer(ByVal iPlr As Integer) As Integer
    Do
        iPlr = GetPrevPlayer(iPlr)
    Loop Until CountCards(Deck(iPlr)) > 0 Or CountCards(CombineDeckSpanish(Deck(iPlr))) > 0
    GetPrevValidPlayer = iPlr
End Function
Sub AI_SelectCards(Player As CardDeck, Dealer As CardDeck, Trick As CardDeck, iRankBottom As Integer)
Dim iCard As Integer, iRank As Integer
Dim nCards As Integer
Dim iAI As Integer
    
    Debug.Print Player.Index + 1; , ;
    
    iAI = Game.AI(Player.Index)
    
    ' first turn, must play lowest rank
    If Game.FirstTurn Then
        SelCards Player, Game.FirstCard
        Debug.Print "aloittaa", Game.FirstCard
        Exit Sub
    End If
    
    If iAI >= 1 Then
    ' check 2
    If CountCards(Dealer) = 0 And Rules.DeucesWild Then
        If CardOnlyInDeck2(Player, 2) Then
            SelCards Player, 2
            Debug.Print "lyˆ", 2
        End If
    End If
    End If
    
    ' AI level 1 random card selection
    If iAI = 0 Then
        nCards = CountCards(Player)
        Do
            iCard = Int(nCards * Rnd)
        Loop Until CheckRules(iRankBottom, GetRank(Player.Card(iCard)))
        Player.Mode(iCard) = cmSelected
        Debug.Print "idiootti", GetRank(Player.Card(iCard))
        Exit Sub
    End If
    
    If iAI = 4 Then
    ' check if next player has only one card and feed if possible
    If CountCards(Deck(GetNextValidPlayer(Player.Index))) = 1 And CountCards(Trick) = 0 Then
        If Rules.GameType <> IDG_SPANISH Then
            If SelCards(Player, 10, 1) Then
                Debug.Print "syˆtt‰‰", 10
                Exit Sub
            End If
            If Not Rules.Only10s Then
                If SelCards(Player, A, 1) Then
                    Debug.Print "syˆtt‰‰", 1
                    Exit Sub
                End If
            End If
            If Not Rules.DeucesWild Then
                If SelCards(Player, 2, 1) Then
                    Debug.Print "syˆtt‰‰", 2
                    Exit Sub
                End If
            End If
        End If
    End If
    End If
    
    If iAI >= 3 Then
    ' check if trick can be killed
    For nCards = 1 To 3
        For iRank = 1 To K
            If IsCardArray(Trick, iRank, nCards) And iRank <> 2 And CheckRules(iRank, iRank) And CheckRules(iRankBottom, iRank) Then
                If SelCards(Player, iRank, 4 - nCards, True) = 4 - nCards Then
                    SelCards Player, iRank, 4 - nCards
                    Debug.Print "t‰ydent‰‰", iRank
                    Exit Sub
                End If
            End If
        Next iRank
    Next nCards
    End If

    If iAI >= 1 Then
    ' check for n-of-a-kind in hand
    For nCards = 4 To 1 Step -1
        For iRank = 1 To K
            If CheckRules(iRankBottom, iRank) And CheckRules(iRank, iRank) And iRank <> 2 Then
                If SelCards(Player, iRank, nCards, True) = nCards Then
                    SelCards Player, iRank, nCards
                    Debug.Print nCards & " samaa", iRank
                    Exit Sub
                End If
            End If
        Next iRank
    Next nCards
    End If
    
    ' check for 10/A/2
    If CheckRules(iRankBottom, 10) And CardInDeck(Player, 10) Then
        SelCards Player, 10, 1
        Debug.Print "kaataa", 10
        Exit Sub
    End If
    
    If Not Rules.Only10s And CheckRules(iRankBottom, A) And CardInDeck(Player, A) Then
        SelCards Player, A, 1
        Debug.Print "kaataa", 1
        Exit Sub
    End If
    
    If CardOnlyInDeck(Player, 2) Then
        SelCards Player, 2
        Exit Sub
    End If
    If CardInDeck(Player, 2) Then
        SelCards Player, 2, 1
        Debug.Print "lyˆ", 2
        If Rules.DeucesWild Then
            If CountCards(Dealer) = 0 Then
                If CardOnlyInDeck(Player, 2) Then
                    SelCards Player, 2
                    Exit Sub
                Else
                    If CardOnlyInDeck2(Player, 2) Then
                        SelCards Player, 2
                    End If
                    AI_SelectCards Player, Dealer, Trick, 0
                    Exit Sub
                End If
            Else
                AI_SelectCards Player, Dealer, Trick, 0
                Exit Sub
            End If
        Else
            If CountCards(Dealer) = 0 Then
                If CardOnlyInDeck(Player, 2) Then
                    SelCards Player, 2
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If iAI >= 1 Then
    ' check if 2 on table
    If GetRank(GetTopCard(Trick)) = 2 Then
        If Rules.DeucesWild Then
            AI_SelectCards Player, Dealer, Trick, 0
            Exit Sub
        Else
            If CardInDeck(Player, 2) Then
                If CountCards(Dealer) = 0 Then
                    If CardOnlyInDeck(Player, 2) Then
                        SelCards Player, 2
                        Exit Sub
                    End If
                Else
                    SelCards Player, 2, 1
                    Exit Sub
                End If
            End If
        End If
    End If
    End If
End Sub
Function AI_Turn() As Boolean
Dim fKill As Boolean
Dim fCPUNext As Boolean
Dim iPlrMode As Integer
    
    DoEvents
    
    ' empty table to trash if necessary
    fKill = CheckKillCards(Deck(IDD_TRICK))
    
    ' fill player hand if necessary or kick out of game
    CheckPlrCards Deck(IDD_DEALER), Deck(Game.Turn)
    
    ' find out if next player is CPU or human
    fCPUNext = IsNextPlayerCPU(Deck(Game.Turn), Deck(IDD_DEALER), Deck(IDD_TRICK), fKill)
    'iPlrMode = GetNextPlrMode(Deck(Game.Turn), Deck(IDD_DEALER), Deck(IDD_TRICK), fKill)
    
    ' check if game has or needs to be ended
    If IsGameOver(Deck(Game.Turn), Deck(IDD_DEALER), Deck(IDD_TRICK)) Then
        fCPUNext = False
        Exit Function
    End If
    
    ' play cards by computer or human
    If fCPUNext Then
        AI_PlayCards Deck(Game.Turn), Deck(IDD_DEALER), Deck(IDD_TRICK)
    Else
        PlayerInput True
    End If
    
    'Select Case iPlrMode
    'Case IDP_HUMAN
    '    PlayerInput True
    'Case IDP_CPU
    '    AI_PlayCards Deck(Game.Turn), Deck(IDD_DEALER), Deck(IDD_TRICK)
    'Case IDP_NETWORK
    '    NET_PlayCards Deck(Game.Turn),Deck(IDD_Dealer),deck(idd_trick)
    'End Select
    
    AI_Turn = fCPUNext

End Function
Function CheckRules(ByVal iRankBottom As Integer, ByVal iRankTop As Integer, Optional ByVal fUpdateStatus As Boolean) As Boolean
Dim sStatus As String
    Select Case iRankBottom
    Case 0                      ' empty table
        sStatus = IDS_STATUS_ROYALS_NOT_ON_EMPTY
        
        Select Case iRankTop
        Case A
            If Rules.HonoursOpen And Rules.Only10s Then
                CheckRules = True
            End If
            If Rules.Only10s = False Then
                CheckRules = True
            End If
        Case 10
            CheckRules = True
        Case J, Q, K
            If Rules.HonoursOpen Then
                CheckRules = True
            End If
        Case Else
            CheckRules = True
        End Select
        
    Case A
        sStatus = IDS_STATUS_ONLY_ONE_KILL
        
        If Rules.Only10s Then
            sStatus = Replace(IDS_STATUS_CHOOSE_SAME_OR_HIGHER_RANK, "%s", GetCardName(iRankBottom))
        End If
        
        Select Case iRankTop
        Case A
            If Rules.Only10s Then
                CheckRules = True
            End If
        Case 2, 10
            If Rules.Only10s Then
                CheckRules = True
            End If
        End Select
    
    Case 2
        sStatus = Replace(IDS_STATUS_CHOOSE_RANK_OR_SUIT, "%s", GetCardName(iRankBottom))
        
        Select Case iRankTop
        Case 2
            CheckRules = True
        Case Else
            If Rules.DeucesWild Then
                CheckRules = True
            End If
        End Select
        
    Case 10
        sStatus = IDS_STATUS_ONLY_ONE_KILL
        
    Case J, Q, K
        sStatus = Replace(IDS_STATUS_CHOOSE_SAME_OR_HIGHER_RANK, "%s", GetCardName(iRankBottom))
        
        Select Case iRankTop
        Case A, 2
            CheckRules = True
        Case 10
            If Rules.Only10s Then
                CheckRules = True
            Else
                sStatus = IDS_STATUS_TENS_DONT_KILL_ROYALS
            End If
        Case J, Q, K
            If iRankTop >= iRankBottom Then
                CheckRules = True
            End If
        End Select
        
    Case Else                    ' 3-9
        sStatus = Replace(IDS_STATUS_CHOOSE_SAME_OR_HIGHER_RANK, "%s", GetCardName(iRankBottom))
        
        Select Case iRankTop
        Case A
            If Rules.Only10s Then
                If iRankBottom >= Rules.HonourLimit Then
                    CheckRules = True
                Else
                    sStatus = IDS_STATUS_ROYALS_ONLY_AFTER
                End If
            Else
                sStatus = IDS_STATUS_ACES_ONLY_KILL_ROYALS
            End If
        Case 2, 10
            CheckRules = True
        Case J, Q, K
            If iRankBottom >= Rules.HonourLimit Then
                CheckRules = True
            Else
                sStatus = IDS_STATUS_ROYALS_ONLY_AFTER
            End If
        Case Else
            If iRankTop >= iRankBottom Then
                CheckRules = True
            End If
        End Select
        
    End Select
    
    ' update status bar text
    If CheckRules = False And fUpdateStatus = True Then
        SetStatus sStatus
    End If
End Function
Function IsNextPlayerCPU(Player As CardDeck, Dealer As CardDeck, Trick As CardDeck, fKill As Boolean) As Boolean
Dim iPlr As Integer
Dim iRank As Integer
Dim PlayerAlt As CardDeck
    If (CountCards(Player) > 0 Or GetCardLayerSpanish(Player) > 0) And fKill Then
        IsNextPlayerCPU = IsPlayerCPU(Player.Index)
    Else
        
        iPlr = Player.Index
        AnimSetNextTurn iPlr
        IsNextPlayerCPU = IsPlayerCPU(iPlr)
        iRank = GetRank(GetTopCard(Trick))
        
        Select Case Rules.GameType
        Case IDG_NORMAL
            ' check if computer is feeding 10/A to plr
            Select Case GetRank(GetTopCard(Trick))
            Case A
                If Not Rules.Only10s Then
                    IsNextPlayerCPU = True
                End If
            Case 10
                IsNextPlayerCPU = True
            End Select
            
            If Not CheckHand(Deck(iPlr), iRank) Then
                If CountCards(Dealer) = 0 And Not IsPlayerCPU(iPlr) Then
                    IsNextPlayerCPU = True
                End If
            End If
        
        Case IDG_SPANISH
            Select Case GetCardLayerSpanish(Deck(iPlr))
            Case 2
                If CountCards(Deck(iPlr)) = 0 Then
                    PlayerAlt = CombineDeckSpanish(Deck(iPlr))
                    PlayerAlt.Index = Deck(iPlr).Index
                    If Not CheckHand(PlayerAlt, iRank) Then
                        IsNextPlayerCPU = True
                    End If
                Else
                    If CountCards(Dealer) = 0 And Not CheckHand(Deck(iPlr), iRank) Then
                        IsNextPlayerCPU = True
                    End If
                End If
                
            Case 1
                If CountCards(Deck(iPlr)) <> 0 And Not CheckHand(Deck(iPlr), iRank) Then
                    IsNextPlayerCPU = True
                End If
            
            Case 0
                If Not CheckHand(Deck(iPlr), iRank) Then
                    If CountCards(Dealer) = 0 And Not IsPlayerCPU(iPlr) Then
                        IsNextPlayerCPU = True
                    End If
                End If
            
            End Select
        End Select
        
    End If
End Function
Sub CheckPlrCards(Dealer As CardDeck, Player As CardDeck)
Dim nCards As Integer
    Select Case Rules.GameType
    Case IDG_NORMAL
        If CountCards(Player) = 0 Or CountPlayers = 1 Then
            ' player is out of the game
            If Game.Pos(Player.Index) = 0 Then
                Game.Pos(Player.Index) = Game.NextPos
                Game.NextPos = Game.NextPos + 1
            End If
            Exit Sub
        End If
    Case IDG_SPANISH
        If (CountCards(Player) = 0 And GetCardLayerSpanish(Player) = 0 And CountCards(Dealer) = 0) Or CountPlayers = 1 Then
            ' player is out of the game
            If Game.Pos(Player.Index) = 0 Then
                Game.Pos(Player.Index) = Game.NextPos
                Game.NextPos = Game.NextPos + 1
            End If
            Exit Sub
        End If
    End Select
    
    ' add cards to player's hand
    If CountCards(Dealer) > 0 And CountCards(Player) < GetCardsInHand Then
        nCards = GetCardsInHand - CountCards(Player)
        If nCards > CountCards(Dealer) Then
            nCards = CountCards(Dealer)
        End If
        AnimPopCards Dealer, Player, nCards, True
    End If
End Sub
Function CheckKillCards(Trick As CardDeck) As Boolean
Dim iRank As Integer
Dim nCards As Integer

    iRank = GetRank(GetTopCard(Trick))
    nCards = CountCards(Trick)
    
    Select Case iRank
    Case A
        If nCards > 1 Then
            If Rules.Only10s = False Then
                CheckKillCards = True
            End If
        End If
        If nCards >= 4 Then
            If IsCardArray(Trick, iRank, 4) Then
                CheckKillCards = True
            End If
        End If
    
    Case 10
        If nCards = 1 And Rules.GameType = IDG_SPANISH Then
            CheckKillCards = True
        End If
        If nCards > 1 Then
            CheckKillCards = True
        End If
    
    Case 2
        If Rules.DeucesWild Then
            CheckKillCards = True
            Exit Function
        End If
    
    Case Else
        If nCards >= 4 Then
            If IsCardArray(Trick, iRank, 4) Then
                CheckKillCards = True
            End If
        End If
    
    End Select
    
    If CheckKillCards Then
        AnimKillCards
    End If

End Function
Sub GetFirstPlayer()
Dim iPlr As Integer, iRnd As Integer, iRank As Integer
    Select Case Game.Mode
    Case IDM_CHOOSE_3
        Game.Turn = IDD_USER
    
    Case Else
        
        iPlr = Game.Dealer
        iRank = 3
        iRnd = 0
        
        Do
            ' rotate player turn
            iPlr = GetNextPlayer(iPlr)
            
            ' if player has rank, make him start
            If CardInDeck(Deck(iPlr), iRank) = True And iPlr <> Game.Dealer Then
                Game.Turn = iPlr
                Exit Do
            End If
        
            ' rotate round
            iRnd = iRnd + 1
            
            ' if round complete, increase rank until 10 and try again
            If iRnd = MAX_PLAYERS Then
                iRank = iRank + 1
                If iRank = 10 Then
                    Game.Turn = GetNextPlayer(Game.Dealer)
                    Exit Do
                End If
                iRnd = 1
            End If
        
        Loop
        
        ' set first turn
        Game.FirstTurn = True
        Game.FirstCard = iRank
        
    End Select
    
    ' rotate one back, so correct player starts
    Game.Turn = GetPrevPlayer(Game.Turn)
    
End Sub
Function IsGameOver(Player As CardDeck, Dealer As CardDeck, Trick As CardDeck) As Boolean
Dim nCards As Integer
Dim iPlr As Integer
    Select Case CountPlayers
    Case 1
        AnimTakeCards Trick, Player
        CheckPlrCards Dealer, Player
        Game.Over = True
        IsGameOver = True
    End Select
    '
    ' lis‰tt‰v‰ :
    '
    ' - onko pelk‰st‰‰n 10/A k‰dess‰ kaikilla?
    ' - voiko mitenk‰‰n p‰‰st‰ pelist‰ pois kukaan?
    '
End Function

