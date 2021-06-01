Attribute VB_Name = "Graphics"
Option Explicit

Sub AnimMoveCardsSpanish(Source As CardDeck, Dest As CardDeck, ByVal Mode As Integer, Optional ByVal Rank As Integer)
Dim iCard As Integer, iDeck As Integer, iDeckDest(2) As Integer
    GetDestDeckSpanish Source, iDeckDest
    For iCard = 0 To CountCards(Source) - 1
        For iDeck = 0 To 2
            If Rank = 0 Then
                If Source.Mode(iCard) = Mode And Deck(iDeckDest(iDeck)).Card(1) = Source.Card(iCard) Then
                    AnimPopCards Deck(iDeckDest(iDeck)), Dest, 1, False
                End If
            Else
                If GetRank(Source.Card(iCard)) = Rank And Source.Mode(iCard) = Mode And Deck(iDeckDest(iDeck)).Card(1) = Source.Card(iCard) Then
                    AnimPopCards Deck(iDeckDest(iDeck)), Dest, 1, False
                End If
            End If
        Next iDeck
    Next iCard
End Sub
Sub AnimMoveSelCardsSpanish(Source As CardDeck, Dest As CardDeck, Optional iRankFirst As Integer)
    If iRankFirst > 0 Then
        AnimMoveCardsSpanish Source, Dest, cmSelected, iRankFirst
    End If
    AnimMoveCardsSpanish Source, Dest, cmSelected
End Sub
Sub AnimMoveTableCardsSpanish()

End Sub
Sub FormMainResize()
Dim iDeck As Integer, iFormWidth As Integer, iFormHeight As Integer, iCardStep As Integer
Dim nCards As Integer
    
    Select Case Rules.GameType
    Case IDG_SPANISH
        With frmMain
            If .WindowState <> vbMinimized Then
                If frmMain.Width < MIN_MAIN_WIDTH_SPANISH Then
                    frmMain.Width = MIN_MAIN_WIDTH_SPANISH
                End If
                If frmMain.Height < MIN_MAIN_HEIGHT_SPANISH Then
                    frmMain.Height = MIN_MAIN_HEIGHT_SPANISH
                End If
            End If
        End With
    End Select
    
    FormMainResizeBasic
    
    ' hide or show spanish decks
    For iDeck = IDD_PLAYER1_ALT1 To IDD_PLAYER4_ALT3
        If Rules.GameType = IDG_SPANISH Then
            frmMain.picDeck(iDeck).Visible = True
        Else
            frmMain.picDeck(iDeck).Visible = False
        End If
    Next iDeck
    
    ' set dealer, trick and trash pos
    For iDeck = MAX_PLAYERS To MAX_DECKS - 1
        With frmMain.picDeck(iDeck)
            
            nCards = CountCards(Deck(iDeck))
            iCardStep = GetCardStep(Deck(iDeck), nCards)
            iFormWidth = frmMain.ScaleWidth
            iFormHeight = frmMain.ScaleHeight - frmMain.sbrStatus.Height
            
            Select Case Deck(iDeck).Index
            Case IDD_DEALER
                .Height = cdHeight + iCardStep
                .Width = cdWidth + iCardStep
                .Left = (iFormWidth - (2 * cdWidth) - 20) / 2
                .Top = (iFormHeight - frmMain.picDeck(IDD_TRICK).Height) / 2
            
            Case IDD_TRICK
                .Height = cdHeight
                .Width = cdWidth + 4 * iCardStep
                .Left = frmMain.picDeck(IDD_DEALER).Left + 94
                .Top = (iFormHeight - .Height) / 2
                
            Case IDD_TRASH
                .Height = cdHeight + iCardStep
                .Width = cdWidth + iCardStep
                Select Case Rules.GameType
                Case IDG_SPANISH
                    .Left = iFormWidth - .Width - 8
                    .Top = iFormHeight - .Height - 8
                Case Else
                    .Left = frmMain.picDeck(IDD_DEALER).Left + 206
                    .Top = (iFormHeight - frmMain.picDeck(IDD_TRICK).Height) / 2
                End Select
                
            Case IDD_PLAYER1_ALT1
                .Height = cdHeight + iCardStep
                .Width = cdWidth
                .Top = frmMain.picDeck(IDD_PLAYER1).Top - .Height - cdHeight / 8
                .Left = (iFormWidth - cdWidth * 3 - cdWidth / 4) / 2
            Case IDD_PLAYER1_ALT2
                .Height = cdHeight + iCardStep
                .Width = cdWidth
                .Top = frmMain.picDeck(IDD_PLAYER1).Top - .Height - cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER1_ALT1).Left + frmMain.picDeck(IDD_PLAYER1_ALT1).Width + cdWidth / 8
            Case IDD_PLAYER1_ALT3
                .Height = cdHeight + iCardStep
                .Width = cdWidth
                .Top = frmMain.picDeck(IDD_PLAYER1).Top - .Height - cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER1_ALT2).Left + frmMain.picDeck(IDD_PLAYER1_ALT2).Width + cdWidth / 8
            
            Case IDD_PLAYER2_ALT1
                .Height = cdHeight
                .Width = cdWidth + iCardStep
                .Top = (iFormHeight - cdHeight * 3 - cdHeight / 8) / 2
                .Left = frmMain.picDeck(IDD_PLAYER2).Left + frmMain.picDeck(IDD_PLAYER2).Width + cdWidth / 4
            Case IDD_PLAYER2_ALT2
                .Height = cdHeight
                .Width = cdWidth + iCardStep
                .Top = frmMain.picDeck(IDD_PLAYER2_ALT1).Top + frmMain.picDeck(IDD_PLAYER2_ALT1).Height + cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER2).Left + frmMain.picDeck(IDD_PLAYER2).Width + cdWidth / 4
            Case IDD_PLAYER2_ALT3
                .Height = cdHeight
                .Width = cdWidth + iCardStep
                .Top = frmMain.picDeck(IDD_PLAYER2_ALT2).Top + frmMain.picDeck(IDD_PLAYER2_ALT2).Height + cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER2).Left + frmMain.picDeck(IDD_PLAYER2).Width + cdWidth / 4
            
            Case IDD_PLAYER3_ALT1
                .Height = cdHeight + iCardStep
                .Width = cdWidth
                .Top = frmMain.picDeck(IDD_PLAYER3).Top + frmMain.picDeck(IDD_PLAYER3).Height + cdHeight / 8
                .Left = (iFormWidth - cdWidth * 3 - cdWidth / 4) / 2
            Case IDD_PLAYER3_ALT2
                .Height = cdHeight + iCardStep
                .Width = cdWidth
                .Top = frmMain.picDeck(IDD_PLAYER3).Top + frmMain.picDeck(IDD_PLAYER3).Height + cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER1_ALT1).Left + frmMain.picDeck(IDD_PLAYER1_ALT1).Width + cdWidth / 8
            Case IDD_PLAYER3_ALT3
                .Height = cdHeight + iCardStep
                .Width = cdWidth
                .Top = frmMain.picDeck(IDD_PLAYER3).Top + frmMain.picDeck(IDD_PLAYER3).Height + cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER1_ALT2).Left + frmMain.picDeck(IDD_PLAYER1_ALT2).Width + cdWidth / 8
            
            Case IDD_PLAYER4_ALT1
                .Height = cdHeight
                .Width = cdWidth + iCardStep
                .Top = (iFormHeight - cdHeight * 3 - cdHeight / 8) / 2
                .Left = frmMain.picDeck(IDD_PLAYER4).Left - frmMain.picDeck(IDD_PLAYER4_ALT1).Width - cdWidth / 4
            Case IDD_PLAYER4_ALT2
                .Height = cdHeight
                .Width = cdWidth + iCardStep
                .Top = frmMain.picDeck(IDD_PLAYER4_ALT1).Top + frmMain.picDeck(IDD_PLAYER4_ALT1).Height + cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER4).Left - frmMain.picDeck(IDD_PLAYER4_ALT1).Width - cdWidth / 4
            Case IDD_PLAYER4_ALT3
                .Height = cdHeight
                .Width = cdWidth + iCardStep
                .Top = frmMain.picDeck(IDD_PLAYER4_ALT2).Top + frmMain.picDeck(IDD_PLAYER4_ALT2).Height + cdHeight / 8
                .Left = frmMain.picDeck(IDD_PLAYER4).Left - frmMain.picDeck(IDD_PLAYER4_ALT1).Width - cdWidth / 4
            
            End Select
        End With
    Next iDeck
    
    ' set action button position
    With frmMain.cmdAction
        Select Case Rules.GameType
        Case IDG_SPANISH
            .Top = frmMain.picDeck(IDD_PLAYER1).Top + cdHeight / 6
            .Left = frmMain.picDeck(IDD_PLAYER1).Left + frmMain.picDeck(IDD_PLAYER1).Width + 15
        Case Else
            .Top = frmMain.picDeck(IDD_DEALER).Top + frmMain.picDeck(IDD_DEALER).Height + 15
            .Left = (iFormWidth - .Width) / 2
        End Select
    End With
End Sub
'Sub SetDeckPosBuffer(Deck As CardDeck)
'Dim iCardStep As Integer
'Dim nCards As Integer
'Dim shScreen As Integer, swScreen As Integer
'    nCards = CountCards(Deck)
'    iCardStep = GetCardStep(Deck, nCards)
'    shScreen = frmScreen.picScreen.ScaleHeight
'    swScreen = frmScreen.picScreen.ScaleWidth
'
'    With Deck
'        Select Case .Index
'        Case IDD_PLAYER1
'
'            .Width = cdWidth + ((nCards - 1) * iCardStep)
'            .Height = cdHeight + ch6
'
'            .x = (swScreen - .Width) / 2
'            .Y = shScreen - .Height - (ch12)
'
'        Case IDD_PLAYER2
'
'            .Width = cdWidth + ch6
'            .Height = cdHeight + ((nCards - 1) * iCardStep)
'
'            .x = ch12
'            .Y = (shScreen - .Height) / 2
'
'        Case IDD_PLAYER3
'
'            .Width = cdWidth + ((nCards - 1) * iCardStep)
'            .Height = cdHeight + ch6
'
'            .x = (swScreen - .Width) / 2
'            .Y = ch12
'
'        Case IDD_PLAYER4
'
'            .Width = cdWidth + ch6
'            .Height = cdHeight + ((nCards - 1) * iCardStep)
'
'            .x = swScreen - .Width - ch12
'            .Y = (shScreen - .Height) / 2
'
'        Case IDD_DEALER
'
'            .Width = 77
'            .Height = 103
'
'            .x = (swScreen - (2 * cdWidth) - 20) / 2
'            .Y = (shScreen - cdHeight) / 2
'
'        Case IDD_TRICK
'
'            .Width = 107
'            .Height = cdHeight
'
'            .x = (swScreen - (2 * cdWidth) - 20) / 2 + 94
'            .Y = (shScreen - cdHeight) / 2
'
'        Case IDD_TRASH
'
'            .Width = 77
'            .Height = 103
'
'            .x = (swScreen - (2 * cdWidth) - 20) / 2 + 206
'            .Y = (shScreen - cdHeight) / 2
'
'        End Select
'    End With
'End Sub
'

Sub DrawCard(Obj As Object, Deck As CardDeck, iCard As Integer, nCards As Integer, X As Integer, Y As Integer)
    DrawCardBasic Obj, Deck, iCard, nCards, X, Y
    If nCards = 0 Then
        Select Case Deck.Index
        Case IDD_DEALER, IDD_TRASH
            cdtDraw Obj.hdc, 0, 0, 0, mdDeckX, IDC_TABLEBG
        Case IDD_TRICK
            cdtDraw Obj.hdc, 0, 0, 0, mdDeckO, IDC_TABLEBG
        Case IDD_PLAYER1_ALT1, IDD_PLAYER1_ALT2, IDD_PLAYER1_ALT3, IDD_PLAYER2_ALT1, IDD_PLAYER2_ALT2, IDD_PLAYER2_ALT3, IDD_PLAYER3_ALT1, IDD_PLAYER3_ALT2, IDD_PLAYER3_ALT3, IDD_PLAYER4_ALT1, IDD_PLAYER4_ALT2, IDD_PLAYER4_ALT3
            GetCardXY Deck, iCard, nCards, X, Y
            cdtDraw Obj.hdc, X, Y, 0, mdGhost, IDC_TABLEBG
'        Case IDD_PLAYER1_ALT, IDD_PLAYER2_ALT, IDD_PLAYER3_ALT, IDD_PLAYER4_ALT
'            GetCardXY Deck, 0, nCards, X, Y
'            cdtDraw Obj.hdc, X, Y, 0, mdGhost, IDC_TABLEBG
'            GetCardXY Deck, 1, nCards, X, Y
'            cdtDraw Obj.hdc, X, Y, 0, mdGhost, IDC_TABLEBG
'            GetCardXY Deck, 2, nCards, X, Y
'            cdtDraw Obj.hdc, X, Y, 0, mdGhost, IDC_TABLEBG
        End Select
    Else
        Select Case Deck.Index
        Case IDD_DEALER, IDD_TRASH
            cdtDraw Obj.hdc, X, Y, Game.CardBack, mdFaceDown, IDC_TABLEBG
        Case IDD_TRICK
            If iCard >= (nCards - MAX_TRICK_CARDS_SHOWN) Then
                cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
            End If
        Case IDD_PLAYER1_ALT1, IDD_PLAYER1_ALT2, IDD_PLAYER1_ALT3, IDD_PLAYER2_ALT1, IDD_PLAYER2_ALT2, IDD_PLAYER2_ALT3, IDD_PLAYER3_ALT1, IDD_PLAYER3_ALT2, IDD_PLAYER3_ALT3, IDD_PLAYER4_ALT1, IDD_PLAYER4_ALT2, IDD_PLAYER4_ALT3
            GetCardXY Deck, iCard, nCards, X, Y
            If iCard = 0 Then
                cdtDraw Obj.hdc, X, Y, Game.CardBack, mdFaceDown, IDC_TABLEBG
            Else
                Select Case Deck.Mode(iCard)
                Case cmNormal, cmSelected
                    cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
                Case cmHilite
                    cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdHilite, IDC_TABLEBG
                End Select
            End If
'        Case IDD_PLAYER1_ALT, IDD_PLAYER2_ALT, IDD_PLAYER3_ALT, IDD_PLAYER4_ALT
'            GetCardXY Deck, iCard, nCards, X, Y
'            If iCard < 3 Then
'                cdtDraw Obj.hdc, X, Y, Game.CardBack, mdFaceDown, IDC_TABLEBG
'            Else
'                cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
'            End If
        End Select
    End If
End Sub
Sub AnimKillCards()
    
    SetStatus Replace(IDS_STATUS_KILL, "%s", Game.Title(Game.Turn))
    
    ' show trash
    frmMain.picDeck(IDD_TRASH).Visible = True
    Delay IDT_SHOW_CARD
    
    ' pop cards from table to trash
    AnimPopCards Deck(IDD_TRICK), Deck(IDD_TRASH), CountCards(Deck(IDD_TRICK)), False
    
    ' hide trash
    Delay IDT_SHOW_CARD
    frmMain.picDeck(IDD_TRASH).Visible = False
    
    SetStatus Replace(IDS_STATUS_TURN, "%s", MakeWordWhose(Game.Title(Game.Turn)))
    
    DoEvents

End Sub

Function AnimTryFromDeck(Player As CardDeck, Dealer As CardDeck, Trick As CardDeck) As Boolean
Dim iCard As Integer
    Debug.Print Game.Turn + 1, "yrittää pakasta"
    SetStatus Replace(IDS_STATUS_TRY, "%s", Player.Name)
    iCard = GetRank(GetTopCard(Trick))
    AnimPopCards Dealer, Trick, 1, False
    AnimTryFromDeck = CheckRules(iCard, GetRank(GetTopCard(Trick)))
    If AnimTryFromDeck = False Then
        DrawDeck Trick
        PlaySound IDSND_KOSH
        Delay IDT_MOVE_CARD
        Delay IDT_MOVE_CARD
        AnimTakeCards Trick, Player
    Else
        DrawDeck Player
    End If
End Function
Sub AnimSetNextTurn(iPlr As Integer)
Dim nCards As Integer
    ' find next player with cards left
    iPlr = GetNextValidPlayer(iPlr)
    Game.Turn = iPlr
    
    ' animate changes
    DrawTitles
    
    If Game.Turn = IDD_USER Then
        SetStatus IDS_STATUS_CHOOSE_CARDS
    Else
        SetStatus Replace(IDS_STATUS_TURN, "%s", MakeWordWhose(Game.Title(iPlr)))
    End If
    
    DoEvents
    
    Delay IDT_MOVE_CARD
End Sub
Function GetCardStep(Deck As CardDeck, ByVal nCards As Integer, Optional ByRef fStepX As Boolean, Optional ByRef fReverse As Boolean) As Integer
    GetCardStep = GetCardStepBasic(Deck, nCards, fStepX, fReverse)
    Select Case Deck.Index
    Case IDD_PLAYER1_ALT1, IDD_PLAYER1_ALT2, IDD_PLAYER1_ALT3, IDD_PLAYER3_ALT1, IDD_PLAYER3_ALT2, IDD_PLAYER3_ALT3
        GetCardStep = cdHeight / 12
    Case IDD_PLAYER2_ALT1, IDD_PLAYER2_ALT2, IDD_PLAYER2_ALT3, IDD_PLAYER4_ALT1, IDD_PLAYER4_ALT2, IDD_PLAYER4_ALT3
        GetCardStep = cdWidth / 8
    Case IDD_DEALER, IDD_TRASH
        GetCardStep = 8
    Case IDD_TRICK
        GetCardStep = cdWidth / 8
    End Select
End Function
Sub GetCardXY(Deck As CardDeck, ByVal iCard As Integer, ByVal nCards As Integer, ByRef X As Integer, ByRef Y As Integer)
Dim iStep As Integer, iPos As Integer
Dim fStepX As Boolean, fReverse As Boolean

    GetCardXYBasic Deck, iCard, nCards, X, Y

    iStep = GetCardStep(Deck, nCards, fStepX, fReverse)
    iPos = iStep * iCard
    
    Select Case Deck.Index
    Case IDD_PLAYER1_ALT1, IDD_PLAYER1_ALT2, IDD_PLAYER1_ALT3
        X = 0
        Y = iStep * (1 - iCard)
    Case IDD_PLAYER2_ALT1, IDD_PLAYER2_ALT2, IDD_PLAYER2_ALT3
        X = iPos
        Y = 0
    Case IDD_PLAYER3_ALT1, IDD_PLAYER3_ALT2, IDD_PLAYER3_ALT3
        X = 0
        Y = iPos
    Case IDD_PLAYER4_ALT1, IDD_PLAYER4_ALT2, IDD_PLAYER4_ALT3
        X = iStep * (1 - iCard)
        Y = 0
    Case IDD_DEALER, IDD_TRASH
        X = iCard / iStep
        Y = iCard / iStep
    Case IDD_TRICK
        If nCards <= MAX_TRICK_CARDS_SHOWN Then
            X = iPos
        Else
            X = (iCard - (nCards - MAX_TRICK_CARDS_SHOWN)) * iStep
        End If
        Y = 0
    End Select
    
End Sub
