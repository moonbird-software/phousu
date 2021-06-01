Attribute VB_Name = "Phousu"
Option Explicit

' Revision History
'
'   started in May 2001
'
'   0.9.2   First public release
'
'   0.9.3   21.12.2001
'
'   - all players' cards are revealed while in demo mode
'   - scores are not shown between rounds while in demo mode
'   - typos fixed in help file
'
'   0.9.4   21.12.2001
'
'   - frmSettings: fixed taborder
'   - frmSettings: added a tab strip
'   - when exiting from demo, the deck of the player last in turn
'     is no longer disabled for the rest of the game
'   - status message "player 1's turn" is omitted
'   - human player's turn is no longer omitted after demo mode
'     if exited from demo mode on the corresponding turn
'   - debug mode activation thru secret player name
'   - debug switches updated at startup
'
'   0.9.5   22.12.2001
'
'   - removed deck index passing thru function calls, now passed
'     thru CardDeck.Index
'   - card animations added
'
'   0.9.6   23.12.2001
'
'   - added "kill cards" status message
'   - frmSettings: added checkbox for enabling card animations
'   - frmSettings: added customizable rules
'   - frmSettings: added changeable card back
'   - card back is drawn instead of last popped card
'     when trying from deck
'
'   0.9.7   2.1.2002
'
'   - tweaked animation a bit
'   - modularized DrawDeck
'   - new game is always started after rules change
'
'   1.0.0   14.3.2002   phousu100.exe
'
'   - frmSettings: icons replaced with good ones
'
'   1.0.1   18.3.2002   phousu101.exe
'
'   - frmSettings: minor UI glitches corrected
'   - frmNetwork: created, but as a hidden menu option
'   - NetPlay: human player can be other than IDD_PLAYER1 (IDD_USER)
'   - GetFirstPlayer fixed
'   - numeric constants renamed to ID?_* scheme
'   - string constants renamed to IDS_* scheme
'   - all deuces in hand are now played if less cards than 5
'   - holddown is no longer used when human is playing
'   - decks are sorted after every dealed card
'
'   1.0.2   28.5.2002   phousu102.exe
'
'   - Rules: with normal rules, deuces are no longer wild
'   - frmSettings: new setting to disable automatic dealing at startup
'   - modularization of AI_Turn
'   - AI: if next plr has only 1 card, he is fed more by cpu
'   - AI: if cpu has 4-of-a-kind, it is always played if possible
'   - AI: cpu now tries from deck if it has no matching cards
'   - Rules: spanish rules card 10 now handled correctly
'   - frmMain: hidden mnuGameSound for obvious reasons
'   - added functionality for future implementation of "Fake" rules
'   - Help: updated
'   - Player 1's turn -> Choose cards
'
'   1.0.3   29.5.2002   phousu103.exe
'
'   - round number is shown in scores
'   - AI: if 4 deuces on table, cpu no longer tries from deck
'   - cpu card selections are now shown correctly
'   - if game in demo mode, always autostart new game
'   - animation improved a LOT
'   - last card in hand can be played with one click
'
'   1.0.4   6.6.2002    phousu104.exe
'
'   - frmSettings: players icon changed
'   - frmNetwork: icon added
'   - max_cardsinhand changed from const to variable
'   - CheckPlrCards: now checks if dealer is empty
'   - AI: 3/2oak now also played before rulebased cards
'   - Locale: "Kolme ässää" -> "Korttihai"
'   - Locale: menu and dialog strings added
'   - Locale: English (US) locale added, not finished yet
'   - "plr picks up" is shown if human tries&fails or picks up voluntarily
'   - AI: trick killing primary goal
'   - "plr tries from deck" is now shown
'   - FirstTurn is now obeyed by cpu too
'   - "Press F2 to start game" displayed
'   - settings are forced on first run time
'   - stupid settings/AI bugs fixed (like playing 2x10 on table, 4 aces not dying)
'
'   1.1.0   7.6.2002    phousu110.exe
'
'   - 4x2 doesn't die anymore
'   - dealer is correctly determined
'   - AI: playing 2's improved greatly
'   - AI: old algorithm removed
'
'   1.1.16  3.7.2002
'
'   - added initial support for single picture box graphics (runko200)
'   - all AI functions moved to AI.bas
'   - made "runko100" a common dir for risti7 and phousu
'   - basic gfx functions moved to RunkoGfx.bas
'   - AnimFlashCard function created
'   - animation delays standardized
'   - adjustable AI added
'   - animation speed is now adjusted too by game.speed
'   - frmSettings: quick dealing option added
'   - AnimObject now uses floating point variables for smoother movement
'   - cards are now sorted in order 3-9,J-K,2,10,A
'   - basic locale moved to RunkoLocFIN.bas
'   - cpu cards are drawn in a uniform fashion
'   - Windows XP style application icon added
'   - help file structure simplified
'   - gfx functions divided in basic and game specific parts
'   - trash deck moved to SE corner
'   - if deuces wild, plr can continue if gets 2 from trying
'   - deck tooltip is no longer shown if deck has no name
'
'   1.1.19  3.1.2003    TJ0!
'
'   - cleaned up AI:DealCardsDoneHook
'   - spanish table cards can now be selected with keys 1 to 3
'   - obsolete code removed
'
'   1.1.35  9.1.2003
'
'   - removed stupid spanish bugs
'   - spanish: only deuces in hand leading to crash fixed
'   - preliminary sound support
'   - when cpu feeds A/10 to human plr, card is now automatically picked up
'   - card sprites are now drawn correctly face up/down
'   - cpu cards are popped with same delay as human cards
'   - action button is now always on right side of bottom deck
'   - selected cards are now all moved before redrawing deck
'
'   1.1.38  3.4.2003
'
'   - sound support
'   - changed default font to Tahoma
'
'   1.2.0   18.10.2003
'
'   - frmSettings: ace of spades is now also drawn in card back picture
'   - drawturnicon added but disabled
'   - changed trash/action button position back to original style
'   - 10/A is now always picked up from trick if fed by other player
'
' TODO:
'
' - joskus jää ikuiseen luuppiin tietokoneiden kesken
' - valepaskahousu puuttuu
' - lisää sijoitusnäyttö
'

' RUNKO definitions
Public Const MAX_PLAYERS = 4
Public Const MAX_DECKS = 20
Public Const MAX_TRICK_CARDS_SHOWN = 4
Public Const MAX_CARDS_IN_HAND = 5
Public Const MAX_CARDS_IN_HAND_SPANISH = 3
Public Const RUNKO_APP = 0
Public Const MIN_MAIN_WIDTH = 8655
Public Const MIN_MAIN_HEIGHT = 7965
Public Const MIN_MAIN_WIDTH_SPANISH = 9825
Public Const MIN_MAIN_HEIGHT_SPANISH = 10215

' rules
Type RuleBook
    GameType As Integer
    DeucesWild As Boolean
    Only10s As Boolean
    HonourLimit As Integer
    HonoursOpen As Boolean
End Type

' deck constants
Public Const IDD_PLAYER1 = 0
Public Const IDD_PLAYER2 = 1
Public Const IDD_PLAYER3 = 2
Public Const IDD_PLAYER4 = 3
Public Const IDD_DEALER = 4
Public Const IDD_TRICK = 5
Public Const IDD_TRASH = 6
Public Const IDD_SPRITE = 7
Public Const IDD_PLAYER1_ALT1 = 8
Public Const IDD_PLAYER1_ALT2 = 9
Public Const IDD_PLAYER1_ALT3 = 10
Public Const IDD_PLAYER2_ALT1 = 11
Public Const IDD_PLAYER2_ALT2 = 12
Public Const IDD_PLAYER2_ALT3 = 13
Public Const IDD_PLAYER3_ALT1 = 14
Public Const IDD_PLAYER3_ALT2 = 15
Public Const IDD_PLAYER3_ALT3 = 16
Public Const IDD_PLAYER4_ALT1 = 17
Public Const IDD_PLAYER4_ALT2 = 18
Public Const IDD_PLAYER4_ALT3 = 19
Public Const IDD_USER = IDD_PLAYER1
Public Const IDD_USER_ALT1 = IDD_PLAYER1_ALT1
Public Const IDD_USER_ALT2 = IDD_PLAYER1_ALT2
Public Const IDD_USER_ALT3 = IDD_PLAYER1_ALT3

' game type constants
Public Const IDG_NORMAL = 0
Public Const IDG_SPANISH = 1
Public Const IDG_FAKE = 2

' game mode constants
Public Const IDM_NORMAL = 0
Public Const IDM_CHOOSE_3 = 1
Sub FormSettingsLoad()
    
    FormSettingsLoadBasic
    
    With frmSettings
        ' rules
        .cboGameType.AddItem IDS_GAME_TYPE_0
        .cboGameType.AddItem IDS_GAME_TYPE_1
        '.cboGameType.AddItem IDS_GAME_TYPE_2
        .cboGameType.ListIndex = Rules.GameType
        .chkRule(0).Value = -Rules.HonoursOpen
        .chkRule(1).Value = -Rules.Only10s
        .chkRule(2).Value = -Rules.DeucesWild
    End With
    
End Sub

Sub FormSettingsSave()
    With frmSettings
        ' prompt to start new game if rules have changed
        If Rules.HonoursOpen <> -.chkRule(0).Value Or _
            Rules.Only10s <> -.chkRule(1).Value Or _
            Rules.DeucesWild <> -.chkRule(2).Value Then
                If .cmdCancel.Enabled Then
                    If MsgBox(IDS_QUERY_RESTART_GAME, vbOKCancel + vbQuestion) = vbOK Then
                        Game.New = True
                    Else
                        Exit Sub
                    End If
                End If
        End If
        
        ' rules
        Rules.GameType = .cboGameType.ListIndex
        Rules.HonoursOpen = -.chkRule(0).Value
        Rules.Only10s = -.chkRule(1).Value
        Rules.DeucesWild = -.chkRule(2).Value
    End With
    
    FormSettingsSaveBasic
End Sub

Sub InitLocale()
    
    InitLocaleBasic
    
    ' settings
    With frmSettings
        .lblGameType.Caption = IDS_DLG_SETTINGS_GAME
        .chkRule(0).Caption = IDS_DLG_SETTINGS_RULE_0
        .chkRule(1).Caption = IDS_DLG_SETTINGS_RULE_1
        .chkRule(2).Caption = IDS_DLG_SETTINGS_RULE_2
    End With
End Sub
Sub ReadSettings()
    ' rules
    Rules.DeucesWild = GetSetting(App.Title, "Rules", "DeucesWild", False)
    Rules.GameType = GetSetting(App.Title, "Rules", "GameType", IDG_NORMAL)
    Rules.Only10s = GetSetting(App.Title, "Rules", "Only10s", False)
    Rules.HonourLimit = GetSetting(App.Title, "Rules", "HonourLimit", 7)
    Rules.HonoursOpen = GetSetting(App.Title, "Rules", "HonoursOpen", True)
    Game.CardSortOrder = csoRank2TA
    
    ReadSettingsBasic
    UpdateDebug
End Sub
Sub SaveSettings()
    SaveSettingsBasic
    
    ' rules
    SaveSetting App.Title, "Rules", "DeucesWild", Rules.DeucesWild
    SaveSetting App.Title, "Rules", "GameType", Rules.GameType
    SaveSetting App.Title, "Rules", "Only10s", Rules.Only10s
    SaveSetting App.Title, "Rules", "HonourLimit", Rules.HonourLimit
    SaveSetting App.Title, "Rules", "HonoursOpen", Rules.HonoursOpen
End Sub
Sub UpdateDebug()
    With frmMain
        .fraDebug.Visible = False
        .chkDebug(0).Value = -Rules.HonoursOpen
        .chkDebug(1).Value = -Rules.Only10s
        .chkDebug(2).Value = -Rules.DeucesWild
        .fraDebug.Visible = Game.Debug
    End With
End Sub
