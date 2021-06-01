Attribute VB_Name = "Locale_US"
' ENGLISH (US) LOCALE

' menu strings

Public Const IDS_MENU_GAME = "&Game"
Public Const IDS_MENU_GAME_NEW = "&New Game"
Public Const IDS_MENU_GAME_NETWORK = "N&etwork Game..."
Public Const IDS_MENU_GAME_SETTINGS = "&Settings..."
Public Const IDS_MENU_GAME_SCORE = "S&cores..."
Public Const IDS_MENU_GAME_SOUND = "S&ound"
Public Const IDS_MENU_GAME_DEMO = "&Demo"
Public Const IDS_MENU_GAME_EXIT = "E&xit"

Public Const IDS_MENU_HELP = "&Help"
Public Const IDS_MENU_HELP_CONTENTS = "&Contents"
Public Const IDS_MENU_HELP_ABOUT = "&About..."

' dialog strings

Public Const IDS_OK = "OK"
Public Const IDS_CANCEL = "Cancel"

Public Const IDS_DLG_SETTINGS = "Settings"
Public Const IDS_DLG_SETTINGS_TAB_GENERAL = "General"
Public Const IDS_DLG_SETTINGS_TAB_ADVANCED = "Advanced"

Public Const IDS_DLG_SETTINGS_PLAYERS = "Player Names"
Public Const IDS_DLG_SETTINGS_PLAYER = "Player"

Public Const IDS_DLG_SETTINGS_DECK = "Cards"
Public Const IDS_DLG_SETTINGS_DECK_BACK = "Choose card back picture:"

Public Const IDS_DLG_SETTINGS_RULES = "Rules"
Public Const IDS_DLG_SETTINGS_GAME = "Game:"
Public Const IDS_DLG_SETTINGS_RULE_0 = "Honours can start a Trick"
Public Const IDS_DLG_SETTINGS_RULE_1 = "Kymppi on ainoa kaatokortti"
Public Const IDS_DLG_SETTINGS_RULE_2 = "Any cards can be played on deuces"

Public Const IDS_DLG_SETTINGS_PERFORMANCE = "Performance"
Public Const IDS_DLG_SETTINGS_GAME_SPEED = "Game Speed:"
Public Const IDS_DLG_SETTINGS_ANIM_CARDS = "Animate Cards"
Public Const IDS_DLG_SETTINGS_AUTOSTART = "Automatically start new game"
Public Const IDS_DLG_SETTINGS_SHOW_SCORE = "Show scores between rounds"

' game type strings

Public Const IDS_GAME_TYPE_0 = "Paskahousu"
Public Const IDS_GAME_TYPE_1 = "Valepaskahousu"
Public Const IDS_GAME_TYPE_2 = "Espanjalainen paskahousu"
Public Const IDS_GAME_TYPE_3 = "Custom Rules"

' card back names

Public Const IDS_CARD_BACK_0 = "Cross Hatch"
Public Const IDS_CARD_BACK_1 = "Plaid"
Public Const IDS_CARD_BACK_2 = "Weave"
Public Const IDS_CARD_BACK_3 = "Robot"
Public Const IDS_CARD_BACK_4 = "Roses"
Public Const IDS_CARD_BACK_5 = "Black Ivy"
Public Const IDS_CARD_BACK_6 = "Blue Ivy"
Public Const IDS_CARD_BACK_7 = "Cyan Fishes"
Public Const IDS_CARD_BACK_8 = "Blue Fishes"
Public Const IDS_CARD_BACK_9 = "Shell"
Public Const IDS_CARD_BACK_10 = "Castle"
Public Const IDS_CARD_BACK_11 = "Beach"
Public Const IDS_CARD_BACK_12 = "Card Shark"

' action button strings

Public Const IDS_ACTION_PLAY_CARD = "Play Card"
Public Const IDS_ACTION_PLAY_CARDS = "Play Cards"
Public Const IDS_ACTION_TAKE_CARDS = "Pick Up Cards"
Public Const IDS_ACTION_TRY_DECK = "Try from Deck"
Public Const IDS_ACTION_CALL = "Call Cards"

' ui strings

Public Const IDS_COPYRIGHT = "Copyright © 2001-2002 Moonbird Software"
Public Const IDS_EMAIL = "moonbirdsoftware@hotmail.com"
Public Const IDS_URL = "http://geocities.com/moonbirdsoftware/"

Public Const IDS_CARD = "card"
Public Const IDS_CARDS = "cards"
Public Const IDS_DECK = "Deck"
Public Const IDS_NAME = "Name"
Public Const IDS_PLAYER = "Player"
Public Const IDS_SCORE = "Points"
Public Const IDS_SCOREBOARD = "Scores"
Public Const IDS_TABLE = "Trick"
Public Const IDS_TRASH = "Kaatopakka"
Public Const IDS_VERSION = "Version"
Public Const IDS_ROUND = "Round"

' game status

Public Const IDS_STATUS_TAKE = "%s is picking up cards..."
Public Const IDS_STATUS_KILL = "%s is doing the kaatamis thing..."
Public Const IDS_STATUS_TURN = "%s turn..."

Public Const IDS_STATUS_DEALING = "Dealing cards..."
Public Const IDS_STATUS_DEMO = "Demo running..."
Public Const IDS_STATUS_NEW_GAME = "Starting new game..."
Public Const IDS_STATUS_WAITING = "Waiting for players..."
Public Const IDS_STATUS_PRESS_F2 = "Press F2 to start a new game."

Public Const IDS_ERROR_CARDS32 = "Failed to initialize cards32.dll."
Public Const IDS_QUERY_FIRST_RUN = "This is the first time you have run Paskahousu. You can now check the game settings and alter them to your liking."
Public Const IDS_QUERY_RESTART_GAME = "Are you sure you want to restart the game? The game must be restarted after the rules have changed."

' card selection error strings

Public Const IDS_STATUS_CHOOSE_CARDS = "Choose cards."
Public Const IDS_STATUS_CHOOSE_LOWEST_RANK = "The lowest rank starts. Choose %s."
Public Const IDS_STATUS_CHOOSE_SAME_RANK = "Choose %s."
Public Const IDS_STATUS_CHOOSE_SAME_OR_HIGHER_RANK = "Choose %s or a higher rank."

Public Const IDS_STATUS_ACES_ONLY_KILL_ROYALS = "You cannot kaataa common cards with an Ace. Choose a Ten."
Public Const IDS_STATUS_TENS_DONT_KILL_ROYALS = "You cannot kaataa Honours with a Ten. Choose an Ace."
Public Const IDS_STATUS_ONLY_ONE_KILL = "You can only play one kaatokortti at a time."

Public Const IDS_STATUS_ROYALS_NOT_ON_EMPTY = "You cannot start a Trick with Honours. Choose a Nine or a lower rank."
Public Const IDS_STATUS_ROYALS_ONLY_AFTER = "You cannot play Honours on top of ranks lower than Seven. Choose a Nine or a lower rank."

' card names

Public Const IDS_CARD_ACE = "an Ace"
Public Const IDS_CARD_DEUX = "a Deux"
Public Const IDS_CARD_THREE = "a Three"
Public Const IDS_CARD_FOUR = "a Four"
Public Const IDS_CARD_FIVE = "a Five"
Public Const IDS_CARD_SIX = "a Six"
Public Const IDS_CARD_SEVEN = "a Seven"
Public Const IDS_CARD_EIGHT = "an Eight"
Public Const IDS_CARD_NINE = "a Nine"
Public Const IDS_CARD_TEN = "a Ten"
Public Const IDS_CARD_JACK = "a Jack"
Public Const IDS_CARD_QUEEN = "a Queen"
Public Const IDS_CARD_KING = "a King"

' debug player name - enter this as player 1's name to activate debug mode

Public Const IDS_DEBUG_PLR_NAME = "Aku Ankka"
Sub SetGameLanguage()
    Game.Language = IDL_ENGLISH
End Sub
