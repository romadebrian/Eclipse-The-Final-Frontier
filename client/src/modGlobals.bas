Attribute VB_Name = "modGlobals"
Option Explicit
' Paperdoll rendering order
Public PaperdollOrder() As Long

'Combo editor
Public COMBO_EDITOR_ITEM_INDEX As Long

' music & sound list cache
Public musicCache() As String
Public soundCache() As String
Public hasPopulated As Boolean

' Buttons
Public LastButtonSound_Menu As Long
Public LastButtonSound_Main As Long

' Graphics
Public OldGuiBars As Boolean

' Book
Public OpeningBook As Boolean
Public Book_PageLeft As Boolean
Public Book_PageRight As Boolean

' Hotbar
Public Hotbar(1 To MAX_HOTBAR) As HotbarRec

' Amount of blood decals
Public BloodCount As Long

' main menu unloading
Public EnteringGame As Boolean

' GUI
Public BarWidth_GuiHP As Long
Public BarWidth_GuiSP As Long
Public BarWidth_GuiEXP As Long
Public BarWidth_NpcHP(1 To MAX_MAP_NPCS) As Long
Public BarWidth_PlayerHP(1 To MAX_PLAYERS) As Long

Public BarWidth_GuiHP_Max As Long
Public BarWidth_GuiSP_Max As Long
Public BarWidth_GuiEXP_Max As Long
Public BarWidth_NpcHP_Max(1 To MAX_MAP_NPCS) As Long
Public BarWidth_PlayerHP_Max(1 To MAX_PLAYERS) As Long

Public BarHeight_GuiHP As Long
Public BarHeight_GuiSP As Long
Public BarHeight_GuiEXP As Long

' NPC Chat
Public chatText As String
Public chatOpt(1 To 4) As String
Public chatOptState(1 To 4) As Byte
Public chatOnlyContinue As Boolean
Public chatContinueState As Byte

' Party GUI
Public Const Party_HPWidth As Long = 182
Public Const Party_SPRWidth As Long = 182

' targetting
Public myTarget As Long
Public myTargetType As Long

' for directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte

' trading
Public TradeTimer As Long
Public InTrade As Long
Public TradeYourOffer(1 To MAX_INV) As PlayerInvRec
Public TradeTheirOffer(1 To MAX_INV) As PlayerInvRec
Public TradeX As Long
Public TradeY As Long

' Cache the Resources in an array
Public MapResource() As MapResourceRec
Public Resource_Index As Long
Public Resources_Init As Boolean

' inv drag + drop
Public DragInvSlotNum As Long
Public InvX As Long
Public InvY As Long

' bank drag + drop
Public DragBankSlotNum As Long
Public BankX As Long
Public BankY As Long

' spell drag + drop
Public DragSpell As Long

' gui
Public EqX As Long
Public EqY As Long
Public SpellX As Long
Public SpellY As Long


' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(1 To MAX_INV) As PlayerInvRec   ' Inventory
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Long
Public InventoryItemSelected As Long
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(1 To MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Debug mode
Public DEBUG_MODE As Boolean

' TCP variables
Public PlayerBuffer As String

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Used for dragging Picture Boxes
Public SOffsetX As Long
Public SOffsetY As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS/Location needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long

' Text vars
Public vbQuote As String

' Mouse cursor tile location
Public CurX As Long
Public CurY As Long

' Game editors
Public Editor As Byte
Public EditorIndex As Long
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long
Public Scroll_Draw As Boolean
Public Scroll_Editor As Byte
Public Scroll_Timer As Integer

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public SpawnNpcNum As Long
Public SpawnNpcDir As Byte
Public EditorShop As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map Resources
Public ResourceEditorNum As Long

' Used for map editor heal & trap & slide tiles
Public MapEditorHealType As Long
Public MapEditorHealAmount As Long
Public MapEditorSlideDir As Long
Public MapEditorSound As String

' Maximum classes
Public Max_Classes As Long
Public Camera As RECT
Public TileView As RECT

' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long

' indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public AnimationIndex As Byte

' fps lock
Public FPS_Lock As Boolean

' Editor edited items array
Public Item_Changed(1 To MAX_ITEMS) As Boolean
Public NPC_Changed(1 To MAX_NPCS) As Boolean
Public Resource_Changed(1 To MAX_RESOURCES) As Boolean
Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean
Public Spell_Changed(1 To MAX_SPELLS) As Boolean
Public Shop_Changed(1 To MAX_SHOPS) As Boolean
Public Combo_Changed(1 To MAX_COMBO) As Boolean

' New char
Public newCharSprite As Long
Public newCharClass As Long
Public newCharSex As Long

' looping saves
Public Player_HighIndex As Long
Public Npc_HighIndex As Long
Public Action_HighIndex As Long

' automation problems
Public ReInitSurfaces As Boolean

' Temp event storage
Public tmpEvent As EventRec
Public isEdit As Boolean

Public curPageNum As Long
Public curCommand As Long
Public GraphicSelX As Long
Public GraphicSelY As Long
Public GraphicSelX2 As Long
Public GraphicSelY2 As Long

Public EventTileX As Long
Public EventTileY As Long

Public EditorEvent As Long

Public GraphicSelType As Long 'Are we selecting a graphic for a move route? A page sprite? What???
Public TempMoveRouteCount As Long
Public TempMoveRoute() As MoveRouteRec
Public IsMoveRouteCommand As Boolean
Public ListOfEvents() As Long

Public EventReplyID As Long
Public EventReplyPage As Long

Public RenameType As Long
Public RenameIndex As Long

Public AnotherChat As Long 'Determines if another showtext/showchoices is comming up, if so, dont close the event chatbox...

' fog
Public fogOffsetX As Long
Public fogOffsetY As Long

'Weather Stuff... events take precedent OVER map settings so we will keep temp map weather settings here.
Public CurrentWeather As Long
Public CurrentWeatherIntensity As Long
Public CurrentFog As Long
Public CurrentFogSpeed As Long
Public CurrentFogOpacity As Long
Public CurrentTintR As Long
Public CurrentTintG As Long
Public CurrentTintB As Long
Public CurrentTintA As Long
Public DrawThunder As Long

' autotiling
Public autoInner(1 To 4) As PointRec
Public autoNW(1 To 4) As PointRec
Public autoNE(1 To 4) As PointRec
Public autoSW(1 To 4) As PointRec
Public autoSE(1 To 4) As PointRec

' Map animations
Public waterfallFrame As Long
Public autoTileFrame As Long

' chat bubble
Public chatBubble(1 To MAX_BYTE) As ChatBubbleRec
Public chatBubbleIndex As Long

Public FadeType As Long
Public FadeAmount As Long
Public FlashTimer As Long

'GUI
Public InvItemFrame(1 To MAX_INV) As Byte ' Used for animated items
Public LastItemDesc As Long ' Stores the last item we showed in desc
Public LastSpellDesc As Long ' Stores the last spell we showed in desc
Public LastBankDesc As Long ' Stores the last bank item we showed in desc
Public tmpCurrencyItem As Long
Public InShop As Long ' is the player in a shop?
Public ShopAction As Byte ' stores the current shop action
Public InBank As Long
Public CurrencyMenu As Byte
Public inChat As Boolean
Public InEvent As Boolean
Public hideGUI As Boolean
Public chatOn As Boolean
Public chatShowLine As String * 1

' Game text buffer
Public MyText As String
Public RenderChatText As String
Public ChatScroll As Long
Public ChatButtonUp As Boolean
Public ChatButtonDown As Boolean
Public GuildScroll As Long
Public CombatScroll As Long
Public totalChatLines As Long

' TempStrings for rendering
Public CurrencyText As String
Public CurrencyAcceptState As Byte
Public CurrencyCloseState As Byte
Public Dialogue_ButtonVisible(1 To 3) As Boolean
Public Dialogue_ButtonState(1 To 3) As Byte
Public Dialogue_TitleCaption As String
Public Dialogue_TextCaption As String
Public TradeStatus As String
Public YourWorth As String
Public TheirWorth As String

' menu
Public sUser As String
Public sPass As String
Public sPass2 As String
Public sChar As String
Public savePass As Boolean
Public inMenu As Boolean
Public curMenu As Long
Public curTextbox As Long

' Cursor
Public GlobalX As Long
Public GlobalY As Long
Public GlobalX_Map As Long
Public GlobalY_Map As Long

' global dialogue index
Public dialogueIndex As Long
Public dialogueData1 As Long
Public sDialogue As String

' GUI consts
Public Const ChatOffsetX As Long = 6
Public Const ChatOffsetY As Long = 38
Public Const ChatWidth As Long = 380
Public TNL As Long
Public CombatTNL(1 To MAX_COMBAT) As Long

Public QuestName As String
Public QuestSay As String
Public QuestSubtitle As String
Public QuestAcceptTag As String
Public QuestAcceptState As Byte
Public QuestAcceptVisible As Boolean
Public QuestExtra As String
Public QuestExtraState As Byte
Public QuestExtraVisible As Boolean
Public QuestCloseState As Byte

'Friends System (By: escfoe2)
Public FriendRequestVisible As Boolean
Public FriendRequestSender As String
Public FriendRequestAcceptState As Byte
Public FriendRequestDeclineState As Byte
Public PlayerInfoX As Byte
Public PlayerInfoText(1 To 7) As String
Public PlayerInfoValue(1 To 6) As Long
