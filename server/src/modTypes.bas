Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public TempEventMap(1 To MAX_MAPS) As GlobalEventsRec
Public MapCache(1 To MAX_MAPS) As Cache
Public temptile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapBlocks(1 To MAX_MAPS) As MapBlockRec
Public Skill(1 To MAX_SKILLS) As SkillDataRec
Public Combo(1 To MAX_COMBOS) As ComboRec

Public Type BookRec
    Name As String * NAME_LENGTH
    Text As String * TEXT_LENGTH
End Type

'combo system
Private Type ComboRec
    Item_1 As Long
    Item_2 As Long
    Item_Given(1 To MAX_COMBO_GIVEN) As Long
    Item_Given_Val(1 To MAX_COMBO_GIVEN) As Long
    Skill As Long
    SkillLevel As Long
    Level As Long
    Take_Item1 As Byte
    Take_Item2 As Byte
    GiveSkill As Long
    GiveSkill_Exp As Long
    ReqItem1 As Long
    ReqItem2 As Long
    ReqItemVal1 As Long
    ReqItemVal2 As Long
    Take_ReqItem1 As Byte
    Take_ReqItem2 As Byte
End Type

Private Type MoveRouteRec
    index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
End Type

Private Type GlobalEventRec
    x As Long
    y As Long
    Dir As Long
    active As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    Position As Long
    
    GraphicType As Long
    GraphicNum As Long
    GraphicX As Long
    GraphicX2 As Long
    GraphicY As Long
    GraphicY2 As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
End Type

Public Type GlobalEventsRec
    EventCount As Long
    Events() As GlobalEventRec
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
    Buy_Cost As Long
    Buy_Lvl As Long
    Buy_Item As Long
    Join_Cost As Long
    Join_Lvl As Long
    Join_Item As Long
    FullScreen As Long
    Projectiles As Long
    OriginalGUIBars As Long
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    Value As Long
    Selected As Byte
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Private Type CombatRec
    Level As Byte
    EXP As Long
End Type

Private Type PlayerSpawnRec
    Map As Byte
    x As Byte
    y As Byte
End Type

Private Type SkillRec
    Level As Long
    EXP As Long
    EXP_Needed As Long
End Type

Private Type FriendRec
    NameOfFriend(1 To MAX_FRIENDS) As String * NAME_LENGTH
    Count As Long
    RequestsSent As Long
End Type

Public Type ProjectileRec
    TravelTime As Long
    Direction As Long
    x As Long
    y As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    
    ' Guilds
    GuildFileId As Long
    GuildMemberId As Long
    
    ' Admins
    Visible As Long
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    
    Spawn As PlayerSpawnRec
    ' Combat
    Combat(1 To MAX_COMBAT) As CombatRec
    
    ' K/D
    MyKills As Long
    MyDeaths As Long
    
    'Walkthrough
    WalkThrough As Byte
    
    'Kill Event
    TopKills As Long
    
    'Follow feature
    Follower As Long
    
    'Skills
    Skills(1 To MAX_SKILLS) As SkillRec
    
    'Friends
    Friends As FriendRec
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Private Type EventCommandRec
    index As Byte
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Private Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Private Type EventPageRec
    'These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    HasItemAmount As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    'End Conditions
    
    'Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    'Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    'Guidelines for the event
    WalkAnim As Long
    DirFix As Long
    WalkThrough As Long
    ShowName As Long
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    'For EventMap
    x As Long
    y As Long
End Type

Private Type EventRec
    Name As String
    Global As Byte
    PageCount As Long
    Pages() As EventPageRec
    x As Long
    y As Long
    'Self Switches re-set on restart.
    SelfSwitches(0 To 4) As Long
End Type

Public Type GlobalMapEvents
    eventID As Long
    pageID As Long
    x As Long
    y As Long
End Type

Private Type MapEventRec
    Dir As Long
    x As Long
    y As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    
    movementspeed As Long
    Position As Long
    Visible As Long
    eventID As Long
    pageID As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
    SelfSwitches(0 To 4) As Long
End Type

Private Type EventMapRec
    CurrentEvents As Long
    EventPages() As MapEventRec
End Type

Private Type EventProcessingRec
    CurList As Long
    CurSlot As Long
    eventID As Long
    pageID As Long
    WaitingForResponse As Long
    ActionTimer As Long
    ListLeftOff() As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    ' guild
    tmpGuildSlot As Long
    tmpGuildInviteSlot As Long
    tmpGuildInviteTimer As Long
    tmpGuildInviteId As Long

    EventMap As EventMapRec
    EventProcessingCount As Long
    EventProcessing() As EventProcessingRec
    
    'Projectiles
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec
    DropItemsOnDeath As Byte
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    CombatTypeReq As Byte
    CombatLvlReq As Byte
    
    Pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    Stackable As Byte
    SkillReq As Long
    
    Element_Light_Dmg As Long
    Element_Light_Res As Long
    Element_Dark_Dmg As Long
    Element_Dark_Res As Long
    Element_Neut_Dmg As Long
    Element_Neut_Res As Long
    
    Book As BookRec
    
    'Projectiles
    ProjecTile As ProjectileRec
End Type

Private Type MapItemRec
    num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
End Type

Private Type DropRec
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
    RandCurrency As Byte
    P_5 As Long
    P_10 As Long
    P_20 As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Drops(1 To MAX_NPC_DROP_ITEMS) As DropRec
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    Speed As Long
    Quest As Byte
    QuestNum As Long
    RandExp As Byte
    Percent_5 As Byte
    Percent_10 As Byte
    Percent_20 As Byte
    RandHP As Byte
    HPMin As Long
    Element_Light_Dmg As Long
    Element_Light_Res As Long
    Element_Dark_Dmg As Long
    Element_Dark_Res As Long
    Element_Neut_Dmg As Long
    Element_Neut_Res As Long
    RndSpawn As Byte
    SpawnSecsMin As Long
End Type

Private Type MapNpcRec
    num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' HP
    HPSetTo As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    CombatTypeReq As Byte
    CombatLvlReq As Byte
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    
    'NEW
    Dmg_Light As Long
    Dmg_Dark As Long
    Dmg_Neut As Long
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    NPC() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
    HPSetTo As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
    ItemRewardAmount As Long
    healthmin As Long
    HPRand As Byte
    DistItems As Byte
    ItemRewardAmountMin As Long
    ItemRewardRand As Byte
    
    'Skill exp
    Exp_Give As Byte
    Exp_Amnt As Long
    Exp_Skill As Long
    Skill_Req As Long
    Skill_LvlReq As Long
    
    'Msg Colors
    Color_Success As Long
    Color_Empty As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Public Type Vector
    x As Long
    y As Long
End Type

Public Type MapBlockRec
    Blocks() As Long
End Type

Private Type SkillDataRec
    Name As String * SKILL_LENGTH
    MaxLvl As Long
    Div As Long
End Type
