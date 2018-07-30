Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Sub Main()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InitialiseGUI True
    ' set loading screen
    frmLoad.Visible = True

    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions

    ' load main menu
    Call SetStatus("Loading Menu...")
    Load frmMenu
    
    ' load gui
    Call SetStatus("Loading interface...")
    InitialiseGUI
    
    setOptionsState
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name
    
    EngineInitFontSettings
    
    InitDX8
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Initializing DirectX...")
    
    ' load music/sound engine
    InitFmod
    
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    'Load frmMainMenu
    Load frmMenu
    
    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    cacheButtons
    resetButtons_Menu
    
    ' we can now see it
    frmMenu.Visible = True
    
    ' hide all pics
    frmMenu.picCredits.Visible = False
    Show_Login False
    frmMenu.picCharacter.Visible = False
    Show_Register False
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' hide the load form
    frmLoad.Visible = False
    frmMain.Width = 12090
    frmMain.Height = 9420
    
    MenuLoop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitialiseGUI(Optional ByVal loadingScreen As Boolean = False)

'Loading Interface.ini data
Dim filename As String
Dim Path As String
filename = App.Path & "\data files\interface.ini"
Dim I As Long
    ' re-set chat scroll
    ChatScroll = 8
    GuildScroll = 1
    CombatScroll = 0
    
    ' loading screen
    If loadingScreen Then
        Path = App.Path & "\data files\graphics\gui\menu\loading.jpg"
        If FileExist(Path, True) = True Then
            Set frmLoad.Picture = LoadPicture(Path)
        End If
        Exit Sub
    End If
     ' menu
    Path = App.Path & "\data files\graphics\gui\menu\background.jpg"""
    If FileExist(Path, True) = True Then
        Set frmLoad.Picture = LoadPicture(Path)
    End If
    ReDim GUIWindow(1 To GUI_Count) As GUIWindowRec
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .x = 10 ' Val(GetVar(filename, "GUI_CHAT", "X"))
        .y = frmMain.ScaleHeight - 155 ' Val(GetVar(filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .Visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .x = Val(GetVar(filename, "GUI_HOTBAR", "X"))
        .y = Val(GetVar(filename, "GUI_HOTBAR", "Y"))
        .Height = Val(GetVar(filename, "GUI_HOTBAR", "Height"))
        .Width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .x = Val(GetVar(filename, "GUI_MENU", "X"))
        .y = Val(GetVar(filename, "GUI_MENU", "Y"))
        .Width = Val(GetVar(filename, "GUI_MENU", "Width"))
        .Height = Val(GetVar(filename, "GUI_MENU", "Height"))
        .Visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .x = Val(GetVar(filename, "GUI_BARS", "X"))
        .y = Val(GetVar(filename, "GUI_BARS", "Y"))
        '.Width = Val(GetVar(filename, "GUI_BARS", "Width"))
        '.Height = Val(GetVar(filename, "GUI_BARS", "Height"))
        .Width = 142
        .Height = 115
        .Visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .x = Val(GetVar(filename, "GUI_INVENTORY", "X"))
        .y = Val(GetVar(filename, "GUI_INVENTORY", "Y"))
        .Width = Val(GetVar(filename, "GUI_INVENTORY", "Width"))
        .Height = Val(GetVar(filename, "GUI_INVENTORY", "Height"))
        .Visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .x = Val(GetVar(filename, "GUI_SPELLS", "X"))
        .y = Val(GetVar(filename, "GUI_SPELLS", "Y"))
        .Width = Val(GetVar(filename, "GUI_SPELLS", "Width"))
        .Height = Val(GetVar(filename, "GUI_SPELLS", "Height"))
        .Visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .x = Val(GetVar(filename, "GUI_CHARACTER", "X"))
        .y = Val(GetVar(filename, "GUI_CHARACTER", "Y"))
        .Width = Val(GetVar(filename, "GUI_CHARACTER", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHARACTER", "Height"))
        .Visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .x = Val(GetVar(filename, "GUI_OPTIONS", "X"))
        .y = Val(GetVar(filename, "GUI_OPTIONS", "Y"))
        .Width = Val(GetVar(filename, "GUI_OPTIONS", "Width"))
        .Height = Val(GetVar(filename, "GUI_OPTIONS", "Height"))
        .Visible = False
    End With
    
    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .x = Val(GetVar(filename, "GUI_PARTY", "X"))
        .y = Val(GetVar(filename, "GUI_PARTY", "Y"))
        .Width = Val(GetVar(filename, "GUI_PARTY", "Width"))
        .Height = Val(GetVar(filename, "GUI_PARTY", "Height"))
        .Visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .x = Val(GetVar(filename, "GUI_DESCRIPTION", "X"))
        .y = Val(GetVar(filename, "GUI_DESCRIPTION", "Y"))
        .Width = Val(GetVar(filename, "GUI_DESCRIPTION", "Width"))
        .Height = Val(GetVar(filename, "GUI_DESCRIPTION", "Height"))
        .Visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .x = Val(GetVar(filename, "GUI_MAINMENU", "X"))
        .y = Val(GetVar(filename, "GUI_MAINMENU", "Y"))
        .Width = Val(GetVar(filename, "GUI_MAINMENU", "Width"))
        .Height = Val(GetVar(filename, "GUI_MAINMENU", "Height"))
        .Visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
        .x = Val(GetVar(filename, "GUI_SHOP", "X"))
        .y = Val(GetVar(filename, "GUI_SHOP", "Y"))
        .Width = Val(GetVar(filename, "GUI_SHOP", "Width"))
        .Height = Val(GetVar(filename, "GUI_SHOP", "Height"))
        .Visible = False
    End With
    
    ' 13 - Bank
    With GUIWindow(GUI_BANK)
        .x = 5
        .y = 62
        .Width = 480
        .Height = 384
        .Visible = False
    End With
    
    ' 14 - Trade
    With GUIWindow(GUI_TRADE)
        .x = 5
        .y = 62
        .Width = 480
        .Height = 384
        .Visible = False
    End With
    
    ' 15 - Currency
    With GUIWindow(GUI_CURRENCY)
        .x = Val(GetVar(filename, "GUI_CHAT", "X"))
        .y = Val(GetVar(filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .Visible = False
    End With
    ' 16 - Dialogue
    With GUIWindow(GUI_DIALOGUE)
        .x = Val(GetVar(filename, "GUI_CHAT", "X"))
        .y = Val(GetVar(filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .Visible = False
    End With
    ' 17 - Event Chat
    With GUIWindow(GUI_EVENTCHAT)
        .x = Val(GetVar(filename, "GUI_CHAT", "X"))
        .y = Val(GetVar(filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .Visible = False
    End With
    ' 18 - Guild Window
    With GUIWindow(GUI_GUILD)
        .x = Val(GetVar(filename, "GUI_GUILD", "X"))
        .y = Val(GetVar(filename, "GUI_GUILD", "Y"))
        .Width = Val(GetVar(filename, "GUI_GUILD", "Width"))
        .Height = Val(GetVar(filename, "GUI_GUILD", "Height"))
    End With
    ' 19 - QuestLog
    With GUIWindow(GUI_QUESTLOG)
        .x = Val(GetVar(filename, "GUI_INVENTORY", "X"))
        .y = Val(GetVar(filename, "GUI_INVENTORY", "Y"))
        .Width = Val(GetVar(filename, "GUI_INVENTORY", "Width"))
        .Height = Val(GetVar(filename, "GUI_INVENTORY", "Height"))
        .Visible = False
    End With
    ' 20 - Quest Dialogue
    With GUIWindow(GUI_QUESTDIALOGUE)
        .x = Val(GetVar(filename, "GUI_CHAT", "X"))
        .y = Val(GetVar(filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .Visible = False
    End With
    ' 21 - Combat Window
    With GUIWindow(GUI_COMBAT)
        .x = Val(GetVar(filename, "GUI_COMBAT", "X"))
        .y = Val(GetVar(filename, "GUI_COMBAT", "Y"))
        .Width = Val(GetVar(filename, "GUI_COMBAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_COMBAT", "Height"))
    End With
    ' 22 - Buddy Window
    With GUIWindow(GUI_FRIENDS)
        .x = Val(GetVar(filename, "GUI_FRIENDS", "X"))
        .y = Val(GetVar(filename, "GUI_FRIENDS", "Y"))
        .Width = Val(GetVar(filename, "GUI_FRIENDS", "Width"))
        .Height = Val(GetVar(filename, "GUI_FRIENDS", "Height"))
    End With
    ' 23 - Friend's Request Window
    With GUIWindow(GUI_FRIENDREQUEST)
        .x = Val(GetVar(filename, "GUI_FRIENDREQUEST", "X"))
        .y = Val(GetVar(filename, "GUI_FRIENDREQUEST", "Y"))
        .Width = Val(GetVar(filename, "GUI_FRIENDREQUEST", "Width"))
        .Height = Val(GetVar(filename, "GUI_FRIENDREQUEST", "Height"))
    End With
    ' 24 - Player Info Window
    With GUIWindow(GUI_PLAYERINFO)
        .x = 600
        .y = 344
        .Width = 195
        .Height = 250
    End With
    ' 25 - Book
    With GUIWindow(GUI_BOOK)
        .x = 100
        .y = 100
        .Width = 600
        .Height = 400
    End With
    
    ' BUTTONS
    With Buttons(1)
        .state = 0 ' normal
        .x = 6
        .y = 6
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 1
    End With
    
    ' main - skills
    With Buttons(2)
        .state = 0 ' normal
        .x = 81
        .y = 6
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 2
    End With
    
    ' main - char
    With Buttons(3)
        .state = 0 ' normal
        .x = 156
        .y = 6
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 3
    End With
    
    ' main - opt
    With Buttons(4)
        .state = 0 ' normal
        .x = 6
        .y = 41
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 4
    End With
    
    ' main - trade
    With Buttons(5)
        .state = 0 ' normal
        .x = 81
        .y = 41
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 5
    End With
    
    ' main - party
    With Buttons(6)
        .state = 0 ' normal
        .x = 156
        .y = 41
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 6
    End With
    
    
    
    ' menu - login
    With Buttons(7)
        .state = 0 ' normal
        .x = 172
        .y = 481
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 7
    End With
    
    ' menu - register
    With Buttons(8)
        .state = 0 ' normal
        .x = 302
        .y = 481
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 8
    End With
    
    ' menu - credits
    With Buttons(9)
        .state = 0 ' normal
        .x = 432
        .y = 481
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 9
    End With
    
    ' menu - exit
    With Buttons(10)
        .state = 0 ' normal
        .x = 562
        .y = 481
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 10
    End With
    
    ' menu - Login Accept
    With Buttons(11)
        .state = 0 ' normal
        .x = 350
        .y = 368
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 11
    End With
    
    ' menu - Register Accept
    With Buttons(12)
        .state = 0 ' normal
        .x = 350
        .y = 373
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Accept
    With Buttons(13)
        .state = 0 ' normal
        .x = 350
        .y = 445
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Next
    With Buttons(14)
        .state = 0 ' normal
        .x = 348
        .y = 445
        .Width = 89
        .Height = 29
        .Visible = True
        .PicNum = 12
    End With
    
    ' menu - NewChar Accept
    With Buttons(15)
        .state = 0 ' normal
        .x = 350
        .y = 373
        .Width = 110
        .Height = 32
        .Visible = True
        .PicNum = 11
    End With
    
    ' main - AddStats
    For I = 16 To 20
        With Buttons(I)
            .state = 0 'normal
            .Width = 12
            .Height = 11
            .Visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For I = 16 To 18 ' first 3
        With Buttons(I)
            .x = 80
            .y = 147 + ((I - 16) * 15)
        End With
    Next
    For I = 19 To 20
        With Buttons(I)
            .x = 165
            .y = 147 + ((I - 19) * 15)
        End With
    Next
    
    ' main - shop buy
    With Buttons(21)
        .state = 0 ' normal
        .x = 12
        .y = 276
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 14
    End With
    
    ' main - shop sell
    With Buttons(22)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 15
    End With
    
    ' main - shop exit
    With Buttons(23)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(24)
        .state = 0 ' normal
        .x = 14
        .y = 209
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(25)
        .state = 0 ' normal
        .x = 101
        .y = 209
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(26)
        .state = 0 ' normal
        .x = 77
        .y = 14
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(27)
        .state = 0 ' normal
        .x = 132
        .y = 14
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(28)
        .state = 0 ' normal
        .x = 77
        .y = 39
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(29)
        .state = 0 ' normal
        .x = 132
        .y = 39
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(30)
        .state = 0 ' normal
        .x = 77
        .y = 64
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(31)
        .state = 0 ' normal
        .x = 132
        .y = 64
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 20
    End With
    
    ' main - player levels on
    With Buttons(32)
        .state = 0 ' normal
        .x = 77
        .y = 89
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 19
    End With
    
    ' main - player levels off
    With Buttons(33)
        .state = 0 ' normal
        .x = 132
        .y = 89
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 20
    End With
    
    ' main - scroll up
    With Buttons(34)
        .state = 0 ' normal
        .x = 391
        .y = 2
        .Width = 19
        .Height = 19
        .Visible = True
        .PicNum = 21
    End With
    
    ' main - scroll down
    With Buttons(35)
        .state = 0 ' normal
        .x = 391
        .y = 105
        .Width = 19
        .Height = 19
        .Visible = True
        .PicNum = 22
    End With
    ' main - Select Gender Left
        With Buttons(36)
            .state = 0 'normal
            .x = 327
            .y = 318
            .Width = 19
            .Height = 19
            .Visible = True
            .PicNum = 23
        End With
        
    ' main - Select Gender Right
        With Buttons(37)
            .state = 0 'normal
            .x = 363
            .y = 318
            .Width = 19
            .Height = 19
            .Visible = True
            .PicNum = 24
        End With
    
    ' main - Select Hair Left
        With Buttons(38)
            .state = 0 'normal
            .x = 327
            .y = 345
            .Width = 19
            .Height = 19
            .Visible = True
            .PicNum = 23
        End With
        
    ' main - Select Gender Right
        With Buttons(39)
            .state = 0 'normal
            .x = 363
            .y = 345
            .Width = 19
            .Height = 19
            .Visible = True
            .PicNum = 24
        End With
    ' main - Accept Trade
        With Buttons(40)
            .state = 0 'normal
            .x = GUIWindow(GUI_TRADE).x + 165
            .y = GUIWindow(GUI_TRADE).y + 335
            .Width = 69
            .Height = 29
            .Visible = True
            .PicNum = 25
        End With
    ' main - Decline Trade
        With Buttons(41)
            .state = 0 'normal
            .x = GUIWindow(GUI_TRADE).x + 245
            .y = GUIWindow(GUI_TRADE).y + 335
            .Width = 69
            .Height = 29
            .Visible = True
            .PicNum = 26
        End With
    
    ' main - guild
    With Buttons(42)
        .state = 0 ' normal
        .x = 231
        .y = 6
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 27
    End With
    
    ' main - guild Up
    With Buttons(43)
        .state = 0 ' normal
        .x = 225
        .y = 155
        .Width = 19
        .Height = 19
        .Visible = True
        .PicNum = 21
    End With
    
    ' main - guild down
    With Buttons(44)
        .state = 0 ' normal
        .x = 225
        .y = 230
        .Width = 19
        .Height = 19
        .Visible = True
        .PicNum = 22
    End With
    
    ' main - Quest buttons
    For I = 45 To 51
        With Buttons(I)
            .state = 0 'normal
            .Width = 12
            .Height = 11
            .x = 46 + ((I - 45) * (.Width + 5))
            .y = 215
            .Visible = True
            .PicNum = 13
        End With
    Next
    
    ' main - combat Up
    With Buttons(52)
        .state = 0 ' normal
        .x = 165
        .y = 50
        .Width = 19
        .Height = 19
        .Visible = True
        .PicNum = 21
    End With
    
    ' main - combat down
    With Buttons(53)
        .state = 0 ' normal
        .x = 165
        .y = 220
        .Width = 19
        .Height = 19
        .Visible = True
        .PicNum = 22
    End With
    
    ' main - defriend
    With Buttons(54)
        .state = 0 ' normal
        .x = 117
        .y = 210
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 28
    End With
    
    ' main - message
    With Buttons(55)
        .state = 0 ' normal
        .x = 10
        .y = 210
        .Width = 69
        .Height = 29
        .Visible = True
        .PicNum = 29
    End With
    
    ' main - book left
    With Buttons(56)
        .state = 0 ' normal
        .x = 60
        .y = 295
        .Width = 30
        .Height = 25
        .Visible = True
        .PicNum = 30
    End With
    
    ' main - book right
    With Buttons(57)
        .state = 0 ' normal
        .x = 500
        .y = 300
        .Width = 30
        .Height = 25
        .Visible = True
        .PicNum = 31
    End With
    
    ' main - book close
    With Buttons(58)
        .state = 0 ' normal
        .x = 490
        .y = 50
        .Width = 19
        .Height = 13
        .Visible = True
        .PicNum = 32
    End With
    
    ' main - minimap on
    With Buttons(59)
        .state = 0 ' normal
        .x = 77
        .y = 114
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 19
    End With
    
    ' main - minimap off
    With Buttons(60)
        .state = 0 ' normal
        .x = 132
        .y = 114
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 20
    End With
    
    ' main - buttons on
    With Buttons(61)
        .state = 0 ' normal
        .x = 77
        .y = 139
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 19
    End With
    
    ' main - buttons off
    With Buttons(62)
        .state = 0 ' normal
        .x = 132
        .y = 139
        .Width = 49
        .Height = 19
        .Visible = True
        .PicNum = 20
    End With
    
    ' main - gui minimap button
    With Buttons(63)
        .state = 0 ' normal
        .x = 39 '34
        .y = 91 '86
        .Width = 18
        .Height = 18
        .Visible = True
        .PicNum = 34
    End With
    
    ' main - gui buttons button
    With Buttons(64)
        .state = 0 ' normal
        .x = 25 '20
        .y = 78 '73
        .Width = 18
        .Height = 18
        .Visible = True
        .PicNum = 33
    End With
    
    'Quest Log List
    frmMain.lstQuestLog.Width = GUIWindow(GUI_QUESTLOG).Width - 20
    frmMain.lstQuestLog.Height = GUIWindow(GUI_QUESTLOG).Height - 50
    frmMain.lstQuestLog.Left = (GUIWindow(GUI_QUESTLOG).x + (GUIWindow(GUI_QUESTLOG).Width / 2)) - (frmMain.lstQuestLog.Width / 2)
    frmMain.lstQuestLog.Top = GUIWindow(GUI_QUESTLOG).y + 10
    
    ' Buddies List
    frmMain.lstFriends.Width = GUIWindow(GUI_FRIENDS).Width - 20
    frmMain.lstFriends.Height = GUIWindow(GUI_FRIENDS).Height - 50
    frmMain.lstFriends.Left = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width / 2)) - (frmMain.lstFriends.Width / 2)
    frmMain.lstFriends.Top = GUIWindow(GUI_FRIENDS).y + 10
    
End Sub

Public Sub MenuState(ByVal state As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmLoad.Visible = True

    Select Case state
        Case MENU_STATE_ADDCHAR
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            Show_Login False
            frmMenu.picCharacter.Visible = False
            Show_Register False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")

                If frmMenu.optMale.value Then
                    Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                Else
                    Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            Show_Login False
            frmMenu.picCharacter.Visible = False
            Show_Register False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            Show_Login False
            frmMenu.picCharacter.Visible = False
            Show_Register False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
                Exit Sub
            End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picCredits.Visible = False
            Show_Login False
            frmMenu.picCharacter.Visible = False
            Show_Register False
            frmLoad.Visible = False
            Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim buffer As clsBuffer, I As Long

    isLogging = True
    InGame = False
    Set buffer = New clsBuffer
    buffer.WriteLong CQuit
    SendData buffer.ToArray()
    Set buffer = Nothing
    Call DestroyTCP
    
    ' destroy the animations loaded
    For I = 1 To MAX_BYTE
        ClearAnimInstance (I)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SpellX = 0
    SpellY = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    
    HideGame
    ' hide main form stuffs
    frmMenu.lblNews.Visible = False
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' Set font
    'Call SetFont(FONT_NAME, FONT_SIZE)
    frmMain.Font = "Arial Bold"
    frmMain.FontSize = 10
    
    ' show the main form
    frmLoad.Visible = False
    frmMain.Show
    
    ' get ping
    GetPing
    DrawPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.max = MAX_ITEMS
    frmMain.scrlAItem.value = 1
    'stop the song playing
    StopMusic
    ShowGame
    chatShowLine = "|"
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub saveGUI()
    'Loading Interface.ini data
    Dim filename As String
    filename = App.Path & "\data files\interface.ini"

    PutVar filename, "GUI_INVENTORY", "X", str(GUIWindow(GUI_INVENTORY).x)
    PutVar filename, "GUI_INVENTORY", "Y", str(GUIWindow(GUI_INVENTORY).y)
    
    PutVar filename, "GUI_SPELLS", "X", str(GUIWindow(GUI_SPELLS).x)
    PutVar filename, "GUI_SPELLS", "Y", str(GUIWindow(GUI_SPELLS).y)
    
    PutVar filename, "GUI_CHARACTER", "X", str(GUIWindow(GUI_CHARACTER).x)
    PutVar filename, "GUI_CHARACTER", "Y", str(GUIWindow(GUI_CHARACTER).y)
    
    PutVar filename, "GUI_PARTY", "X", str(GUIWindow(GUI_PARTY).x)
    PutVar filename, "GUI_PARTY", "Y", str(GUIWindow(GUI_PARTY).y)
    
    PutVar filename, "GUI_OPTIONS", "X", str(GUIWindow(GUI_OPTIONS).x)
    PutVar filename, "GUI_OPTIONS", "Y", str(GUIWindow(GUI_OPTIONS).y)
    
    PutVar filename, "GUI_GUILD", "X", str(GUIWindow(GUI_GUILD).x)
    PutVar filename, "GUI_GUILD", "Y", str(GUIWindow(GUI_GUILD).y)
    
    PutVar filename, "GUI_COMBAT", "X", str(GUIWindow(GUI_COMBAT).x)
    PutVar filename, "GUI_COMBAT", "Y", str(GUIWindow(GUI_COMBAT).y)
    
    PutVar filename, "GUI_FRIENDS", "X", str(GUIWindow(GUI_FRIENDS).x)
    PutVar filename, "GUI_FRIENDS", "Y", str(GUIWindow(GUI_FRIENDS).y)
End Sub

Public Sub DestroyGame()
Dim frm As Form
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    saveGUI
    
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    HideGame
    
    'destroy objects in reverse order
    DestroyDX8
    
    DestroyFmod

    'Call UnloadAllForms
    For Each frm In VB.Forms
        If frm.name <> "frmMenu" Then Unload frm
    Next frm
    
    Unload frmMenu
    DoEvents
    
    End
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmLoad.lblStatus.Caption = Caption
    DoEvents
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal txt As TextBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If NewLine Then
        txt.text = txt.text + Msg + vbCrLf
    Else
        txt.text = txt.text + Msg
    End If

    txt.SelStart = Len(txt.text) - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Prevent high ascii chars
    For I = 1 To Len(sInput)

        If Asc(Mid$(sInput, I, 1)) < vbKeySpace Or Asc(Mid$(sInput, I, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' menu - login
    With MenuButton(1)
        .filename = "login"
        .state = 0 ' normal
    End With
    
    ' menu - register
    With MenuButton(2)
        .filename = "register"
        .state = 0 ' normal
    End With
    
    ' menu - credits
    With MenuButton(3)
        .filename = "credits"
        .state = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(4)
        .filename = "exit"
        .state = 0 ' normal
    End With
    
    ' main - inv
    With MainButton(1)
        .filename = "inv"
        .state = 0 ' normal
    End With
    
    ' main - skills
    With MainButton(2)
        .filename = "skills"
        .state = 0 ' normal
    End With
    
    ' main - char
    With MainButton(3)
        .filename = "char"
        .state = 0 ' normal
    End With
    
    ' main - opt
    With MainButton(4)
        .filename = "opt"
        .state = 0 ' normal
    End With
    
    ' main - trade
    With MainButton(5)
        .filename = "trade"
        .state = 0 ' normal
    End With
    
    ' main - party
    With MainButton(6)
        .filename = "party"
        .state = 0 ' normal
    End With
    
    ' main - guild
    With MainButton(7)
        .filename = "guild"
        .state = 0 ' normal
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub resetClickedButtons()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' loop through entire array
    For I = 1 To MAX_BUTTONS
        Select Case I
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33, 55, 56
            Case 51, 52, 53, 54, 59, 60, 61, 62
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(I).state = 0 'normal
        End Select
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' loop through entire array
    For I = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(I).state = 0 And Not I = exceptionNum Then
            ' reset state and render
            MenuButton(I).state = 0 'normal
            renderButton_Menu I
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal buttonnum As Long)
Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' get the suffix
    Select Case MenuButton(buttonnum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(buttonnum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonnum).filename & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal buttonnum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonnum).state = bState Then Exit Sub
        ' change and render
        MenuButton(buttonnum).state = bState
        renderButton_Menu buttonnum
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub PopulateLists()
Dim strLoad As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Cache music list
    strLoad = Dir(App.Path & MUSIC_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To I) As String
        musicCache(I) = strLoad
        strLoad = Dir
        I = I + 1
    Loop
    
    ' Cache sound list
    strLoad = Dir(App.Path & SOUND_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To I) As String
        soundCache(I) = strLoad
        strLoad = Dir
        I = I + 1
    Loop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShowGame()
Dim I As Long

    For I = 1 To 4
        GUIWindow(I).Visible = True
    Next
End Sub

Public Sub HideGame()
Dim I As Long
    For I = 1 To GUI_Count - 1
        GUIWindow(I).Visible = False
    Next
    
    frmMain.lstQuestLog.Visible = False
    frmMain.lstFriends.Clear
    frmMain.lstFriends.Visible = False
End Sub

Public Sub Show_Register(ByVal YesNo As Boolean)
    frmMenu.txtRAccept.Visible = YesNo
    frmMenu.txtRPass.Visible = YesNo
    frmMenu.txtRPass2.Visible = YesNo
    frmMenu.txtRUser.Visible = YesNo
    frmMenu.lblBlank(8).Visible = YesNo
    frmMenu.lblBlank(9).Visible = YesNo
    frmMenu.lblBlank(11).Visible = YesNo
End Sub

Public Sub Show_Login(ByVal YesNo As Boolean)
    frmMenu.txtLUser.Visible = YesNo
    frmMenu.txtLPass.Visible = YesNo
    frmMenu.chkPass.Visible = YesNo
    frmMenu.lblLAccept.Visible = YesNo
    frmMenu.lblBlank(0).Visible = YesNo
    frmMenu.lblBlank(3).Visible = YesNo
    frmMenu.lblBlank(12).Visible = YesNo
End Sub
