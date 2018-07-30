Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub SetData()
Dim tempStr As String, Path As String
    With frmServer
        Call .UsersOnline_Start
        .cboColor_Start.ListIndex = 15
        .cboColor_End.ListIndex = 15
        .cboColor_ActionMsg.ListIndex = 4
        .cboColor_PlayerMsg.ListIndex = 2
        Placements(1) = "1st"
        Placements(2) = "2nd"
        Placements(3) = "3rd"
        Placements(4) = "4th"
        Placements(5) = "5th"
        SetMainSkillData
        
        'Control Data
        Path = App.Path & "\data\options.ini"
        tempStr = GetVar(Path, "OPTIONS", "FriendSystem")
        If Len(tempStr) > 0 Then .chkFriendSystem.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "DropOnDeath")
        If Len(tempStr) > 0 Then .chkDropInvItems.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "FullScreen")
        If Len(tempStr) > 0 Then .chkFS.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "Projectiles")
        If Len(tempStr) > 0 Then .chkProj.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "OriginalGUIBars")
        If Len(tempStr) > 0 Then .chkGUIBars.Value = Val(tempStr)
    End With
End Sub

Public Sub InitServer()
    Dim I As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "guilds"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "skills"

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Mega"
        Options.Port = 7001
        Options.MOTD = "Welcome to Eclipse Mega."
        Options.Website = "http://www.touchofdeathforums.com/smf/"
        Options.Buy_Cost = 5000
        Options.Buy_Lvl = 20
        Options.Buy_Item = 1
        Options.Join_Cost = 1000
        Options.Join_Lvl = 20
        Options.Join_Item = 1
        SaveOptions
    Else
        LoadOptions
    End If
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For I = 1 To MAX_PLAYERS
        Call ClearPlayer(I)
        Load frmServer.Socket(I)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Spawning global events...")
    Call SpawnAllMapGlobalEvents
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    ' Setup Guild ranks
    Call Set_Default_Guild_Ranks

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    
    Dim Time3 As Double
    Dim Output As String
    Time3 = time2 - time1
    If Time3 > 1000 Then
        Time3 = Time3 / 1000
        Output = FormatNumber(Time3, 3)
        Call SetStatus("Initialization complete. Server loaded in " & Output & " seconds.")
    Else
        Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & " milliseconds.")
    End If
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim I As Long
    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For I = 1 To MAX_PLAYERS
        Unload frmServer.Socket(I)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing guilds...")
    Call ClearGuilds
    Call SetStatus("Clearing quests...")
    Call ClearQuests
    Call SetStatus("Clearing combos...")
    Call ClearCombos
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading switches...")
    Call LoadSwitches
    Call SetStatus("Loading variables...")
    Call LoadVariables
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Loading combos...")
    Call LoadCombos
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function



