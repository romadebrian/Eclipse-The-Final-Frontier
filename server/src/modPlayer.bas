Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long)
    If Not isPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim I As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendCombos(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendQuests(index)
    Call SendSkills(index)
    
    ' send vitals, exp + stats
    For I = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, I)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)
    
    'Do all the guild start up checks
    Call GuildLoginCheck(index)

    ' Send Resource cache
    For I = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, I
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, I As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For I = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(I).num = 0
                TempPlayer(tradeTarget).TradeOffer(I).Value = 0
                Player(index).Inv(I).Selected = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave index
        
        If Player(index).GuildFileId > 0 Then
            'Set player online flag off
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Online = False
            Call CheckUnloadGuild(TempPlayer(index).tmpGuildSlot)
        End If

        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    For I = 1 To Player_HighIndex
        If TempPlayer(I).target = index Then
            TempPlayer(I).target = 0
            TempPlayer(I).targetType = 0
            Call SendTarget(I)
        End If
    Next I
    
    Call ClearPlayer(index)
    SendUpdateFriendsLists
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If isPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim I As Long
    Dim n As Long

    If GetPlayerEquipment(index, weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            I = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= I Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim I As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            I = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= I Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim I As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If isPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
    If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    TempPlayer(index).EventProcessingCount = 0
    TempPlayer(index).EventMap.CurrentEvents = 0
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If
    
    UpdateMapBlock OldMap, GetPlayerX(index), GetPlayerY(index), False
    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    UpdateMapBlock mapnum, x, y, True
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For I = 1 To Player_HighIndex
            If isPlaying(I) Then
                If GetPlayerMap(I) = mapnum Then
                    SendMapEquipmentTo I, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For I = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).NPC(I).num > 0 Then
                MapNpc(OldMap).NPC(I).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).NPC(I).num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
    TempPlayer(index).GettingMap = YES
    Set buffer = New clsBuffer
    buffer.WriteLong SCheckForMap
    buffer.WriteLong mapnum
    buffer.WriteLong Map(mapnum).Revision
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer, mapnum As Long, I As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long, begineventprocessing As Boolean

    ' Check for subscript out of range
    If isPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = NO
    mapnum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
    
                            ' Check to see if the tile is a key and if it is, check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                                Call SetPlayerY(index, GetPlayerY(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                'Is someone following me?
                                If Player(index).Follower > 0 Then
                                    'only move follower if follower is next to me
                                    If FollowerIsNearMe(index, Player(index).Follower) Then
                                        'make the follower follow you
                                        SendPlayerFollow Player(index).Follower, GetProperDir(index, Player(index).Follower, DIR_UP)
                                    Else
                                        'then stop the follower and tell them why
                                        Call PlayerMsg(Player(index).Follower, "You must be beside the player to follow them.", Red)
                                        Player(index).Follower = 0
                                    End If
                                End If
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(mapnum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                                Call SetPlayerY(index, GetPlayerY(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                'Is someone following me?
                                If Player(index).Follower > 0 Then
                                    'only move follower if follower is next to me
                                    If FollowerIsNearMe(index, Player(index).Follower) Then
                                        'make the follower follow you
                                        SendPlayerFollow Player(index).Follower, GetProperDir(index, Player(index).Follower, DIR_DOWN)
                                    Else
                                        'then stop the follower and tell them why
                                        Call PlayerMsg(Player(index).Follower, "You must be beside the player to follow them.", Red)
                                        Player(index).Follower = 0
                                    End If
                                End If
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                                Call SetPlayerX(index, GetPlayerX(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                'Is someone following me?
                                If Player(index).Follower > 0 Then
                                    'only move follower if follower is next to me
                                    If FollowerIsNearMe(index, Player(index).Follower) Then
                                        'make the follower follow you
                                        SendPlayerFollow Player(index).Follower, GetProperDir(index, Player(index).Follower, DIR_LEFT)
                                    Else
                                        'then stop the follower and tell them why
                                        Call PlayerMsg(Player(index).Follower, "You must be beside the player to follow them.", Red)
                                        Player(index).Follower = 0
                                    End If
                                End If
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(mapnum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Or GetPlayerAccess(index) > 0 And Player(index).WalkThrough = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                                Call SetPlayerX(index, GetPlayerX(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                'Is someone following me?
                                If Player(index).Follower > 0 Then
                                    'only move follower if follower is next to me
                                    If FollowerIsNearMe(index, Player(index).Follower) Then
                                        'make the follower follow you
                                        SendPlayerFollow Player(index).Follower, GetProperDir(index, Player(index).Follower, DIR_RIGHT)
                                    Else
                                        'then stop the follower and tell them why
                                        Call PlayerMsg(Player(index).Follower, "You must be beside the player to follow them.", Red)
                                        Player(index).Follower = 0
                                    End If
                                End If
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, mapnum, x, y)
            Moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
            Call PlayerWarp(index, mapnum, x, y)
            Moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                temptile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                temptile(GetPlayerMap(index)).DoorTimer = GetTickCount
                SendMapKey index, x, y, 1
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & Amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + Amount
                PlayerMsg index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - Amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - Amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
                ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If
    
    x = GetPlayerX(index)
    y = GetPlayerY(index)
    
    If Moved = YES Then
        If TempPlayer(index).EventMap.CurrentEvents > 0 Then
            For I = 1 To TempPlayer(index).EventMap.CurrentEvents
                If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Global = 1 Then
                    If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).x = x And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).y = y And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Pages(TempPlayer(index).EventMap.EventPages(I).pageID).Trigger = 1 And TempPlayer(index).EventMap.EventPages(I).Visible = 1 Then begineventprocessing = True
                Else
                    If TempPlayer(index).EventMap.EventPages(I).x = x And TempPlayer(index).EventMap.EventPages(I).y = y And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Pages(TempPlayer(index).EventMap.EventPages(I).pageID).Trigger = 1 And TempPlayer(index).EventMap.EventPages(I).Visible = 1 Then begineventprocessing = True
                End If
                If begineventprocessing = True Then
                    'Process this event, it is on-touch and everything checks out.
                    If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Pages(TempPlayer(index).EventMap.EventPages(I).pageID).CommandListCount > 0 Then
                        TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
                        ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).ActionTimer = GetTickCount
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).CurList = 1
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).CurSlot = 1
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).eventID = TempPlayer(index).EventMap.EventPages(I).eventID
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).pageID = TempPlayer(index).EventMap.EventPages(I).pageID
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).WaitingForResponse = 0
                        ReDim TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Pages(TempPlayer(index).EventMap.EventPages(I).pageID).CommandListCount)
                    End If
                    begineventprocessing = False
                End If
            Next
        End If
    End If

End Sub

Function GetProperDir(ByVal index As Long, ByVal NeedsMoved As Long, ByVal Dir As Byte) As Byte
    'if the follower was above the player and the player moves left
    If Dir = DIR_LEFT Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) - 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) + 1 Then
                'move the follower down, not left
                GetProperDir = DIR_DOWN
                Exit Function
            End If
        End If
    End If
    
    'if the follower was above the player and the player moves right
    If Dir = DIR_RIGHT Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) - 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) - 1 Then
                'move the follower down, not right
                GetProperDir = DIR_DOWN
                Exit Function
            End If
        End If
    End If
    
    'if the follower was below the player and the player moves left
    If Dir = DIR_LEFT Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) + 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) + 1 Then
                'move the follower up, not left
                GetProperDir = DIR_UP
                Exit Function
            End If
        End If
    End If
    
    'if the follower was below the player and the player moves right
    If Dir = DIR_RIGHT Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) + 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) - 1 Then
                'move the follower up, not right
                GetProperDir = DIR_UP
                Exit Function
            End If
        End If
    End If
    
    'if the follower was to the left of the player and the player moves up
    If Dir = DIR_UP Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) + 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) - 1 Then
                'move the follower right, not up
                GetProperDir = DIR_RIGHT
                Exit Function
            End If
        End If
    End If
    
    'if the follower was to the left of the player and the player moves down
    If Dir = DIR_DOWN Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) - 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) - 1 Then
                'move the follower right, not down
                GetProperDir = DIR_RIGHT
                Exit Function
            End If
        End If
    End If
    
    'if the follower was to the right of the player and the player moves up
    If Dir = DIR_UP Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) + 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) + 1 Then
                'move the follower left, not up
                GetProperDir = DIR_LEFT
                Exit Function
            End If
        End If
    End If
    
    'if the follower was to the right of the player and the player moves down
    If Dir = DIR_DOWN Then
        If GetPlayerY(NeedsMoved) = GetPlayerY(index) - 1 Then
            If GetPlayerX(NeedsMoved) = GetPlayerX(index) + 1 Then
                'move the follower left, not down
                GetProperDir = DIR_LEFT
                Exit Function
            End If
        End If
    End If
    
    ' if walking straight keep it the same
    GetProperDir = Dir
End Function

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, movement, True
End Sub

Function FollowerIsNearMe(ByVal index As Long, Who As Long, Optional ByVal Space As Boolean = True) As Boolean
Dim XisGd As Boolean
Dim YisGd As Boolean
    If index < 0 Or index > MAX_PLAYERS Then Exit Function
    If Who < 0 Or Who > MAX_PLAYERS Then Exit Function
    
    ' check Y's
    If GetPlayerY(Who) = GetPlayerY(index) Then YisGd = True
    If GetPlayerY(Who) = GetPlayerY(index) + 1 Then YisGd = True
    If GetPlayerY(Who) = GetPlayerY(index) - 1 Then YisGd = True
    If Space Then If GetPlayerY(Who) = GetPlayerY(index) + 2 Then YisGd = True
    If Space Then If GetPlayerY(Who) = GetPlayerY(index) - 2 Then YisGd = True
    
    ' check x's
    If GetPlayerX(Who) = GetPlayerX(index) Then XisGd = True
    If GetPlayerX(Who) = GetPlayerX(index) + 1 Then XisGd = True
    If GetPlayerX(Who) = GetPlayerX(index) - 1 Then XisGd = True
    If Space Then If GetPlayerX(Who) = GetPlayerX(index) + 2 Then XisGd = True
    If Space Then If GetPlayerX(Who) = GetPlayerX(index) - 2 Then XisGd = True
    
    ' do both check out?
    If YisGd = True And XisGd = True Then FollowerIsNearMe = True
End Function

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim I As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For I = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(index, I)

        If itemnum > 0 Then

            Select Case I
                Case Equipment.weapon

                    If Item(itemnum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, I
                Case Equipment.Armor

                    If Item(itemnum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, I
                Case Equipment.Helmet

                    If Item(itemnum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, I
                Case Equipment.Shield

                    If Item(itemnum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, I
            End Select

        Else
            SetPlayerEquipment index, 0, I
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range
    If isPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For I = 1 To MAX_INV

            If GetPlayerInvItemNum(index, I) = itemnum Then
                FindOpenInvSlot = I
                Exit Function
            End If

        Next

    End If

    For I = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, I) = 0 Then
            FindOpenInvSlot = I
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim I As Long

    If Not isPlaying(index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For I = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, I) = itemnum Then
                FindOpenBankSlot = I
                Exit Function
            End If
        Next I

    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, I) = 0 Then
            FindOpenBankSlot = I
            Exit Function
        End If
    Next I

End Function

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range
    If isPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, I) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
                HasItem = GetPlayerInvItemValue(index, I)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function HasItems(ByVal index As Long, ByVal itemnum As Long, Amnt As Long) As Boolean
    Dim I As Long
    Dim cnt As Long

    ' Check for subscript out of range
    If isPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, I) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
                If GetPlayerInvItemValue(index, I) >= Amnt Then HasItems = True
                Exit Function
            Else
                cnt = cnt + 1
            End If
        End If

    Next
    
    If cnt >= Amnt Then HasItems = True

End Function

Function FindItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range
    If isPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, I) = itemnum Then
            FindItem = I
            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim I As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If isPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, I) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, I) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) - ItemVal)
                    Call SendInventoryUpdate(index, I)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, I, 0)
                Call SetPlayerInvItemValue(index, I, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(index, I)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim I As Long
    Dim n As Long
    Dim itemnum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If isPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemnum = GetPlayerInvItemNum(index, invSlot)

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendupdate As Boolean = True) As Boolean
    Dim I As Long

    ' Check for subscript out of range
    If isPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    I = FindOpenInvSlot(index, itemnum)

    ' Check to see if inventory is full
    If I <> 0 Then
        Call SetPlayerInvItemNum(index, I, itemnum)
        Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) + ItemVal)
        If sendupdate Then Call SendInventoryUpdate(index, I)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal spellnum As Long) As Boolean
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, I) = spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, I) = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim I As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not isPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)

    For I = MAX_MAP_ITEMS To 1 Step -1
        ' See if theres even an item here
        If (MapItem(mapnum, I).num > 0) And (MapItem(mapnum, I).num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, I) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, I).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, I).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, I).num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, I).num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, n)).Stackable > 0 Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, I).Value)
                                Msg = MapItem(mapnum, I).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem I, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(I, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Call CheckTasks(index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(Item(GetPlayerInvItemNum(index, n)).Name)))
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim I As Long

    ' Check for subscript out of range
    If isPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check and make sure the player isn't doing anything
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            I = FindOpenMapItemSlot(GetPlayerMap(index))

            If I <> 0 Then
                MapItem(GetPlayerMap(index), I).num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), I).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), I).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), I).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), I).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), I).canDespawn = True
                MapItem(GetPlayerMap(index), I).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, invNum)).Stackable > 0 Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), I).Value = GetPlayerInvItemValue(index, invNum)
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), I).Value = Amount
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), I).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(I, MapItem(GetPlayerMap(index), I).num, Amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), I).canDespawn)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim I As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP index
        SendPlayerData index
    End If
End Sub
 
Sub CheckCombatLevelUp(ByVal index As Long, ByVal skillType As Byte)
    Dim I As Long
    Dim expRollover As Long
    Dim level_count As Byte
    
    level_count = 0
    
    Do While GetPlayerCombatExp(index, skillType) >= GetPlayerNextCombatLevel(index, skillType)
        expRollover = GetPlayerCombatExp(index, skillType) - GetPlayerNextCombatLevel(index, skillType)
        
        If Not SetPlayerCombatLevel(index, GetPlayerCombatLevel(index, skillType) + 1, skillType) Then
            Exit Sub
        End If
                
        Call SetPlayerCombatExp(index, skillType, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        SendCombatEXP index
    End If

End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerFriends(ByVal index As Long) As Long
    GetPlayerFriends = 0
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerFriends = Player(index).Friends.Count
End Function

Function GetPlayerFriendRequests(ByVal index As Long) As Long
    GetPlayerFriendRequests = 0
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerFriendRequests = Player(index).Friends.RequestsSent
End Function

Sub SetPlayerFriendRequests(ByVal index As Long, ByVal rNum As Long, Optional ByVal PlusVal As Boolean = True)
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    If rNum = 0 Then Exit Sub
    
    If PlusVal Then
        Player(index).Friends.RequestsSent = Player(index).Friends.RequestsSent + rNum
    Else
        Player(index).Friends.RequestsSent = rNum
    End If
End Sub

Sub SetPlayerFriends(ByVal index As Long, ByVal fNum As Long, Optional ByVal PlusVal As Boolean = True)
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    
    If PlusVal Then
        If fNum = 0 Then Exit Sub
        Player(index).Friends.Count = Player(index).Friends.Count + fNum
    Else
        Player(index).Friends.Count = fNum
    End If
End Sub

Sub SetPlayerFriendName(ByVal index As Long, ByVal fNum As Long, Optional ByVal fName As String = vbNullString)
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    If fNum = 0 Then fNum = 1
    Player(index).Friends.NameOfFriend(fNum) = fName
End Sub

Function GetPlayerFriendName(ByVal index As Long, ByVal fNum As Long)
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerFriendName = Trim$(Player(index).Friends.NameOfFriend(fNum))
End Function

Function GetPlayerLevel(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).EXP
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal EXP As Long)
    Player(index).EXP = EXP
    If GetPlayerLevel(index) = MAX_LEVELS And Player(index).EXP > GetPlayerNextLevel(index) Then
        Player(index).EXP = GetPlayerNextLevel(index)
    End If
End Sub

Function GetPlayerCombatLevel(ByVal index As Long, ByVal skillType As Byte) As Long

    If index > MAX_PLAYERS Then Exit Function
    If skillType < 1 Then Exit Function
    GetPlayerCombatLevel = Player(index).Combat(skillType).Level
End Function

Function SetPlayerCombatLevel(ByVal index As Long, ByVal Level As Long, ByVal skillType As Byte) As Boolean
    SetPlayerCombatLevel = False
    If Level > MAX_COMBAT_LEVEL Then Exit Function
    Player(index).Combat(skillType).Level = Level
    SetPlayerCombatLevel = True
End Function

Function GetPlayerNextCombatLevel(ByVal index As Long, ByVal skillType As Byte) As Long
    GetPlayerNextCombatLevel = (50 / 3) * ((GetPlayerCombatLevel(index, skillType) + 1) ^ 3 - (6 * (GetPlayerCombatLevel(index, skillType) + 1) ^ 2) + 17 * (GetPlayerCombatLevel(index, skillType) + 1) - 12)
End Function

Function GetPlayerCombatExp(ByVal index As Long, ByVal skillType As Byte) As Long
    GetPlayerCombatExp = Player(index).Combat(skillType).EXP
End Function

Sub SetPlayerCombatExp(ByVal index As Long, ByVal skillType As Byte, ByVal EXP As Long)
    Player(index).Combat(skillType).EXP = EXP
    If GetPlayerCombatLevel(index, skillType) = MAX_LEVELS And Player(index).Combat(skillType).EXP > GetPlayerNextCombatLevel(index, skillType) Then
        Player(index).Combat(skillType).EXP = GetPlayerNextCombatLevel(index, skillType)
    End If
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVisible(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerVisible = Player(index).Visible
End Function

Sub SetPlayerVisible(ByVal index As Long, ByVal Visible As Long)
    Player(index).Visible = Visible
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Sub SetPlayerKills(ByVal index As Long, ByVal Value As Long, Optional ByVal PlusVal As Boolean = True)
    If PlusVal Then
        Player(index).MyKills = Player(index).MyKills + Value
    Else
        Player(index).MyKills = Value
    End If
End Sub

Function GetPlayerKills(ByVal index As Long) As Long
    GetPlayerKills = Player(index).MyKills
End Function

Sub SetPlayerDeaths(ByVal index As Long, ByVal Value As Long, Optional PlusVal As Boolean = True)
    If PlusVal Then
        Player(index).MyDeaths = Player(index).MyDeaths + Value
    Else
        Player(index).MyDeaths = Value
    End If
End Sub

Function GetPlayerDeaths(ByVal index As Long) As Long
    GetPlayerDeaths = Player(index).MyDeaths
End Function

Public Function GetPlayerStat(ByVal index As Long, ByVal stat As Stats) As Long
    Dim x As Long, I As Long
    If index > MAX_PLAYERS Then Exit Function
    
    x = Player(index).stat(stat)
    
    For I = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(I) > 0 Then
            If Item(Player(index).Equipment(I)).Add_Stat(stat) > 0 Then
                x = x + Item(Player(index).Equipment(I)).Add_Stat(stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).stat(stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal stat As Stats, ByVal Value As Long)
    Player(index).stat(stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long

    If index < 1 Or index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Sub SetPlayerSpawn(ByVal index As Long, ByVal Map As Byte, ByVal x As Byte, ByVal y As Byte)
    Player(index).Spawn.Map = Map
    Player(index).Spawn.x = x
    Player(index).Spawn.y = y
End Sub

Function GetPlayerSpawnMap(ByVal index As Long) As Byte

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpawnMap = Player(index).Spawn.Map
End Function

Function GetPlayerSpawnX(ByVal index As Long) As Byte

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpawnX = Player(index).Spawn.x
End Function

Function GetPlayerSpawnY(ByVal index As Long) As Byte

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpawnY = Player(index).Spawn.y
End Function

Function GetPlayerDir(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    Player(index).Inv(invSlot).num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal spellnum As Long)
    Player(index).Spell(spellslot) = spellnum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
    If GetPlayerEquipment > 1000000 Then GetPlayerEquipment = 0
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim I As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)
    
    ' Set Deaths
    Call SetPlayerDeaths(index, 1)

    'Drop inventory items if allowed
    If frmServer.chkDropInvItems.Value = vbChecked Then
        If Map(GetPlayerMap(index)).DropItemsOnDeath = 1 Then
            For I = 1 To MAX_INV
                PlayerMapDropItem index, I, GetPlayerInvItemValue(index, I)
            Next
    
    
            'Send all equiped items to the inventory to be dumped.
            For I = 1 To Equipment.Equipment_Count - 1
                If GetPlayerEquipment(index, I) > 0 Then
                    PlayerMapDropItem index, GetPlayerEquipment(index, I), 0
                End If
               
                'Send Weapon
                GiveInvItem index, GetPlayerEquipment(index, weapon), 0
                SetPlayerEquipment index, 0, weapon
                'Send Armor
                GiveInvItem index, GetPlayerEquipment(index, Armor), 0
                SetPlayerEquipment index, 0, Armor
                'Send Shield
                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                SetPlayerEquipment index, 0, Shield
                'Send Helmet
                GiveInvItem index, GetPlayerEquipment(index, Helmet), 0
                SetPlayerEquipment index, 0, Helmet
            Next
        
            'Drop *equipped* inventory items
            For I = 1 To MAX_INV
                PlayerMapDropItem index, I, 0
            Next
        End If
    End If

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, GetPlayerSpawnMap(index), GetPlayerSpawnX(index), GetPlayerSpawnY(index))
        End If
    End With
    
    ' clear all DoTs and HoTs
    For I = 1 To MAX_DOTS
        With TempPlayer(index).DoT(I)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(I)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim I As Long, II As Long
    Dim Damage As Long
    Dim Divisor As Long
    Static LastDiv As Long
    Dim RandItemAmnt As Long
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1

        ' Get the cache number
        For I = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(I).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(I).y = y Then
                    Resource_num = I
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(index, weapon) > 0 Then
                If Item(GetPlayerEquipment(index, weapon)).Data3 = Resource(Resource_index).ToolRequired Then
                    
                    'Make sure player has the skill level required
                    If Resource(Resource_index).Skill_Req > 0 Then
                        If Resource(Resource_index).Skill_LvlReq > 0 Then
                            If Resource(Resource_index).Skill_Req > 0 Then
                                If GetPlayerSkillLevel(index, Resource(Resource_index).Skill_Req) < Resource(Resource_index).Skill_LvlReq Then
                                    PlayerMsg index, "You do not have the skill level to interact with this resource.", BrightRed
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                        ' inv space?
                        If Resource(Resource_index).ItemReward > 0 Then
                            If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                                PlayerMsg index, "You have no inventory space.", BrightRed
                                Exit Sub
                            End If
                        End If
    
                        ' check if already cut down
                        If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                        
                            rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                            rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                            
                            Damage = Item(GetPlayerEquipment(index, weapon)).Data2
                        
                            ' check if damage is more than health
                            If Damage > 0 Then
                                ' check for items to give
                                If Resource(Resource_index).ItemRewardRand = 1 Then
                                    RandItemAmnt = rand(Resource(Resource_index).ItemRewardAmountMin, Resource(Resource_index).ItemRewardAmount)
                                Else
                                    RandItemAmnt = Resource(Resource_index).ItemRewardAmount
                                End If
                            
                                ' cut it down!
                                If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                    SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                    SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                    ' send message if it exists
                                    If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), Resource(Resource_index).Color_Success, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                    End If
                                    ' carry on/give item amount specified
                                    If Resource(Resource_index).DistItems = 0 Then
                                        For II = 1 To RandItemAmnt
                                            Call GiveInvItem(index, Resource(Resource_index).ItemReward, 1)
                                        Next II
                                    Else
                                        GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                    End If
                                    
                                    ' Give skill experience
                                    If Resource(Resource_index).Exp_Give = 1 Then
                                        If Resource(Resource_index).Exp_Skill > 0 Then
                                            If Resource(Resource_index).Exp_Amnt > 0 Then
                                                Call SetPlayerSkillExp(index, Resource(Resource_index).Exp_Skill, Resource(Resource_index).Exp_Amnt)
                                                Call SendPlayerData(index)
                                            End If
                                        End If
                                    End If
                                    SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                                    
                                    'clear the tracker shit
                                    LastDiv = 0
                                Else
                                    ' just do the damage
                                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                                    SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                    SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                                    
                                    ' are we giving items during attack?
                                    If Resource(Resource_index).DistItems = 1 Then
                                        If RandItemAmnt > 1 Then
                                            If Item(Resource(Resource_index).ItemReward).Type <> ITEM_TYPE_CURRENCY Then
                                                ' divide damage and give items appropriately
                                                For II = RandItemAmnt - 1 To 1 Step -1
                                                    ' run through divisors to give items
                                                    If LastDiv = 0 Then LastDiv = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).HPSetTo
                                                    Divisor = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).HPSetTo / RandItemAmnt
                                                    Divisor = Divisor * II
                                                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health <= Divisor Then
                                                        If LastDiv > Divisor Then
                                                            Call GiveInvItem(index, Resource(Resource_index).ItemReward, 1)
                                                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), Resource(Resource_index).Color_Success, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                                            LastDiv = Divisor
                                                            Exit For
                                                        End If
                                                    End If
                                                Next II
                                            End If
                                        End If
                                    End If
                                End If
                                ' send the sound
                                SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                                Call CheckTasks(index, QUEST_TYPE_GOTRAIN, Resource_index)
                            Else
                                ' too weak
                                SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                            End If
                        Else
                            ' send message if it exists
                            If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                                SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), Resource(Resource_index).Color_Empty, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            End If
                        End If

                Else
                    PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Bank(index).Item(BankSlot).num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal Amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, invSlot)).Stackable > 0 Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(index, BankSlot)).Stackable > 0 Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim EXP As Long

    ' Calculate exp to give attacker
    EXP = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If EXP < 0 Then EXP = 0
    If EXP = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - EXP)
        SendEXP index
        Call PlayerMsg(index, "You lost " & EXP & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, I As Long, tempItem As Long, x As Long, y As Long, itemnum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        itemnum = GetPlayerInvItemNum(index, invNum)
        
        Player(index).Inv(invNum).Selected = 0
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 2
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                                
                'skill Requirement
                If Item(itemnum).CombatTypeReq > 0 Then
                    If GetPlayerCombatLevel(index, Item(itemnum).CombatTypeReq) < Item(itemnum).CombatLvlReq Then
                        PlayerMsg index, "You do not meet the combat level requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to equip this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemnum, Armor
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 2
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to equip this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If

                If GetPlayerEquipment(index, weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, weapon)
                End If

                SetPlayerEquipment index, itemnum, weapon
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' Check if item is two handed
                If Item(itemnum).Handed > 0 Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        tempItem = GetPlayerEquipment(index, Shield)
        
                        SetPlayerEquipment index, 0, Shield
        
                        GiveInvItem index, tempItem, 0 ' give back the stored item
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 2
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to equip this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemnum, Helmet
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 2
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to equip this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, itemnum, Shield
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' Check if player has on two handed weapon
                If GetPlayerEquipment(index, weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, weapon)).Handed > 0 Then
                        tempItem = GetPlayerEquipment(index, weapon)
        
                        SetPlayerEquipment index, 0, weapon
        
                        GiveInvItem index, tempItem, 0 ' give back the stored item
                        tempItem = 0
                    End If
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 2
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to use this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).num, 1)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 2
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to use this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If itemnum = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                        temptile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        temptile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        SendMapKey index, x, y, 1
                        Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, itemnum, 1)
                            Call PlayerMsg(index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SPELL
            
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to use this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        I = Spell(n).LevelReq

                        If I <= GetPlayerLevel(index) Then
                            I = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If I > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Call SetPlayerSpell(index, I, n)
                                    Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemnum, 1)
                                    Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & I & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_BOOK
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, I) < Item(itemnum).Stat_Req(I) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                'another skill requirement
                If Item(itemnum).Stat_Req(6) > 0 Then
                    If Item(itemnum).SkillReq > 0 Then
                        If GetPlayerSkillLevel(index, Item(itemnum).SkillReq) < Item(itemnum).Stat_Req(6) Then
                            PlayerMsg index, "You do not meet the skill level requirement to use this item.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If
                
                'we're good so yeah
                SendOpenBook index
        End Select
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Client Character Editor ::
' :::::::::::::::::::::::::::::

