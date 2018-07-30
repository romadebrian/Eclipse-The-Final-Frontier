Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Sub CheckHighlight(ByVal Index As Long, ByVal invNum As Long)
Dim reSet As Boolean, Sel1 As Boolean, Sel2 As Boolean
Dim Sel_Index As Long, itemnum As Long
Dim i As Long, II As Long
Dim aiiSelected As Boolean
    ' if selected, unselect
    If Player(Index).Inv(invNum).Selected = 1 Then
        Player(Index).Inv(invNum).Selected = 0
    Else
        ' highlight the item
        Player(Index).Inv(invNum).Selected = 1
        reSet = False
        
        ' see if another one is selected, if so, get ready to use item combo system
        For i = 1 To MAX_INV
            If Player(Index).Inv(i).Selected = 1 And i <> invNum Then
            
                ' Run through combos to see if we have one for these items
                For II = 1 To MAX_COMBOS
                    If Combo(II).Item_1 > 1 Or Combo(II).Item_2 > 1 Then
                        itemnum = GetPlayerInvItemNum(Index, invNum)
                        
                        ' Check if we have item 1 in the slot we just clicked
                        If itemnum = Combo(II).Item_1 Then Sel1 = True
                        
                        ' Check if we have item 2 in the slot we just clicked
                        If itemnum = Combo(II).Item_2 Then Sel1 = True
                        
                        
                        itemnum = GetPlayerInvItemNum(Index, i)
                            
                        ' Check if we have item 1 in the other slot
                        If itemnum = Combo(II).Item_1 Then Sel2 = True
                            
                        ' Check if we have item 2 in the other slot
                        If itemnum = Combo(II).Item_2 Then Sel2 = True
                    End If
                    
                    
                    ' Leave the loop if we found a combo
                    If Sel1 = True And Sel2 = True Then
                        Sel_Index = II
                        Exit For
                    End If
                    
                    Sel1 = False
                    Sel2 = False
                Next II
                
                ' If both items are part of a combo then we're moving on
                If Sel1 = True And Sel2 = True Then
                    aiiSelected = True
                End If
                
                Player(Index).Inv(i).Selected = 0
                reSet = True
            End If
        Next i
    End If
    
    'use item combo system
    If aiiSelected Then
        ' Check requirements
        If Combo(Sel_Index).Level > 0 And GetPlayerLevel(Index) < Combo(Sel_Index).Level Then
            Call PlayerMsg(Index, "You must have a combat level equal to or higher than " & Combo(Sel_Index).Level & " to combine these items.", BrightRed)
            GoTo Continue
        End If
        If Combo(Sel_Index).Skill > 0 And GetPlayerSkillLevel(Index, Combo(Sel_Index).Skill) < Combo(Sel_Index).SkillLevel Then
            Call PlayerMsg(Index, "Your " & Trim$(Skill(Combo(Sel_Index).Skill).Name) & " level must equal to or higher than " & Combo(Sel_Index).SkillLevel & " to combine these items.", BrightRed)
            GoTo Continue
        End If
        If Combo(Sel_Index).ReqItem1 > 0 And HasItems(Index, Combo(Sel_Index).ReqItem1, Combo(Sel_Index).ReqItemVal1) = False Then
            If Item(Combo(Sel_Index).ReqItem1).Type = ITEM_TYPE_CURRENCY Then
                Call PlayerMsg(Index, "You need " & Combo(Sel_Index).ReqItemVal1 & " " & Trim$(Item(Combo(Sel_Index).ReqItem1).Name) & " to combine these items.", BrightRed)
            Else
                Call PlayerMsg(Index, "You need " & CheckGrammar(Trim$(Item(Combo(Sel_Index).ReqItem1).Name)) & " to combine these items.", BrightRed)
            End If
            GoTo Continue
        End If
        If Combo(Sel_Index).ReqItem2 > 0 And HasItems(Index, Combo(Sel_Index).ReqItem2, Combo(Sel_Index).ReqItemVal2) = False Then
            If Item(Combo(Sel_Index).ReqItem2).Type = ITEM_TYPE_CURRENCY Then
                Call PlayerMsg(Index, "You need " & Combo(Sel_Index).ReqItemVal2 & " " & Trim$(Item(Combo(Sel_Index).ReqItem2).Name) & " to combine these items.", BrightRed)
            Else
                Call PlayerMsg(Index, "You need " & CheckGrammar(Trim$(Item(Combo(Sel_Index).ReqItem2).Name)) & " to combine these items.", BrightRed)
            End If
            GoTo Continue
        End If
                
    
        ' Take items
        If Combo(Sel_Index).Take_Item1 = 1 Then Call TakeInvItem(Index, Combo(Sel_Index).Item_1, 1)
        If Combo(Sel_Index).Take_Item2 = 1 Then Call TakeInvItem(Index, Combo(Sel_Index).Item_2, 1)
        If Combo(Sel_Index).Take_ReqItem1 = 1 Then Call TakeInvItem(Index, Combo(Sel_Index).ReqItem1, 1)
        If Combo(Sel_Index).Take_ReqItem2 = 1 Then Call TakeInvItem(Index, Combo(Sel_Index).ReqItem2, 1)
                
        ' Give items
        For i = 1 To MAX_COMBO_GIVEN
            If Combo(Sel_Index).Item_Given(i) > 0 Then
                If Item(Combo(Sel_Index).Item_Given(i)).Type = ITEM_TYPE_CURRENCY Then
                    Call GiveInvItem(Index, Combo(Sel_Index).Item_Given(i), Combo(Sel_Index).Item_Given_Val(i))
                Else
                    For II = 1 To Combo(Sel_Index).Item_Given_Val(i)
                        Call GiveInvItem(Index, Combo(Sel_Index).Item_Given(i), 1)
                    Next II
                End If
            End If
        Next i
        
        If Combo(Sel_Index).GiveSkill > 0 Then
            Call SetPlayerSkillExp(Index, Combo(Sel_Index).GiveSkill, Combo(Sel_Index).GiveSkill_Exp)
            Call SendPlayerData(Index)
            Call PlayerMsg(Index, "You gain " & Combo(Sel_Index).GiveSkill_Exp & " " & Trim$(Skill(Combo(Sel_Index).GiveSkill).Name) & " experience.", Cyan)
        End If
    End If
Continue:
    
    If reSet Then
        ' Remove all highlights
        For i = 1 To MAX_INV
            Player(Index).Inv(i).Selected = 0
            SendHighlight Index, i
        Next i
    End If
End Sub

Sub RemoveFriend(ByVal Index As Long, ByVal fName As String)
Dim i As Long, Place As Long, pI As Long, fOther As String
    pI = FindPlayer(fName)
    fOther = GetPlayerName(Index)
    
    ' Do the first player
    For i = 1 To GetPlayerFriends(Index)
        If GetPlayerFriendName(Index, i) = fName Then
            Place = i
            Call SetPlayerFriends(Index, -1)
        End If
        
        If Place > 0 And Place < i Then
            Call SetPlayerFriendName(Index, i - 1, GetPlayerFriendName(Index, i))
            SetPlayerFriendName Index, i
        End If
    Next i
    Place = 0
    
    ' Do the other player
    For i = 1 To GetPlayerFriends(pI)
        If GetPlayerFriendName(pI, i) = fOther Then
            Place = i
            SetPlayerFriendName pI, i
            Call SetPlayerFriends(pI, -1)
        End If
        
        If Place > 0 And Place < i Then
            GetPlayerFriendName(pI, i - 1) = GetPlayerFriendName(pI, i)
        End If
    Next i
End Sub

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If isPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If isPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).canDespawn = canDespawn
            MapItem(mapnum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).num = itemnum
            MapItem(mapnum, i).Value = ItemVal
            MapItem(mapnum, i).x = x
            MapItem(mapnum, i).y = y
            ' send to map
            SendSpawnItemToMap mapnum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(mapnum).Tile(x, y).Data1).Stackable > 0 And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Dim temp As Long
    
    'Make sure (High) is actually the high number
    If Low > High Then
        temp = High
        High = Low
        Low = temp
    End If
    
    'continue
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapnum As Long, Optional ForcedSpawn As Boolean = False)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean
    Dim HPRndNum As Long
    Dim NText As String
    Dim SEP_CHAR As String * 1

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapnum).NPC(mapNpcNum)
    If ForcedSpawn = False And Map(mapnum).NpcSpawnType(mapNpcNum) = 1 Then npcNum = 0
    If npcNum > 0 Then
        NText = Replace$(NPC(npcNum).Name, SEP_CHAR, vbNullString)
        If Len(NText) < 1 Then Exit Sub
    
        MapNpc(mapnum).NPC(mapNpcNum).num = npcNum
        MapNpc(mapnum).NPC(mapNpcNum).target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0 ' clear
        
        If NPC(mapNpcNum).RandHP = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).HPSetTo = GetNpcMaxVital(npcNum, Vitals.HP)
            MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).HPSetTo
        Else
            HPRndNum = rand(NPC(npcNum).HPMin, NPC(npcNum).HP)
            MapNpc(mapnum).NPC(mapNpcNum).HPSetTo = HPRndNum
            MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = HPRndNum
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapnum).MaxX
            For y = 0 To Map(mapnum).MaxY
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapnum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(mapnum).NPC(mapNpcNum).x = x
                        MapNpc(mapnum).NPC(mapNpcNum).y = y
                        MapNpc(mapnum).NPC(mapNpcNum).Dir = Map(mapnum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(mapnum).MaxX)
                y = Random(0, Map(mapnum).MaxY)
    
                If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
                If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).NPC(mapNpcNum).x = x
                    MapNpc(mapnum).NPC(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapnum).MaxX
                For y = 0 To Map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).NPC(mapNpcNum).x = x
                        MapNpc(mapnum).NPC(mapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).num
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).HPSetTo
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
            UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, True
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    Else
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0 ' clear
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong mapNpcNum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

Public Sub SpawnMapEventsFor(Index As Long, mapnum As Long)
Dim i As Long, x As Long, y As Long, z As Long, spawncurrentevent As Boolean, p As Long
Dim Buffer As clsBuffer
    
    TempPlayer(Index).EventMap.CurrentEvents = 0
    ReDim TempPlayer(Index).EventMap.EventPages(0)
    
    If Map(mapnum).EventCount <= 0 Then Exit Sub
    For i = 1 To Map(mapnum).EventCount
        If Map(mapnum).Events(i).PageCount > 0 Then
            For z = Map(mapnum).Events(i).PageCount To 1 Step -1
                With Map(mapnum).Events(i).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        If Player(Index).Variables(.VariableIndex) < .VariableCondition Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSwitch = 1 Then
                        If Player(Index).Switches(.SwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(Index, .HasItemIndex) < .HasItemIndex Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(i).SelfSwitches(.SelfSwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        'spawn the event... send data to player
                        TempPlayer(Index).EventMap.CurrentEvents = TempPlayer(Index).EventMap.CurrentEvents + 1
                        ReDim Preserve TempPlayer(Index).EventMap.EventPages(TempPlayer(Index).EventMap.CurrentEvents)
                        With TempPlayer(Index).EventMap.EventPages(TempPlayer(Index).EventMap.CurrentEvents)
                            If Map(mapnum).Events(i).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(i).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            .GraphicNum = Map(mapnum).Events(i).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(i).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(i).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(i).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(i).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(i).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(i).Pages(z).MoveSpeed
                                Case 0
                                    .movementspeed = 2
                                Case 1
                                    .movementspeed = 3
                                Case 2
                                    .movementspeed = 4
                                Case 3
                                    .movementspeed = 6
                                Case 4
                                    .movementspeed = 12
                                Case 5
                                    .movementspeed = 24
                            End Select
                            If Map(mapnum).Events(i).Global Then
                                .x = TempEventMap(mapnum).Events(i).x
                                .y = TempEventMap(mapnum).Events(i).y
                                .Dir = TempEventMap(mapnum).Events(i).Dir
                                .MoveRouteStep = TempEventMap(mapnum).Events(i).MoveRouteStep
                            Else
                                .x = Map(mapnum).Events(i).x
                                .y = Map(mapnum).Events(i).y
                                .MoveRouteStep = 0
                            End If
                            .Position = Map(mapnum).Events(i).Pages(z).Position
                            .eventID = i
                            .pageID = z
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(mapnum).Events(i).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount)
                                If Map(mapnum).Events(i).Pages(z).MoveRouteCount > 0 Then
                                    For p = 0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(mapnum).Events(i).Pages(z).MoveRoute(p)
                                    Next
                                End If
                            End If
                            
                            .RepeatMoveRoute = Map(mapnum).Events(i).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(i).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(i).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(mapnum).Events(i).Pages(z).WalkAnim
                            .WalkThrough = Map(mapnum).Events(i).Pages(z).WalkThrough
                            .ShowName = Map(mapnum).Events(i).Pages(z).ShowName
                            .FixedDir = Map(mapnum).Events(i).Pages(z).DirFix
                            
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
        For i = 1 To TempPlayer(Index).EventMap.CurrentEvents
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnEvent
            Buffer.WriteLong i
            With TempPlayer(Index).EventMap.EventPages(i)
                Buffer.WriteString Map(GetPlayerMap(Index)).Events(i).Name
                Buffer.WriteLong .Dir
                Buffer.WriteLong .GraphicNum
                Buffer.WriteLong .GraphicType
                Buffer.WriteLong .GraphicX
                Buffer.WriteLong .GraphicX2
                Buffer.WriteLong .GraphicY
                Buffer.WriteLong .GraphicY2
                Buffer.WriteLong .movementspeed
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .Position
                Buffer.WriteLong .Visible
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkAnim
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).DirFix
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkThrough
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).ShowName
            End With
            SendDataTo Index, Buffer.ToArray
            Set Buffer = Nothing
        Next
    End If
End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).NPC(LoopI).num > 0 Then
            If MapNpc(mapnum).NPC(LoopI).x = x Then
                If MapNpc(mapnum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next
    
    For LoopI = 1 To TempEventMap(mapnum).EventCount
        If TempEventMap(mapnum).Events(LoopI).active = 1 Then
            If MapNpc(mapnum).NPC(LoopI).x = TempEventMap(mapnum).Events(LoopI).x Then
                If MapNpc(mapnum).NPC(LoopI).y = TempEventMap(mapnum).Events(LoopI).y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next
    
    CacheMapBlocks mapnum

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Sub SpawnAllMapGlobalEvents()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnGlobalEvents(i)
    Next

End Sub

Sub SpawnGlobalEvents(ByVal mapnum As Long)
    Dim i As Long, z As Long
    
    If Map(mapnum).EventCount > 0 Then
        TempEventMap(mapnum).EventCount = 0
        ReDim TempEventMap(mapnum).Events(0)
        For i = 1 To Map(mapnum).EventCount
            TempEventMap(mapnum).EventCount = TempEventMap(mapnum).EventCount + 1
            ReDim Preserve TempEventMap(mapnum).Events(0 To TempEventMap(mapnum).EventCount)
            If Map(mapnum).Events(i).PageCount > 0 Then
                If Map(mapnum).Events(i).Global = 1 Then
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).x = Map(mapnum).Events(i).x
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).y = Map(mapnum).Events(i).y
                    If Map(mapnum).Events(i).Pages(1).GraphicType = 1 Then
                        Select Case Map(mapnum).Events(i).Pages(1).GraphicY
                            Case 0
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                    End If
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).active = 1
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = Map(mapnum).Events(i).Pages(1).MoveType
                    
                    If TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = 2 Then
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRouteCount = Map(mapnum).Events(i).Pages(1).MoveRouteCount
                        ReDim TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount)
                        For z = 0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount
                            TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(z) = Map(mapnum).Events(i).Pages(1).MoveRoute(z)
                        Next
                    End If
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).RepeatMoveRoute = Map(mapnum).Events(i).Pages(1).RepeatMoveRoute
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveFreq = Map(mapnum).Events(i).Pages(1).MoveFreq
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveSpeed = Map(mapnum).Events(i).Pages(1).MoveSpeed
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkThrough = Map(mapnum).Events(i).Pages(1).WalkThrough
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).FixedDir = Map(mapnum).Events(i).Pages(1).DirFix
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkingAnim = Map(mapnum).Events(i).Pages(1).WalkAnim
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).ShowName = Map(mapnum).Events(i).Pages(1).ShowName
                    
                End If
            End If
        Next
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).NPC(mapNpcNum).x
    y = MapNpc(mapnum).NPC(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x - 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x + 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, False

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).NPC(mapNpcNum).y = MapNpc(mapnum).NPC(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).NPC(mapNpcNum).y = MapNpc(mapnum).NPC(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).NPC(mapNpcNum).x = MapNpc(mapnum).NPC(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).NPC(mapNpcNum).x = MapNpc(mapnum).NPC(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select
    
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, True

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If isPlaying(i) And GetPlayerMap(i) = mapnum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    temptile(mapnum).DoorTimer = 0
    ReDim temptile(mapnum).DoorOpen(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            temptile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(Map(mapnum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(Index, oldSlot)
    NewNum = GetPlayerSpell(Index, newSlot)
    SetPlayerSpell Index, oldSlot, NewNum
    SetPlayerSpell Index, newSlot, OldNum
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        If Item(GetPlayerEquipment(Index, EqSlot)).Stackable > 0 Then
            GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 1
        Else
            GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 0
        End If
        PlayerMsg Index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(Index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SendWornEquipment Index
        SendMapEquipment Index
        SendStats Index
        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Function FindNPCInRange(ByVal Index As Long, ByVal mapnum As Long, ByVal Range As Long)
Dim x As Long, y As Long
Dim npcX As Long, npcY As Long
Dim i As Long

    FindNPCInRange = 0
    
    For i = 1 To MAX_MAP_NPCS
        x = GetPlayerX(Index)
        y = GetPlayerY(Index)
        npcX = MapNpc(mapnum).NPC(i).x
        npcY = MapNpc(mapnum).NPC(i).y
    
        If isInRange(Range, x, y, npcX, npcY) Then
            If MapNpc(mapnum).NPC(i).num > 0 Then
                If NPC(MapNpc(mapnum).NPC(i).num).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(MapNpc(mapnum).NPC(i).num).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
                    FindNPCInRange = i
                    Exit Function
                End If
            End If
        End If
        
    Next i
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(partyNum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> Index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(Index) & " left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(Index).inParty = 0
                        TempPlayer(Index).partyInvite = 0
                        Exit For
                        End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(Index) & " left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(Index).inParty = 0
                        TempPlayer(Index).partyInvite = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party has been disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(partyNum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).partyInvite = 0
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not isPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partyNum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, Index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, Index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = Index
        Party(partyNum).Member(1) = Index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, Index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal targetPlayer As Long)
    PlayerMsg Index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal EXP As Long, ByVal Index As Long, ByVal mapnum As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long, LoseMemberCount As Byte

    ' check if it's worth sharing
    If Not EXP >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, EXP
        Exit Sub
    End If
    
    ' check members in outhers maps
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And isPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> mapnum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next i
    
    ' find out the equal share
    expShare = EXP \ (Party(partyNum).MemberCount - LoseMemberCount)
    leftOver = EXP Mod (Party(partyNum).MemberCount - LoseMemberCount)
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And isPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = mapnum Then
                    ' give them their share
                    GivePlayerEXP tmpIndex, expShare
                End If
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(rand(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal EXP As Long)
    ' give the exp
    Call SetPlayerExp(Index, GetPlayerExp(Index) + EXP)
    SendEXP Index
    SendActionMsg GetPlayerMap(Index), "+" & EXP & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp Index
End Sub

Public Sub GivePlayerCombatEXP(ByVal Index As Long, ByVal skillType As Byte, ByVal EXP As Long)
    If EXP < 0 Then Exit Sub
    If Player(Index).Combat(Item(skillType).CombatTypeReq).Level = MAX_COMBAT_LEVEL Then Exit Sub
    Call SetPlayerCombatExp(Index, Item(skillType).CombatTypeReq, GetPlayerCombatExp(Index, Item(skillType).CombatTypeReq) + EXP)
    SendCombatEXP Index
    CheckCombatLevelUp Index, Item(skillType).CombatTypeReq
End Sub

Function CanEventMove(Index As Long, ByVal mapnum As Long, x As Long, y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional globalevent As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long, z As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    CanEventMove = True
    
    

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_NPCAVOID Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y - 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x) And (MapNpc(mapnum).NPC(i).y = y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y + 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x) And (MapNpc(mapnum).NPC(i).y = y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x - 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x - 1) And (MapNpc(mapnum).NPC(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x - 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x - 1) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If isPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x + 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x + 1) And (MapNpc(mapnum).NPC(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x + 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x + 1) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

    End Select

End Function

Sub EventDir(playerindex As Long, ByVal mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional globalevent As Boolean = False)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(playerindex).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(playerindex).EventMap.EventPages(eventID).Dir = Dir
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventDir
    Buffer.WriteLong eventID
    If globalevent Then
        Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
    Else
        Buffer.WriteLong TempPlayer(playerindex).EventMap.EventPages(eventID).Dir
    End If
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub EventMove(Index As Long, mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, movementspeed As Long, Optional globalevent As Boolean = False)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
        UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, False
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(Index).EventMap.EventPages(eventID).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventID).y = TempPlayer(Index).EventMap.EventPages(eventID).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
            
        Case DIR_DOWN
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventID).y = TempPlayer(Index).EventMap.EventPages(eventID).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_LEFT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventID).x = TempPlayer(Index).EventMap.EventPages(eventID).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_RIGHT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventID).x = TempPlayer(Index).EventMap.EventPages(eventID).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
    End Select

End Sub

Public Sub ClearTopKillData()
Dim i As Integer
    
    On Error GoTo ErrHandler

    For i = 1 To 5
        hkPlace(i) = vbNullString
        hkKills(i) = 0
    Next i

    Reg_Kills = 0
    
    For i = 1 To MAX_PLAYERS
        Player(i).TopKills = 0
    Next i
    
Exit Sub
ErrHandler:
    HandleError "ClearTopKillData", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
End Sub

Sub InitTopKillEvent()
    'On Error GoTo ErrHandler

    With frmServer
        If TopKill_Activated = False And .btnStart.Caption <> "RESET" And .btnStart.Caption <> "STOP" Then
            ' do little shit
            .btnStart.Caption = "STOP"
            .lblStatus.Caption = "NO"
            .shpColor.Visible = True
            .tmrGetTime.Enabled = True
            .lblTime.Visible = True
            
            'Tell everyone
            If Len(.txtStartMsg.Text) > 0 Then
                If .cboColor_Start.ListIndex > -1 Then
                    Call GlobalMsg(ConfigureTopKillMsg(.txtStartMsg.Text), .cboColor_Start.ListIndex)
                Else
                    Call GlobalMsg(ConfigureTopKillMsg(.txtStartMsg.Text), White)
                End If
            Else
                Call GlobalMsg("Server: TopKill Event has been initialized.  First person to " & frmServer.scrlNeeded.Value & " kills wins.", White)
                Call GlobalMsg("TopKill Event -- First Place Exp: " & .scrlFirst.Value, White)
                Call GlobalMsg("TopKill Event -- Second Place Exp: " & .scrlSecond.Value, White)
                If .chk3.Value = vbChecked Then Call GlobalMsg("TopKill Event -- Third Place Exp: " & .scrlThird.Value, White)
                If .chk4.Value = vbChecked Then Call GlobalMsg("TopKill Event -- Fourth Place Exp: " & .scrlFourth.Value, White)
                If .chk5.Value = vbChecked Then Call GlobalMsg("TopKill Event -- Fifth Place Exp: " & .scrlFifth.Value, White)
            End If
            
            'Show in server
            Call TextAdd("Server: TopKill Event Initiated. (Made By. escfoe2)")

            'Disable opts
            .scrlNeeded.Enabled = False
            .scrlFirst.Enabled = False
            .scrlSecond.Enabled = False
            .scrlThird.Enabled = False
            .scrlFourth.Enabled = False
            .scrlFifth.Enabled = False
            .chk3.Enabled = False
            .chk4.Enabled = False
            .chk5.Enabled = False
            
            'Activate catch variable
            TopKill_Activated = True
        Else
            'do if event was stopped early
            If .btnStart.Caption = "STOP" Then
                'Deactivate catch variable
                TopKill_Activated = False
                'Stop timer
                .tmrGetTime.Enabled = False
                'Tell Players
                Call GlobalMsg("Server: The Kill Event has been ended early. We apologize for the incovenience. 1/2 exp will be awarded to those who placed... These people are as follows:", White)
                If Len(hkPlace(1)) > 0 Then Call GlobalMsg(hkPlace(1) & " -- 1st Place -- " & (.scrlFirst.Value / 2) & " exp.", Green)
                If Len(hkPlace(2)) > 0 Then Call GlobalMsg(hkPlace(2) & " -- 2nd Place -- " & (.scrlSecond.Value / 2) & " exp.", Green)
                If .chk3.Value = vbChecked Then If Len(hkPlace(3)) > 0 Then Call GlobalMsg(hkPlace(3) & " -- 3rd Place -- " & (.scrlThird.Value / 2) & " exp.", Green)
                If .chk4.Value = vbChecked Then If Len(hkPlace(4)) > 0 Then Call GlobalMsg(hkPlace(4) & " -- 4th Place -- " & (.scrlFourth.Value / 2) & " exp.", Green)
                If .chk5.Value = vbChecked Then If Len(hkPlace(5)) > 0 Then Call GlobalMsg(hkPlace(5) & " -- 5th Place -- " & (.scrlFifth.Value / 2) & " exp.", Green)
                'Give exp
                If Len(hkPlace(1)) > 0 Then Call SetPlayerExp(FindPlayer(hkPlace(1)), GetPlayerExp(FindPlayer(hkPlace(1))) + (.scrlFirst.Value / 2))
                If Len(hkPlace(1)) > 0 Then Call SendEXP(FindPlayer(hkPlace(1)))
                If Len(hkPlace(2)) > 0 Then Call SetPlayerExp(FindPlayer(hkPlace(2)), GetPlayerExp(FindPlayer(hkPlace(2))) + (.scrlSecond.Value / 2))
                If Len(hkPlace(2)) > 0 Then Call SendEXP(FindPlayer(hkPlace(2)))
                If Len(hkPlace(3)) > 0 Then Call SetPlayerExp(FindPlayer(hkPlace(3)), GetPlayerExp(FindPlayer(hkPlace(3))) + (.scrlThird.Value / 2))
                If Len(hkPlace(3)) > 0 And .chk3.Value = vbChecked Then Call SendEXP(FindPlayer(hkPlace(3)))
                If Len(hkPlace(4)) > 0 Then Call SetPlayerExp(FindPlayer(hkPlace(4)), GetPlayerExp(FindPlayer(hkPlace(4))) + (.scrlFourth.Value / 2))
                If Len(hkPlace(4)) > 0 And .chk4.Value = vbChecked Then Call SendEXP(FindPlayer(hkPlace(4)))
                If Len(hkPlace(5)) > 0 Then Call SetPlayerExp(FindPlayer(hkPlace(5)), GetPlayerExp(FindPlayer(hkPlace(5))) + (.scrlFifth.Value / 2))
                If Len(hkPlace(5)) > 0 And .chk5.Value = vbChecked Then Call SendEXP(FindPlayer(hkPlace(5)))
                'Set button text
                .btnStart.Caption = "RESET"
                'show in server
                Call TextAdd("Server: TopKill Event has been stopped prematurely.")
                Exit Sub
            End If
            
            'show in server
            Call TextAdd("Server: TopKill Event reset successfully. (Made By: escfoe2)")
            
            'do little shit
            .btnStart.Caption = "Activate TopKill Event"
            .lblStatus.Caption = "??"
            .shpColor.BackColor = vbRed
            .shpColor.Visible = False
            .tmrGetTime.Enabled = False
            Time_Seconds = 0
            Time_Minutes = 0
            Time_Hours = 0
            .lblTime.Caption = "Run Time:  00:00:00"
            .lblTime.Visible = False
            
            'Deactivate catch variable
            TopKill_Activated = False
            
            'Reset items
            .scrlNeeded.Enabled = True
            .scrlFirst.Enabled = True
            .scrlSecond.Enabled = True
            .scrlThird.Enabled = True
            .scrlFourth.Enabled = True
            .scrlFifth.Enabled = True
            .chk3.Enabled = True
            .chk4.Enabled = True
            .chk5.Enabled = True
            .lblFirst.Caption = "N/A"
            .lblSecond.Caption = "N/A"
            .lblThird.Caption = "N/A"
            .lblFourth.Caption = "N/A"
            .lblFifth.Caption = "N/A"
            .lblTotal = "Total Kills: N/A"
            .lblWinner.Caption = "Waiting..."
            
            ' clear cache
            Call ClearTopKillData
            'Send data thus far to players.
            
        End If
    End With
    
Exit Sub
ErrHandler:
    HandleError "InitTopKillEvent", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
End Sub

Public Sub CheckTopKill(ByVal Index As Long)
Dim tempName1 As String, MyName As String, Color As Integer
Dim tempKills1 As Long, TempIndex As Long, tempExp As Long, MyKills As Long, MyPlace As Byte

    On Error GoTo ErrHandler
    
    'set mykills
    MyKills = Player(Index).TopKills
    MyName = Trim$(Player(Index).Name)
    
    ' check what place player should be in and update score table
    If MyKills >= hkKills(5) And MyKills < hkKills(4) Then
        If Not frmServer.chk5.Value = vbChecked Then GoTo Continue
        ' Show player what place he/she is in.
        MyPlace = 5
        'Update Fifth Place
        hkPlace(5) = MyName
        hkKills(5) = MyKills
    ElseIf MyKills >= hkKills(4) And MyKills < hkKills(3) Then
        If Not frmServer.chk4.Value = vbChecked Then GoTo Continue
        ' Show player what place he/she is in.
        MyPlace = 4
        ' See if the only thing necessary is updating a current table position
        If hkPlace(4) = MyName Then
            hkKills(4) = MyKills
            GoTo Continue
        End If
        'Update Fourth Place
        tempName1 = hkPlace(4)
        tempKills1 = hkKills(4)
        hkPlace(4) = MyName
        hkKills(4) = MyKills
        'Update Fifth Place
        hkPlace(5) = tempName1
        hkKills(5) = tempKills1
    ElseIf MyKills >= hkKills(3) And MyKills < hkKills(2) Then
        If Not frmServer.chk3.Value = vbChecked Then GoTo Continue
        ' Show player what place he/she is in.
        MyPlace = 3
        ' See if the only thing necessary is updating a current table position
        If hkPlace(3) = MyName Then
            hkKills(3) = MyKills
            GoTo Continue
        End If
        'Update Third Place
        tempName1 = hkPlace(3)
        tempKills1 = hkKills(3)
        hkPlace(3) = MyName
        hkKills(3) = MyKills
        'Update Fourth Place
        hkPlace(4) = tempName1
        hkKills(4) = tempKills1
    ElseIf MyKills >= hkKills(2) And MyKills < hkKills(1) Then
        ' Show player what place he/she is in.
        MyPlace = 2
        ' See if the only thing necessary is updating a current table position
        If hkPlace(2) = MyName Then
            hkKills(2) = MyKills
            GoTo Continue
        End If
        'Update Second Place
        tempName1 = hkPlace(2)
        tempKills1 = hkKills(2)
        hkPlace(2) = MyName
        hkKills(2) = MyKills
        'Update Third Place
        hkPlace(3) = tempName1
        hkKills(3) = tempKills1
    ElseIf MyKills >= hkKills(1) Then
        ' Show player what place he/she is in.
        MyPlace = 1
        ' See if the only thing necessary is updating a current table position
        If hkPlace(1) = MyName Then
            hkKills(1) = MyKills
            GoTo Continue
        End If
        ' Update first Place
        tempName1 = hkPlace(1)
        tempKills1 = hkKills(1)
        hkPlace(1) = MyName
        hkKills(1) = MyKills
        'Update Second Place
        hkPlace(2) = tempName1
        hkKills(2) = tempKills1
    Else
        'Didn't place, oh well, add to total kill count anyway :p
        Reg_Kills = Reg_Kills + 1
    End If
Continue:

    'if selected and player has actually placed, send actionmsg
    If MyPlace > 0 Then
        If frmServer.chkActionMsg.Value = vbChecked Then
            'get the color
            If frmServer.cboColor_ActionMsg.ListIndex < 0 Then
                Color = Red
            Else
                Color = frmServer.cboColor_ActionMsg.ListIndex
            End If
            
            'check for msg to send, if nothing, revert to default
            If Len(frmServer.txtActionMsg.Text) > 0 Then
                'custom
                If frmServer.opStat.Value Then
                    Call SendActionMsg(GetPlayerMap(Index), ConfigureTopKillMsg(frmServer.txtActionMsg.Text, MyPlace), Color, 0, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
                Else
                    Call SendActionMsg(GetPlayerMap(Index), ConfigureTopKillMsg(frmServer.txtActionMsg.Text, MyPlace), Color, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
                End If
            Else
                'default
                If frmServer.opStat.Value Then
                    Call SendActionMsg(GetPlayerMap(Index), Placements(MyPlace) & " Place", Color, 0, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
                Else
                    Call SendActionMsg(GetPlayerMap(Index), Placements(MyPlace) & " Place", Color, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
                End If
            End If
        End If
    End If
    'Update Server Data
    With frmServer
        If Len(Trim$(hkPlace(1))) > 0 Then .lblFirst.Caption = hkPlace(1) & " (" & hkKills(1) & ")" ' If placement 1 is being used do this
        If Len(Trim$(hkPlace(1))) < 1 Then .lblFirst.Caption = "N/A" 'If not do this
        If Len(Trim$(hkPlace(2))) > 0 Then .lblSecond.Caption = hkPlace(2) & " (" & hkKills(2) & ")" ' If placement 2 is being used do this
        If Len(Trim$(hkPlace(2))) < 1 Then .lblSecond.Caption = "N/A" 'If not do this
        If Len(Trim$(hkPlace(3))) > 0 Then .lblThird.Caption = hkPlace(3) & " (" & hkKills(3) & ")" ' If placement 3 is being used do this
        If Len(Trim$(hkPlace(3))) < 1 Then .lblThird.Caption = "N/A" 'If not do this
        If Len(Trim$(hkPlace(4))) > 0 Then .lblFourth.Caption = hkPlace(4) & " (" & hkKills(4) & ")" ' If placement 4 is being used do this
        If Len(Trim$(hkPlace(4))) < 1 Then .lblFourth.Caption = "N/A" 'If not do this
        If Len(Trim$(hkPlace(5))) > 0 Then .lblFifth.Caption = hkPlace(5) & " (" & hkKills(5) & ")" ' If placement 5 is being used do this
        If Len(Trim$(hkPlace(5))) < 1 Then .lblFifth.Caption = "N/A" 'If not do this
        
        .lblTotal.Caption = "Total Kills: " & hkKills(1) + hkKills(2) + hkKills(3) + hkKills(4) + hkKills(5) + Reg_Kills
    End With
        
    'If selected, tell the player how close (s)he is to winning
    If frmServer.chkPlayerMsg.Value = vbChecked Then
        ' get the color
        If frmServer.cboColor_PlayerMsg.ListIndex < 0 Then
            Color = Green
        Else
            Color = frmServer.cboColor_PlayerMsg.ListIndex
        End If
        
        'check for msg to send, if nothing, revert to default
        If Len(frmServer.txtPlayerMsg.Text) > 0 Then
            'custom
            Call PlayerMsg(Index, ConfigureTopKillMsg(frmServer.txtPlayerMsg.Text, MyPlace, Index), Color)
        Else
        'default
            Call PlayerMsg(Index, "TopKill Event: " & MyKills & "/" & frmServer.scrlNeeded.Value, Color)
        End If
    End If
    
    'Check if limit has been reached
    If hkKills(1) >= frmServer.scrlNeeded.Value Then
        TopKill_Activated = False
        frmServer.lblWinner.Caption = hkPlace(1)
        frmServer.lblStatus.Caption = "YES"
        frmServer.shpColor.BackColor = vbGreen
        frmServer.lblTotal = "^WINNER^"
        frmServer.btnStart.Caption = "RESET"
        frmServer.tmrGetTime.Enabled = False
        
        'dynamic win msg
        If Len(frmServer.txtEndMsg.Text) > 0 Then
            'custom
            If frmServer.cboColor_End.ListIndex > -1 Then
                Call GlobalMsg(ConfigureTopKillMsg(frmServer.txtEndMsg.Text, MyPlace, Index), frmServer.cboColor_End.ListIndex)
            Else
                Call GlobalMsg(ConfigureTopKillMsg(frmServer.txtEndMsg.Text, MyPlace, Index), White)
            End If
        Else
            'default
            tempKills1 = hkKills(1) - hkKills(2)
            If hkKills(1) - hkKills(2) > 0 Then
                Call GlobalMsg("Server: The TopKill Event has ended... Our winner is " & hkPlace(1) & " who won by " & tempKills1 & " kill(s)!!  Good game everyone!!", White)
            Else
                Call GlobalMsg("Server: The TopKill Event has ended... The game was a tie!! Our winners are " & hkPlace(1) & " and " & hkPlace(2) & "!!  Great game everyone!!", White)
            End If
        End If
        
        'Show on server
        Call TextAdd("Server: TopKill Event has finished. There is a winner.")
        
        'give exp
        If Len(hkPlace(1)) > 0 Then
            TempIndex = FindPlayer(hkPlace(1))
            tempExp = frmServer.scrlFirst.Value
            If tempExp > 0 Then
                Call SetPlayerExp(TempIndex, GetPlayerExp(Index) + tempExp)
                Call SendEXP(TempIndex)
                Call PlayerMsg(TempIndex, "You receive " & tempExp & " exp for First Place. Perfect!!!", Yellow)
            End If
        End If
        If Len(hkPlace(2)) > 0 Then
            TempIndex = FindPlayer(hkPlace(2))
            tempExp = frmServer.scrlSecond.Value
            If tempExp > 0 Then
                Call SetPlayerExp(TempIndex, GetPlayerExp(TempIndex) + tempExp)
                Call SendEXP(TempIndex)
                Call PlayerMsg(TempIndex, "You receive " & tempExp & " exp for Second Place. Fantastic!!", Yellow)
            End If
        End If
        If Len(hkPlace(3)) > 0 Then
            If Not frmServer.chk3.Value = vbChecked Then GoTo Continue2
            TempIndex = FindPlayer(hkPlace(3))
            tempExp = frmServer.scrlThird.Value
            Call SetPlayerExp(TempIndex, GetPlayerExp(TempIndex) + tempExp)
            Call SendEXP(TempIndex)
            Call PlayerMsg(TempIndex, "You receive " & tempExp & " exp for Third Place. Nicely Done!", Yellow)
        End If
        If Len(hkPlace(4)) > 0 Then
            If Not frmServer.chk4.Value = vbChecked Then GoTo Continue2
            TempIndex = FindPlayer(hkPlace(4))
            tempExp = frmServer.scrlFourth.Value
            If tempExp > 0 Then
                Call SetPlayerExp(TempIndex, GetPlayerExp(TempIndex) + tempExp)
                Call SendEXP(TempIndex)
                Call PlayerMsg(TempIndex, "You receive " & tempExp & " exp for Fourth Place. Great Job", Yellow)
            End If
        End If
        If Len(hkPlace(5)) > 0 Then
            If Not frmServer.chk5.Value = vbChecked Then GoTo Continue2
            TempIndex = FindPlayer(hkPlace(5))
            tempExp = frmServer.scrlFifth.Value
            If tempExp > 0 Then
                Call SetPlayerExp(TempIndex, GetPlayerExp(TempIndex) + tempExp)
                Call SendEXP(TempIndex)
                Call PlayerMsg(TempIndex, "You receive " & tempExp & " exp for Fifth Place. Good Job", Yellow)
            End If
        End If
        
Continue2:
        'Clear variable data
        Call ClearTopKillData
    End If
    Exit Sub
'Handle Error
ErrHandler:
    Call HandleError("modPlayer", "CheckTopKill", Err.Number, Err.Description, Err.Source, Err.HelpContext)
End Sub

Public Function ConfigureTopKillMsg(ByVal Msg As String, Optional ByVal Place As Byte, Optional ByVal Index As Long) As String
Dim NewMsg(0 To 18) As String, Kills As Long
    On Error GoTo ErrHandler
    
    'if nothing in msg exit out
    If Len(Msg) < 1 Then Exit Function

    'Withdraw and replace keywords
    With frmServer
        ' Start message keywords
        NewMsg(0) = Replace$(Msg, "#1stexp#", CStr(.scrlFirst.Value))
        NewMsg(1) = Replace$(NewMsg(0), "#2ndexp#", CStr(.scrlSecond.Value))
        NewMsg(2) = Replace$(NewMsg(1), "#3rdexp#", CStr(.scrlThird.Value))
        NewMsg(3) = Replace$(NewMsg(2), "#4thexp#", CStr(.scrlFourth.Value))
        NewMsg(4) = Replace$(NewMsg(3), "#5thexp#", CStr(.scrlFifth.Value))
        NewMsg(5) = Replace$(NewMsg(4), "#getkills#", CStr(.scrlNeeded.Value))
        
        ' End message keywords
        NewMsg(6) = Replace$(NewMsg(5), "#1stkills#", CStr(hkKills(1)))
        NewMsg(7) = Replace$(NewMsg(6), "#2ndkills#", CStr(hkKills(2)))
        NewMsg(8) = Replace$(NewMsg(7), "#3rdkills#", CStr(hkKills(3)))
        NewMsg(9) = Replace$(NewMsg(8), "#4thkills#", CStr(hkKills(4)))
        NewMsg(10) = Replace$(NewMsg(9), "#5thkills#", CStr(hkKills(5)))
        NewMsg(11) = Replace$(NewMsg(10), "#1stname#", CStr(hkPlace(1)))
        NewMsg(12) = Replace$(NewMsg(11), "#2ndname#", CStr(hkPlace(2)))
        NewMsg(13) = Replace$(NewMsg(12), "#3rdname#", CStr(hkPlace(3)))
        NewMsg(14) = Replace$(NewMsg(13), "#4thname#", CStr(hkPlace(4)))
        NewMsg(15) = Replace$(NewMsg(14), "#5thname#", CStr(hkPlace(5)))
        Kills = Reg_Kills + hkKills(1) + hkKills(2) + hkKills(3) + hkKills(4) + hkKills(5)
        
        'end msg/action msg/player msg
        NewMsg(16) = Replace$(NewMsg(15), "#totalkills#", CStr(Kills))
        NewMsg(17) = NewMsg(16)
        If Place > 0 Then NewMsg(17) = Replace$(NewMsg(16), "#placement#", Placements(Place))
        NewMsg(18) = NewMsg(17)
        If Index > 0 Then NewMsg(18) = Replace$(NewMsg(17), "#playerkills#", Player(Index).TopKills)
    End With
    
    'Send back the filtered msg
    ConfigureTopKillMsg = NewMsg(18)
    
    Exit Function
ErrHandler:
    HandleError "ConfigTopKillStartMessage", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
End Function

Public Sub HandleProjecTile(ByVal Index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, i As Long

    ' check for subscript out of range
    If Index < 1 Or Index > MAX_PLAYERS Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
        
    ' check to see if it's time to move the Projectile
    If GetTickCount > TempPlayer(Index).ProjecTile(PlayerProjectile).TravelTime Then
        With TempPlayer(Index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case DIR_DOWN
                    .y = .y + 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(Index) + .Range) + 1 Then
                        ClearProjectile Index, PlayerProjectile
                        Exit Sub
                    End If
                ' up
                Case DIR_UP
                    .y = .y - 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(Index) - .Range) - 1 Then
                        ClearProjectile Index, PlayerProjectile
                        Exit Sub
                    End If
                ' right
                Case DIR_RIGHT
                    .x = .x + 1
                    ' check if they reached max range
                    If .x = (GetPlayerX(Index) + .Range) + 1 Then
                        ClearProjectile Index, PlayerProjectile
                        Exit Sub
                    End If
                ' left
                Case DIR_LEFT
                    .x = .x - 1
                    ' check if they reached maxrange
                    If .x = (GetPlayerX(Index) - .Range) - 1 Then
                        ClearProjectile Index, PlayerProjectile
                        Exit Sub
                    End If
            End Select
            .TravelTime = GetTickCount + .Speed
        End With
    End If
    
    x = TempPlayer(Index).ProjecTile(PlayerProjectile).x
    y = TempPlayer(Index).ProjecTile(PlayerProjectile).y
    
    ' check if left map
    If x > Map(GetPlayerMap(Index)).MaxX Or y > Map(GetPlayerMap(Index)).MaxY Or x < 0 Or y < 0 Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if hit player
    For i = 1 To Player_HighIndex
        ' make sure they're actually playing
        If isPlaying(i) Then
            ' check coordinates
            If x = Player(i).x And y = GetPlayerY(i) Then
                ' make sure it's not the attacker
                If Not x = Player(Index).x Or Not y = GetPlayerY(Index) Then
                    ' check if player can attack
                    If CanPlayerAttackPlayer(Index, i, False, True) = True Then
                        ' attack the player and kill the project tile
                        PlayerAttackPlayer Index, i, TempPlayer(Index).ProjecTile(PlayerProjectile).Damage
                        ClearProjectile Index, PlayerProjectile
                        Exit Sub
                    Else
                        ClearProjectile Index, PlayerProjectile
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If x = MapNpc(GetPlayerMap(Index)).NPC(i).x And y = MapNpc(GetPlayerMap(Index)).NPC(i).y Then
            ' they're hit, remove it and deal that damage ;)
            If CanPlayerAttackNpc(Index, i, True) Then
                PlayerAttackNpc Index, i, TempPlayer(Index).ProjecTile(PlayerProjectile).Damage
                ClearProjectile Index, PlayerProjectile
                Exit Sub
            Else
                ClearProjectile Index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' hit a block
    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        ' hit a block, clear it.
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
End Sub
