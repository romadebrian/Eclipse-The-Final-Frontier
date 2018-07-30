Attribute VB_Name = "modEventLogic"
Option Explicit
Public Sub RemoveDeadEvents()
Dim I As Long, mapnum As Long, buffer As clsBuffer, x As Long, id As Long, page As Long
    'See if we should remove any events....
    For I = 1 To Player_HighIndex
        If TempPlayer(I).EventMap.CurrentEvents > 0 Then
            mapnum = GetPlayerMap(I)
            For x = 1 To TempPlayer(I).EventMap.CurrentEvents
                id = TempPlayer(I).EventMap.EventPages(x).eventID
                page = TempPlayer(I).EventMap.EventPages(x).pageID
                If Map(mapnum).Events(id).PageCount >= page Then
                
                    'See if there is any reason to delete this event....
                    'In other words, go back through conditions and make sure they all check up.
                    If TempPlayer(I).EventMap.EventPages(x).Visible = 1 Then
                        If Map(mapnum).Events(id).Pages(page).chkHasItem = 1 Then
                            If HasItem(I, Map(mapnum).Events(id).Pages(page).HasItemIndex) = 0 Then
                                TempPlayer(I).EventMap.EventPages(x).Visible = 0
                            End If
                        End If
                        
                        If Map(mapnum).Events(id).Pages(page).chkSelfSwitch = 1 Then
                            If Map(mapnum).Events(id).Pages(page).SelfSwitchCompare = 0 Then
                                If Map(mapnum).Events(id).SelfSwitches(Map(mapnum).Events(id).Pages(page).SelfSwitchIndex) = 0 Then
                                    TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                End If
                            Else
                                If Map(mapnum).Events(id).SelfSwitches(Map(mapnum).Events(id).Pages(page).SelfSwitchIndex) = 1 Then
                                    TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                End If
                            End If
                        End If
                        
                        If Map(mapnum).Events(id).Pages(page).chkVariable = 1 Then
                            Select Case Map(mapnum).Events(id).Pages(page).VariableCompare
                                Case 0
                                    If Player(I).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) <> Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 1
                                    If Player(I).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) < Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 2
                                    If Player(I).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) > Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 3
                                    If Player(I).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) <= Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 4
                                    If Player(I).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) >= Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 5
                                    If Player(I).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) = Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                    End If
                            End Select
                        End If
                        
                        If Map(mapnum).Events(id).Pages(page).chkSwitch = 1 Then
                            If Map(mapnum).Events(id).Pages(page).SwitchCompare = 1 Then
                                If Player(I).Switches(Map(mapnum).Events(id).Pages(page).SwitchIndex) = 1 Then
                                    TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                End If
                            Else
                                If Player(I).Switches(Map(mapnum).Events(id).Pages(page).SwitchIndex) = 0 Then
                                    TempPlayer(I).EventMap.EventPages(x).Visible = 0
                                End If
                            End If
                        End If
                        
                        If TempPlayer(I).EventMap.EventPages(x).Visible = 0 Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong SSpawnEvent
                            buffer.WriteLong id
                            With TempPlayer(I).EventMap.EventPages(x)
                                buffer.WriteString Map(GetPlayerMap(I)).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Name
                                buffer.WriteLong .Dir
                                buffer.WriteLong .GraphicNum
                                buffer.WriteLong .GraphicType
                                buffer.WriteLong .GraphicX
                                buffer.WriteLong .GraphicX2
                                buffer.WriteLong .GraphicY
                                buffer.WriteLong .GraphicY2
                                buffer.WriteLong .movementspeed
                                buffer.WriteLong .x
                                buffer.WriteLong .y
                                buffer.WriteLong .Position
                                buffer.WriteLong .Visible
                                buffer.WriteLong Map(mapnum).Events(id).Pages(page).WalkAnim
                                buffer.WriteLong Map(mapnum).Events(id).Pages(page).DirFix
                                buffer.WriteLong Map(mapnum).Events(id).Pages(page).WalkThrough
                                buffer.WriteLong Map(mapnum).Events(id).Pages(page).ShowName
                            End With
                            SendDataTo I, buffer.ToArray
                            Set buffer = Nothing
                        End If
                    End If
                End If
            Next
        End If
    Next
    
End Sub

Public Sub SpawnNewEvents()
Dim buffer As clsBuffer, pageID As Long, id As Long, compare As Long, I As Long, mapnum As Long, x As Long, z As Long, spawnevent As Boolean, p As Long
    'That was only removing events... now we gotta worry about spawning them again, luckily, it is almost the same exact thing, but backwards!
    For I = 1 To Player_HighIndex
        If TempPlayer(I).EventMap.CurrentEvents > 0 Then
            mapnum = GetPlayerMap(I)
            For x = 1 To TempPlayer(I).EventMap.CurrentEvents
                id = TempPlayer(I).EventMap.EventPages(x).eventID
                pageID = TempPlayer(I).EventMap.EventPages(x).pageID
                
                'See if there is any reason to delete this event....
                'In other words, go back through conditions and make sure they all check up.
                For z = Map(mapnum).Events(id).PageCount To 1 Step -1
                        
                    spawnevent = True
                        
                    If Map(mapnum).Events(id).Pages(z).chkHasItem = 1 Then
                        If HasItem(I, Map(mapnum).Events(id).Pages(z).HasItemIndex) = 0 Then
                            spawnevent = False
                        End If
                    End If
                        
                    If Map(mapnum).Events(id).Pages(z).chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(id).Pages(z).SelfSwitchCompare = 0 Then
                            compare = 1
                        Else
                            compare = 0
                        End If
                        If Map(mapnum).Events(id).Global = 1 Then
                            If Map(mapnum).Events(id).SelfSwitches(Map(mapnum).Events(id).Pages(z).SelfSwitchIndex) <> compare Then
                                spawnevent = False
                            End If
                        Else
                            If TempPlayer(I).EventMap.EventPages(id).SelfSwitches(Map(mapnum).Events(id).Pages(z).SelfSwitchIndex) <> compare Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If Map(mapnum).Events(id).Pages(z).chkVariable = 1 Then
                        Select Case Map(mapnum).Events(id).Pages(z).VariableCompare
                            Case 0
                                If Player(I).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) <> Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 1
                                If Player(I).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) < Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 2
                                If Player(I).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) > Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 3
                                If Player(I).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) <= Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 4
                                If Player(I).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) >= Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 5
                                If Player(I).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) = Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                        End Select
                    End If
                        
                    If Map(mapnum).Events(id).Pages(z).chkSwitch = 1 Then
                        If Map(mapnum).Events(id).Pages(z).SwitchCompare = 0 Then
                            If Player(I).Switches(Map(mapnum).Events(id).Pages(z).SwitchIndex) = 0 Then
                                spawnevent = False
                            End If
                        Else
                            If Player(I).Switches(Map(mapnum).Events(id).Pages(z).SwitchIndex) = 1 Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If spawnevent = True Then
                        If TempPlayer(I).EventMap.EventPages(x).Visible = 1 Then
                            If z <= pageID Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If spawnevent = True Then
                        With TempPlayer(I).EventMap.EventPages(x)
                            If Map(mapnum).Events(id).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(id).Pages(z).GraphicY
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
                            .GraphicNum = Map(mapnum).Events(id).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(id).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(id).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(id).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(id).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(id).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(id).Pages(z).MoveSpeed
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
                            .Position = Map(mapnum).Events(id).Pages(z).Position
                            .eventID = id
                            .pageID = z
                            .Visible = 1
                                
                            .MoveType = Map(mapnum).Events(id).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(id).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(id).Pages(z).MoveRouteCount)
                                For p = 0 To Map(mapnum).Events(id).Pages(z).MoveRouteCount
                                    .MoveRoute(p) = Map(mapnum).Events(id).Pages(z).MoveRoute(p)
                                Next
                            End If
                                
                            .RepeatMoveRoute = Map(mapnum).Events(id).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(id).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(id).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(id).Pages(z).MoveSpeed
                                
                            .WalkThrough = Map(mapnum).Events(id).Pages(z).WalkThrough
                            .ShowName = Map(mapnum).Events(id).Pages(z).ShowName
                            .WalkingAnim = Map(mapnum).Events(id).Pages(z).WalkAnim
                            .FixedDir = Map(mapnum).Events(id).Pages(z).DirFix
                            
                        End With
                        
                        
                        
                        Set buffer = New clsBuffer
                        buffer.WriteLong SSpawnEvent
                        buffer.WriteLong id
                        With TempPlayer(I).EventMap.EventPages(x)
                            buffer.WriteString Map(GetPlayerMap(I)).Events(.eventID).Name
                            buffer.WriteLong .Dir
                            buffer.WriteLong .GraphicNum
                            buffer.WriteLong .GraphicType
                            buffer.WriteLong .GraphicX
                            buffer.WriteLong .GraphicX2
                            buffer.WriteLong .GraphicY
                            buffer.WriteLong .GraphicY2
                            buffer.WriteLong .movementspeed
                            buffer.WriteLong .x
                            buffer.WriteLong .y
                            buffer.WriteLong .Position
                            buffer.WriteLong .Visible
                            buffer.WriteLong Map(mapnum).Events(id).Pages(z).WalkAnim
                            buffer.WriteLong Map(mapnum).Events(id).Pages(z).DirFix
                            buffer.WriteLong Map(mapnum).Events(id).Pages(z).WalkThrough
                            buffer.WriteLong Map(mapnum).Events(id).Pages(z).ShowName
                        End With
                        SendDataTo I, buffer.ToArray
                        Set buffer = Nothing
                        GoTo nextevent
                    End If
                Next
nextevent:
            Next
        End If
    Next
    
End Sub

Public Sub ProcessEventMovement()
Dim rand As Long, x As Long, I As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, mapnum As Long, actualmovespeed As Long, buffer As clsBuffer, z As Long, sendupdate As Boolean
    'Process Movement if needed for each player/each map/each event....
    For I = 1 To MAX_MAPS
        If PlayersOnMap(I) Then
            'Manage Global Events First, then all the others.....
            If TempEventMap(I).EventCount > 0 Then
                For x = 1 To TempEventMap(I).EventCount
                    If TempEventMap(I).Events(x).active = 1 Then
                        If TempEventMap(I).Events(x).MoveTimer <= GetTickCount Then
                            'Real event! Lets process it!
                            Select Case TempEventMap(I).Events(x).MoveType
                                Case 0
                                    'Nothing, fixed position
                                Case 1 'Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    If CanEventMove(0, I, TempEventMap(I).Events(x).x, TempEventMap(I).Events(x).y, x, TempEventMap(I).Events(x).WalkThrough, rand, True) Then
                                        Select Case TempEventMap(I).Events(x).MoveSpeed
                                            Case 0
                                                EventMove 0, I, x, rand, 2, True
                                            Case 1
                                                EventMove 0, I, x, rand, 3, True
                                            Case 2
                                                EventMove 0, I, x, rand, 4, True
                                            Case 3
                                                EventMove 0, I, x, rand, 6, True
                                            Case 4
                                                EventMove 0, I, x, rand, 12, True
                                            Case 5
                                                EventMove 0, I, x, rand, 24, True
                                        End Select
                                    Else
                                        EventDir 0, I, x, rand, True
                                    End If
                                Case 2 'Move Route - later
                                    With TempEventMap(I).Events(x)
                                        isglobal = True
                                        mapnum = I
                                        playerID = 0
                                        eventID = x
                                        WalkThrough = TempEventMap(I).Events(x).WalkThrough
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                GoTo donotprocessmoveroute
                                            End If
                                            .MoveRouteStep = .MoveRouteStep + 1
                                            Select Case .MoveSpeed
                                                Case 0
                                                    actualmovespeed = 2
                                                Case 1
                                                    actualmovespeed = 3
                                                Case 2
                                                    actualmovespeed = 4
                                                Case 3
                                                    actualmovespeed = 6
                                                Case 4
                                                    actualmovespeed = 12
                                                Case 5
                                                    actualmovespeed = 24
                                            End Select
                                            Select Case .MoveRoute(.MoveRouteStep).index
                                                Case 1
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_UP, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 2
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 3
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 4
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 5
                                                    z = Random(0, 3)
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 6
                                                    If isglobal = False Then
                                                        If IsOneBlockAway(.x, .y, GetPlayerX(playerID), GetPlayerY(playerID)) = True Then
                                                            EventDir playerID, GetPlayerMap(playerID), eventID, GetDirToPlayer(playerID, GetPlayerMap(playerID), eventID), False
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        Else
                                                            z = CanEventMoveTowardsPlayer(playerID, mapnum, eventID)
                                                            If z >= 4 Then
                                                                'No
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            Else
                                                                'i is the direct, lets go...
                                                                If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                    EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                                Else
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 7
                                                    If isglobal = False Then
                                                        z = CanEventMoveAwayFromPlayer(playerID, mapnum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            'i is the direct, lets go...
                                                            If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 8
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, .Dir, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, .Dir, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 9
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                    End Select
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 10
                                                    .MoveTimer = GetTickCount + 100
                                                Case 11
                                                    .MoveTimer = GetTickCount + 500
                                                Case 12
                                                    .MoveTimer = GetTickCount + 1000
                                                Case 13
                                                    EventDir playerID, mapnum, eventID, DIR_UP, isglobal
                                                Case 14
                                                    EventDir playerID, mapnum, eventID, DIR_DOWN, isglobal
                                                Case 15
                                                    EventDir playerID, mapnum, eventID, DIR_LEFT, isglobal
                                                Case 16
                                                    EventDir playerID, mapnum, eventID, DIR_RIGHT, isglobal
                                                Case 17
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_DOWN
                                                        Case DIR_LEFT
                                                            z = DIR_UP
                                                        Case DIR_DOWN
                                                            z = DIR_LEFT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 18
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_LEFT
                                                        Case DIR_RIGHT
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_RIGHT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 19
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 20
                                                    z = Random(0, 3)
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 21
                                                    If isglobal = False Then
                                                        z = GetDirToPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 22
                                                    If isglobal = False Then
                                                        z = GetDirAwayFromPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 23
                                                    .MoveSpeed = 0
                                                Case 24
                                                    .MoveSpeed = 1
                                                Case 25
                                                    .MoveSpeed = 2
                                                Case 26
                                                    .MoveSpeed = 3
                                                Case 27
                                                    .MoveSpeed = 4
                                                Case 28
                                                    .MoveSpeed = 5
                                                Case 29
                                                    .MoveFreq = 0
                                                Case 30
                                                    .MoveFreq = 1
                                                Case 31
                                                    .MoveFreq = 2
                                                Case 32
                                                    .MoveFreq = 3
                                                Case 33
                                                    .MoveFreq = 4
                                                Case 34
                                                    .WalkingAnim = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 35
                                                    .WalkingAnim = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 36
                                                    .FixedDir = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 37
                                                    .FixedDir = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 38
                                                    .WalkThrough = 1
                                                Case 39
                                                    .WalkThrough = 0
                                                Case 40
                                                    .Position = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 41
                                                    .Position = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 42
                                                    .Position = 2
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 43
                                                    .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                    .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                    .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                    .GraphicX2 = .MoveRoute(.MoveRouteStep).Data4
                                                    .GraphicY = .MoveRoute(.MoveRouteStep).Data5
                                                    .GraphicY2 = .MoveRoute(.MoveRouteStep).Data6
                                                    If .GraphicType = 1 Then
                                                        Select Case .GraphicY
                                                            Case 0
                                                                .Dir = DIR_DOWN
                                                            Case 1
                                                                .Dir = DIR_LEFT
                                                            Case 2
                                                                .Dir = DIR_RIGHT
                                                            Case 3
                                                                .Dir = DIR_UP
                                                        End Select
                                                    End If
                                                    'Need to Send Update to client
                                                    sendupdate = True
                                            End Select
                                            
                                            If sendupdate Then
                                                Set buffer = New clsBuffer
                                                buffer.WriteLong SSpawnEvent
                                                buffer.WriteLong eventID
                                                With TempEventMap(I).Events(x)
                                                    buffer.WriteString Map(GetPlayerMap(I)).Events(eventID).Name
                                                    buffer.WriteLong .Dir
                                                    buffer.WriteLong .GraphicNum
                                                    buffer.WriteLong .GraphicType
                                                    buffer.WriteLong .GraphicX
                                                    buffer.WriteLong .GraphicX2
                                                    buffer.WriteLong .GraphicY
                                                    buffer.WriteLong .GraphicY2
                                                    buffer.WriteLong .MoveSpeed
                                                    buffer.WriteLong .x
                                                    buffer.WriteLong .y
                                                    buffer.WriteLong .Position
                                                    buffer.WriteLong .active
                                                    buffer.WriteLong .WalkingAnim
                                                    buffer.WriteLong .FixedDir
                                                    buffer.WriteLong .WalkThrough
                                                    buffer.WriteLong .ShowName
                                                End With
                                                SendDataToMap I, buffer.ToArray
                                                Set buffer = Nothing
                                            End If
donotprocessmoveroute:
                                        End If
                                    End With
                            End Select
                            
                            Select Case TempEventMap(I).Events(x).MoveFreq
                                Case 0
                                    TempEventMap(I).Events(x).MoveTimer = GetTickCount + 4000
                                Case 1
                                    TempEventMap(I).Events(x).MoveTimer = GetTickCount + 2000
                                Case 2
                                    TempEventMap(I).Events(x).MoveTimer = GetTickCount + 1000
                                Case 3
                                    TempEventMap(I).Events(x).MoveTimer = GetTickCount + 500
                                Case 4
                                    TempEventMap(I).Events(x).MoveTimer = GetTickCount + 250
                            End Select
                        End If
                    End If
                Next
            End If
            'HOPEFULLY this will not reduce CPS too much!
        End If
        DoEvents
    Next
End Sub

Public Sub ProcessLocalEventMovement()
Dim rand As Long, x As Long, I As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, mapnum As Long, actualmovespeed As Long, buffer As clsBuffer, z As Long, sendupdate As Boolean
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            playerID = I
            If TempPlayer(I).EventMap.CurrentEvents > 0 Then
                For x = 1 To TempPlayer(I).EventMap.CurrentEvents
                    If Map(GetPlayerMap(I)).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Global = 1 Then GoTo nextevent1
                    If TempPlayer(I).EventMap.EventPages(x).Visible = 1 Then
                        If TempPlayer(I).EventMap.EventPages(x).MoveTimer <= GetTickCount Then
                            'Real event! Lets process it!
                            Select Case TempPlayer(I).EventMap.EventPages(x).MoveType
                                Case 0
                                    'Nothing, fixed position
                                Case 1 'Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    playerID = I
                                    If CanEventMove(I, GetPlayerMap(I), TempPlayer(I).EventMap.EventPages(x).x, TempPlayer(I).EventMap.EventPages(x).y, x, TempPlayer(I).EventMap.EventPages(x).WalkThrough, rand, False) Then
                                        Select Case TempPlayer(I).EventMap.EventPages(x).MoveSpeed
                                            Case 0
                                                EventMove I, GetPlayerMap(I), x, rand, 2, False
                                            Case 1
                                                EventMove I, GetPlayerMap(I), x, rand, 3, False
                                            Case 2
                                                EventMove I, GetPlayerMap(I), x, rand, 4, False
                                            Case 3
                                                EventMove I, GetPlayerMap(I), x, rand, 6, False
                                            Case 4
                                                EventMove I, GetPlayerMap(I), x, rand, 12, False
                                            Case 5
                                                EventMove I, GetPlayerMap(I), x, rand, 24, False
                                        End Select
                                    Else
                                        EventDir 0, GetPlayerMap(I), x, rand, True
                                    End If
                                Case 2 'Move Route - later!
                                    With TempPlayer(I).EventMap.EventPages(x)
                                        isglobal = False
                                        mapnum = GetPlayerMap(I)
                                        playerID = I
                                        eventID = .eventID
                                        WalkThrough = .WalkThrough
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                GoTo donotprocessmoveroute1
                                            End If
                                            .MoveRouteStep = .MoveRouteStep + 1
                                            Select Case .MoveSpeed
                                                Case 0
                                                    actualmovespeed = 2
                                                Case 1
                                                    actualmovespeed = 3
                                                Case 2
                                                    actualmovespeed = 4
                                                Case 3
                                                    actualmovespeed = 6
                                                Case 4
                                                    actualmovespeed = 12
                                                Case 5
                                                    actualmovespeed = 24
                                            End Select
                                            Select Case .MoveRoute(.MoveRouteStep).index
                                                Case 1
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_UP, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 2
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 3
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 4
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 5
                                                    z = Random(0, 3)
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 6
                                                    If isglobal = False Then
                                                        If IsOneBlockAway(.x, .y, GetPlayerX(playerID), GetPlayerY(playerID)) = True Then
                                                            EventDir playerID, GetPlayerMap(playerID), eventID, GetDirToPlayer(playerID, GetPlayerMap(playerID), eventID), False
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        Else
                                                            z = CanEventMoveTowardsPlayer(playerID, mapnum, eventID)
                                                            If z >= 4 Then
                                                                'No
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            Else
                                                                'i is the direct, lets go...
                                                                If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                    EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                                Else
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 7
                                                    If isglobal = False Then
                                                        z = CanEventMoveAwayFromPlayer(playerID, mapnum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            'i is the direct, lets go...
                                                            If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 8
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, .Dir, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, .Dir, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 9
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                    End Select
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 10
                                                    .MoveTimer = GetTickCount + 100
                                                Case 11
                                                    .MoveTimer = GetTickCount + 500
                                                Case 12
                                                    .MoveTimer = GetTickCount + 1000
                                                Case 13
                                                    EventDir playerID, mapnum, eventID, DIR_UP, isglobal
                                                Case 14
                                                    EventDir playerID, mapnum, eventID, DIR_DOWN, isglobal
                                                Case 15
                                                    EventDir playerID, mapnum, eventID, DIR_LEFT, isglobal
                                                Case 16
                                                    EventDir playerID, mapnum, eventID, DIR_RIGHT, isglobal
                                                Case 17
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_DOWN
                                                        Case DIR_LEFT
                                                            z = DIR_UP
                                                        Case DIR_DOWN
                                                            z = DIR_LEFT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 18
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_LEFT
                                                        Case DIR_RIGHT
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_RIGHT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 19
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 20
                                                    z = Random(0, 3)
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 21
                                                    If isglobal = False Then
                                                        z = GetDirToPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 22
                                                    If isglobal = False Then
                                                        z = GetDirAwayFromPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 23
                                                    .MoveSpeed = 0
                                                Case 24
                                                    .MoveSpeed = 1
                                                Case 25
                                                    .MoveSpeed = 2
                                                Case 26
                                                    .MoveSpeed = 3
                                                Case 27
                                                    .MoveSpeed = 4
                                                Case 28
                                                    .MoveSpeed = 5
                                                Case 29
                                                    .MoveFreq = 0
                                                Case 30
                                                    .MoveFreq = 1
                                                Case 31
                                                    .MoveFreq = 2
                                                Case 32
                                                    .MoveFreq = 3
                                                Case 33
                                                    .MoveFreq = 4
                                                Case 34
                                                    .WalkingAnim = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 35
                                                    .WalkingAnim = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 36
                                                    .FixedDir = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 37
                                                    .FixedDir = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 38
                                                    .WalkThrough = 1
                                                Case 39
                                                    .WalkThrough = 0
                                                Case 40
                                                    .Position = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 41
                                                    .Position = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 42
                                                    .Position = 2
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 43
                                                    .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                    .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                    .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                    .GraphicX2 = .MoveRoute(.MoveRouteStep).Data4
                                                    .GraphicY = .MoveRoute(.MoveRouteStep).Data5
                                                    .GraphicY2 = .MoveRoute(.MoveRouteStep).Data6
                                                    If .GraphicType = 1 Then
                                                        Select Case .GraphicY
                                                            Case 0
                                                                .Dir = DIR_DOWN
                                                            Case 1
                                                                .Dir = DIR_LEFT
                                                            Case 2
                                                                .Dir = DIR_RIGHT
                                                            Case 3
                                                                .Dir = DIR_UP
                                                        End Select
                                                    End If
                                                    'Need to Send Update to client
                                                    sendupdate = True
                                            End Select
                                            
                                            If sendupdate Then
                                                Set buffer = New clsBuffer
                                                buffer.WriteLong SSpawnEvent
                                                buffer.WriteLong eventID
                                                With TempPlayer(playerID).EventMap.EventPages(eventID)
                                                    buffer.WriteString Map(GetPlayerMap(playerID)).Events(eventID).Name
                                                    buffer.WriteLong .Dir
                                                    buffer.WriteLong .GraphicNum
                                                    buffer.WriteLong .GraphicType
                                                    buffer.WriteLong .GraphicX
                                                    buffer.WriteLong .GraphicX2
                                                    buffer.WriteLong .GraphicY
                                                    buffer.WriteLong .GraphicY2
                                                    buffer.WriteLong .MoveSpeed
                                                    buffer.WriteLong .x
                                                    buffer.WriteLong .y
                                                    buffer.WriteLong .Position
                                                    buffer.WriteLong .Visible
                                                    buffer.WriteLong .WalkingAnim
                                                    buffer.WriteLong .FixedDir
                                                    buffer.WriteLong .WalkThrough
                                                    buffer.WriteLong .ShowName
                                                End With
                                                SendDataTo playerID, buffer.ToArray
                                                Set buffer = Nothing
                                            End If
                                        End If
                                    End With
                            End Select

donotprocessmoveroute1:
                            Select Case TempPlayer(playerID).EventMap.EventPages(x).MoveFreq
                                Case 0
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 4000
                                Case 1
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 2000
                                Case 2
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 1000
                                Case 3
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 500
                                Case 4
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 250
                            End Select
                        End If
                    End If
nextevent1:
                Next
            End If
        End If
        DoEvents
    Next
End Sub

Public Sub ProcessEventCommands()
Dim buffer As clsBuffer, I As Long, x As Long, z As Long, removeEventProcess As Boolean, w As Long, v As Long, p As Long
Dim II As Long, Amnt As Long, wgCon As Boolean
    'Now, we process the damn things for commands :P
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            For x = 1 To TempPlayer(I).EventMap.CurrentEvents
                If TempPlayer(I).EventMap.EventPages(x).Visible Then
                    If Map(Player(I).Map).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Pages(TempPlayer(I).EventMap.EventPages(x).pageID).Trigger = 2 Then 'Parallel Process baby!
                        If TempPlayer(I).EventProcessingCount > 0 Then
                            For z = 1 To TempPlayer(I).EventProcessingCount
                                If TempPlayer(I).EventProcessing(z).eventID = TempPlayer(I).EventMap.EventPages(x).eventID And TempPlayer(I).EventMap.EventPages(x).pageID = TempPlayer(I).EventProcessing(z).pageID Then
                                    'Exit For
                                Else
                                    If z = TempPlayer(I).EventProcessingCount Then
                                        If Map(GetPlayerMap(I)).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Pages(TempPlayer(I).EventMap.EventPages(x).pageID).CommandListCount > 0 Then
                                            'start new event processing
                                            TempPlayer(I).EventProcessingCount = TempPlayer(I).EventProcessingCount + 1
                                            ReDim Preserve TempPlayer(I).EventProcessing(TempPlayer(I).EventProcessingCount)
                                            With TempPlayer(I).EventProcessing(TempPlayer(I).EventProcessingCount)
                                                .ActionTimer = GetTickCount
                                                .CurList = 1
                                                .CurSlot = 1
                                                .eventID = TempPlayer(I).EventMap.EventPages(x).eventID
                                                .pageID = TempPlayer(I).EventMap.EventPages(x).pageID
                                                .WaitingForResponse = 0
                                                ReDim .ListLeftOff(0 To Map(GetPlayerMap(I)).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Pages(TempPlayer(I).EventMap.EventPages(x).pageID).CommandListCount)
                                            End With
                                        End If
                                        Exit For
                                        
                                    End If
                                End If
                            Next
                        Else
                            If Map(GetPlayerMap(I)).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Pages(TempPlayer(I).EventMap.EventPages(x).pageID).CommandListCount > 0 Then
                                'Clearly need to start it!
                                TempPlayer(I).EventProcessingCount = 1
                                ReDim Preserve TempPlayer(I).EventProcessing(TempPlayer(I).EventProcessingCount)
                                With TempPlayer(I).EventProcessing(TempPlayer(I).EventProcessingCount)
                                    .ActionTimer = GetTickCount
                                    .CurList = 1
                                    .CurSlot = 1
                                    .eventID = TempPlayer(I).EventMap.EventPages(x).eventID
                                    .pageID = TempPlayer(I).EventMap.EventPages(x).pageID
                                    .WaitingForResponse = 0
                                    ReDim .ListLeftOff(0 To Map(GetPlayerMap(I)).Events(TempPlayer(I).EventMap.EventPages(x).eventID).Pages(TempPlayer(I).EventMap.EventPages(x).pageID).CommandListCount)
                                End With
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    'That is it for starting parallel processes :D now we just have to make the code that actually processes the events to their fullest
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            If TempPlayer(I).EventProcessingCount > 0 Then
restartloop:
                For x = 1 To TempPlayer(I).EventProcessingCount
                    With TempPlayer(I).EventProcessing(x)
                        If TempPlayer(I).EventProcessingCount = 0 Then Exit Sub
                        removeEventProcess = False
                        If .WaitingForResponse = 2 Then
                            If TempPlayer(I).InShop = 0 Then
                                .WaitingForResponse = 0
                            End If
                        End If
                        If .WaitingForResponse = 3 Then
                            If TempPlayer(I).InBank = False Then
                                .WaitingForResponse = 0
                            End If
                        End If
                        If .WaitingForResponse = 0 Then
                            If .ActionTimer <= GetTickCount Then
restartlist:
                                If .ListLeftOff(.CurList) > 0 Then
                                    .CurSlot = .ListLeftOff(.CurList) + 1
                                End If
                                If .CurList > Map(Player(I).Map).Events(.eventID).Pages(.pageID).CommandListCount Then
                                    'Get rid of this event, it is bad
                                    removeEventProcess = True
                                    GoTo endprocess
                                End If
                                If .CurSlot > Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount Then
                                    If .CurList = 1 Then
                                        'Get rid of this event, it is bad
                                        removeEventProcess = True
                                        GoTo endprocess
                                    Else
                                        .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).ParentList
                                        .CurSlot = 1
                                        GoTo restartlist
                                    End If
                                End If
                                'If we are still here, then we are good to process shit :D
                                Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).index
                                    Case EventType.evAddText
                                        Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                            Case 0
                                                PlayerMsg I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case 1
                                                MapMsg GetPlayerMap(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case 2
                                                GlobalMsg Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        End Select
                                    Case EventType.evShowText
                                        Set buffer = New clsBuffer
                                        buffer.WriteLong SEventChat
                                        buffer.WriteLong .eventID
                                        buffer.WriteLong .pageID
                                        buffer.WriteString ParseEventText(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1)
                                        buffer.WriteLong 0
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                            If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowText Or Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowChoices Then
                                                buffer.WriteLong 1
                                            ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evCondition Then
                                                buffer.WriteLong 2
                                            Else
                                                buffer.WriteLong 0
                                            End If
                                        Else
                                            buffer.WriteLong 2
                                        End If
                                        SendDataTo I, buffer.ToArray
                                        Set buffer = Nothing
                                        .WaitingForResponse = 1
                                    Case EventType.evShowChoices
                                        Set buffer = New clsBuffer
                                        buffer.WriteLong SEventChat
                                        buffer.WriteLong .eventID
                                        buffer.WriteLong .pageID
                                        buffer.WriteString ParseEventText(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1)
                                        If Len(Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text2)) > 0 Then
                                            w = 1
                                            If Len(Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text3)) > 0 Then
                                                w = 2
                                                If Len(Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text4)) > 0 Then
                                                    w = 3
                                                    If Len(Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text5)) > 0 Then
                                                        w = 4
                                                    End If
                                                End If
                                            End If
                                        End If
                                        buffer.WriteLong w
                                        For v = 1 To w
                                            Select Case v
                                                Case 1
                                                    buffer.WriteString ParseEventText(I, Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text2))
                                                Case 2
                                                    buffer.WriteString ParseEventText(I, Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text3))
                                                Case 3
                                                    buffer.WriteString ParseEventText(I, Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text4))
                                                Case 4
                                                    buffer.WriteString ParseEventText(I, Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text5))
                                            End Select
                                        Next
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                            If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowText Or Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowChoices Then
                                                buffer.WriteLong 1
                                            ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evCondition Then
                                                buffer.WriteLong 2
                                            Else
                                                buffer.WriteLong 0
                                            End If
                                        Else
                                            buffer.WriteLong 2
                                        End If
                                        SendDataTo I, buffer.ToArray
                                        Set buffer = Nothing
                                        .WaitingForResponse = 1
                                    Case EventType.evPlayerVar
                                        Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                            Case 0
                                                Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                            Case 1
                                                Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) + Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                            Case 2
                                                Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) - Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                            Case 3
                                                Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Random(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4)
                                        End Select
                                    Case EventType.evPlayerSwitch
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            Player(I).Switches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 1
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            Player(I).Switches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 0
                                        End If
                                    Case EventType.evSelfSwitch
                                        If Map(GetPlayerMap(I)).Events(.eventID).Global = 1 Then
                                            If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                Map(GetPlayerMap(I)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                            ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                Map(GetPlayerMap(I)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                            End If
                                        Else
                                            If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                TempPlayer(I).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                            ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                TempPlayer(I).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                            End If
                                        End If
                                    Case EventType.evCondition
                                        Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Condition
                                            Case 0 'Player Var
                                                Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1 'Player Switch
                                                        If Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) >= Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 2 'Has Item
                                                        If Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <= Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 3 'Class Is
                                                        If Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) > Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 4 'Learnt Skill
                                                        If Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) < Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 5 'Level Is
                                                        If Player(I).Variables(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <> Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 1
                                                Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If Player(I).Switches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If Player(I).Switches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 0 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 2
                                                If HasItems(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2) Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 3
                                                If Player(I).Class = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 4
                                                If HasSpell(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = True Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 5
                                                Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If GetPlayerLevel(I) = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If GetPlayerLevel(I) >= Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 2
                                                        If GetPlayerLevel(I) <= Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 3
                                                        If GetPlayerLevel(I) > Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 4
                                                        If GetPlayerLevel(I) < Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 5
                                                        If GetPlayerLevel(I) <> Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 6
                                                If Map(GetPlayerMap(I)).Events(.eventID).Global = 1 Then
                                                    Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                        Case 0 'Self Switch is true
                                                            If Map(GetPlayerMap(I)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                        Case 1  'self switch is false
                                                            If Map(GetPlayerMap(I)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                    End Select
                                                Else
                                                    Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                        Case 0 'Self Switch is true
                                                            If TempPlayer(I).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                        Case 1  'self switch is false
                                                            If TempPlayer(I).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                    End Select
                                                End If
                                            Case 7 'Has skill level
                                                If GetPlayerSkillLevel(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) >= Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2 Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 8 'Check quest status
                                                Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                
                                                    Case QUEST_NOT_STARTED
                                                        If Player(I).PlayerQuest(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).Status = QUEST_NOT_STARTED Then
                                                            wgCon = True
                                                        End If
                                                    Case QUEST_STARTED
                                                        If Player(I).PlayerQuest(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).Status = QUEST_STARTED Then
                                                            wgCon = True
                                                        End If
                                                    Case QUEST_COMPLETED
                                                        If Player(I).PlayerQuest(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).Status = QUEST_COMPLETED Or Player(I).PlayerQuest(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).Status = QUEST_COMPLETED_BUT Then
                                                            wgCon = True
                                                        End If
                                                    
                                                End Select
                                                
                                                If wgCon Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                        End Select
                                        GoTo endprocess
                                    Case EventType.evExitProcess
                                        removeEventProcess = True
                                        GoTo endprocess
                                    Case EventType.evChangeItems
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            If FindItem(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                                Call SetPlayerInvItemValue(I, FindItem(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3)
                                            End If
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If Item(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Type <> ITEM_TYPE_CURRENCY Then
                                                For II = 1 To Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                    GiveInvItem I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, True
                                                Next II
                                            Else
                                                Amnt = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                GiveInvItem I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, True
                                                SendActionMsg GetPlayerMap(I), "+ " & Amnt & "g", Green, 1, (GetPlayerX(I) * 32), (GetPlayerY(I) * 32)
                                            End If
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                            ' Check Amount
                                            If HasItem(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                                If Item(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Type <> ITEM_TYPE_CURRENCY Then
                                                    For II = 1 To Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                        TakeInvItem I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                    Next II
                                                Else
                                                    Amnt = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                    TakeInvItem I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                    SendActionMsg GetPlayerMap(I), "- " & Amnt & "g", Red, 1, (GetPlayerX(I) * 32), (GetPlayerY(I) * 32)
                                                End If
                                            End If
                                        End If
                                        SendInventory I
                                    Case EventType.evRestoreHP
                                        SetPlayerVital I, HP, GetPlayerMaxVital(I, HP)
                                        SendVital I, HP
                                    Case EventType.evRestoreMP
                                        SetPlayerVital I, MP, GetPlayerMaxVital(I, MP)
                                        SendVital I, MP
                                    Case EventType.evLevelUp
                                        SetPlayerExp I, GetPlayerNextLevel(I)
                                        CheckPlayerLevelUp I
                                        SendEXP I
                                        SendPlayerData I
                                    Case EventType.evChangeLevel
                                        SetPlayerLevel I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SetPlayerExp I, 0
                                        SendEXP I
                                        SendPlayerData I
                                    Case EventType.evChangeSkills
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            If FindOpenSpellSlot(I) > 0 Then
                                                If HasSpell(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = False Then
                                                    SetPlayerSpell I, FindOpenSpellSlot(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                Else
                                                    'Error, already knows spell
                                                End If
                                            Else
                                                'Error, no room for spells
                                            End If
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If HasSpell(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = True Then
                                                For p = 1 To MAX_PLAYER_SPELLS
                                                    If Player(I).Spell(p) = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 Then
                                                        SetPlayerSpell I, p, 0
                                                    End If
                                                Next
                                            End If
                                        End If
                                        SendPlayerSpells I
                                    Case EventType.evChangeClass
                                        Player(I).Class = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData I
                                    Case EventType.evChangeSprite
                                        Player(I).Sprite = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData I
                                    Case EventType.evChangeSex
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                            Player(I).Sex = SEX_MALE
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                            Player(I).Sex = SEX_FEMALE
                                        End If
                                        SendPlayerData I
                                    Case EventType.evChangePK
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                            Player(I).PK = NO
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                            Player(I).PK = YES
                                        End If
                                        SendPlayerData I
                                    Case EventType.evWarpPlayer
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4 = 0 Then
                                            PlayerWarp I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        Else
                                            Player(I).Dir = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4 - 1
                                            PlayerWarp I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        End If
                                        
                                    Case EventType.evSetMoveRoute
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 <= Map(GetPlayerMap(I)).EventCount Then
                                            If Map(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Global = 1 Then
                                                TempEventMap(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                TempEventMap(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                TempEventMap(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                TempEventMap(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                TempEventMap(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                TempEventMap(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                            Else
                                                TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                            End If
                                        End If
                                    Case EventType.evPlayAnimation
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            SendAnimation GetPlayerMap(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerX(I), GetPlayerY(I), TARGET_TYPE_PLAYER, I, I
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If Map(GetPlayerMap(I)).Events(.eventID).Global = 1 Then
                                                SendAnimation GetPlayerMap(I), Map(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).x, Map(GetPlayerMap(I)).Events(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).y
                                            Else
                                                SendAnimation GetPlayerMap(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).x, TempPlayer(I).EventMap.EventPages(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).y, 0, 0, I
                                            End If
                                        ElseIf Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                            SendAnimation GetPlayerMap(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4, 0, 0, I
                                        End If
                                    Case EventType.evCustomScript
                                        'Runs Through Cases for a script
                                        Call CustomScript(I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                    Case EventType.evPlayBGM
                                        Set buffer = New clsBuffer
                                        buffer.WriteLong SPlayBGM
                                        buffer.WriteString Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        SendDataTo I, buffer.ToArray
                                        Set buffer = Nothing
                                    Case EventType.evFadeoutBGM
                                        Set buffer = New clsBuffer
                                        buffer.WriteLong SFadeoutBGM
                                        SendDataTo I, buffer.ToArray
                                        Set buffer = Nothing
                                    Case EventType.evPlaySound
                                        Set buffer = New clsBuffer
                                        buffer.WriteLong SPlaySound
                                        buffer.WriteString Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        SendDataTo I, buffer.ToArray
                                        Set buffer = Nothing
                                    Case EventType.evStopSound
                                        Set buffer = New clsBuffer
                                        buffer.WriteLong SStopSound
                                        SendDataTo I, buffer.ToArray
                                        Set buffer = Nothing
                                    Case EventType.evSetAccess
                                        Player(I).Access = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData I
                                    Case EventType.evOpenShop
                                        If Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 > 0 Then ' shop exists?
                                            If Len(Trim$(Shop(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Name)) > 0 Then ' name exists?
                                                SendOpenShop I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                TempPlayer(I).InShop = Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 ' stops movement and the like
                                                .WaitingForResponse = 2
                                            End If
                                        End If
                                    Case EventType.evOpenBank
                                        SendBank I
                                        TempPlayer(I).InBank = True
                                        .WaitingForResponse = 3
                                    Case EventType.evGiveExp
                                        If Abs(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2) = 1 Then
                                            GivePlayerEXP I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Call SendPlayerData(I)
                                        Else
                                            SetPlayerSkillExp I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Call SendPlayerData(I)
                                        End If
                                    Case EventType.evShowChatBubble
                                        Select Case Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case TARGET_TYPE_PLAYER
                                                SendChatBubble GetPlayerMap(I), I, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                            Case TARGET_TYPE_NPC
                                                SendChatBubble GetPlayerMap(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                            Case TARGET_TYPE_EVENT
                                                SendChatBubble GetPlayerMap(I), Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                        End Select
                                    Case EventType.evLabel
                                        'Do nothing, just a label
                                    Case EventType.evGotoLabel
                                        'Find the label's list of commands and slot
                                        FindEventLabel Trim$(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1), GetPlayerMap(I), .eventID, .pageID, .CurSlot, .CurList, .ListLeftOff
                                    Case EventType.evSpawnNpc
                                        If Map(GetPlayerMap(I)).NPC(Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                            SpawnNpc Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerMap(I), True
                                        End If
                                    Case EventType.evFadeIn
                                        SendSpecialEffect I, EFFECT_TYPE_FADEIN
                                    Case EventType.evFadeOut
                                        SendSpecialEffect I, EFFECT_TYPE_FADEOUT
                                    Case EventType.evFlashWhite
                                        SendSpecialEffect I, EFFECT_TYPE_FLASH
                                    Case EventType.evSetFog
                                        SendSpecialEffect I, EFFECT_TYPE_FOG, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                    Case EventType.evSetweather
                                        SendSpecialEffect I, EFFECT_TYPE_WEATHER, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                    Case EventType.evSetTint
                                        SendSpecialEffect I, EFFECT_TYPE_TINT, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4
                                    Case EventType.evWait
                                        .ActionTimer = GetTickCount + Map(GetPlayerMap(I)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                End Select
                                .CurSlot = .CurSlot + 1
                            End If
                        End If
endprocess:
                    End With
                    If removeEventProcess = True Then
                        TempPlayer(I).EventProcessingCount = TempPlayer(I).EventProcessingCount - 1
                        If TempPlayer(I).EventProcessingCount <= 0 Then
                            ReDim TempPlayer(I).EventProcessing(0)
                            GoTo restartloop:
                        Else
                            For z = x To TempPlayer(I).EventProcessingCount - 1
                                TempPlayer(I).EventProcessing(x) = TempPlayer(I).EventProcessing(x + 1)
                            Next
                            ReDim Preserve TempPlayer(I).EventProcessing(TempPlayer(I).EventProcessingCount)
                            GoTo restartloop
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub

Public Sub UpdateEventLogic()
Dim I As Long, x As Long, y As Long, z As Long, mapnum As Long, id As Long
Dim page As Long, buffer As clsBuffer, spawnevent As Boolean, p As Long, rand As Long, isglobal As Boolean, actualmovespeed As Long, playerID As Long, WalkThrough As Long, eventID As Long, sendupdate As Boolean, removeEventProcess As Boolean, w As Long, v As Long
    'Check Removing and Adding of Events (Did switches change or something?)
    RemoveDeadEvents
    SpawnNewEvents
    ProcessEventMovement
    ProcessLocalEventMovement
    ProcessEventCommands
    
End Sub

Sub SendSwitchesAndVariables(index As Long, Optional everyone As Boolean = False)
Dim buffer As clsBuffer, I As Long

Set buffer = New clsBuffer
buffer.WriteLong SSwitchesAndVariables

For I = 1 To MAX_SWITCHES
    buffer.WriteString Switches(I)
Next

For I = 1 To MAX_VARIABLES
    buffer.WriteString Variables(I)
Next

If everyone Then
    SendDataToAll buffer.ToArray
Else
    SendDataTo index, buffer.ToArray
End If

Set buffer = Nothing
End Sub

Sub SendMapEventData(index As Long)
Dim buffer As clsBuffer, I As Long, x As Long, y As Long, z As Long, mapnum As Long, w As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SMapEventData
    mapnum = GetPlayerMap(index)
    'Event Data
    buffer.WriteLong Map(mapnum).EventCount
        
    If Map(mapnum).EventCount > 0 Then
        For I = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(I)
                buffer.WriteString .Name
                buffer.WriteLong .Global
                buffer.WriteLong .x
                buffer.WriteLong .y
                buffer.WriteLong .PageCount
            End With
            If Map(mapnum).Events(I).PageCount > 0 Then
                For x = 1 To Map(mapnum).Events(I).PageCount
                    With Map(mapnum).Events(I).Pages(x)
                        buffer.WriteLong .chkVariable
                        buffer.WriteLong .VariableIndex
                        buffer.WriteLong .VariableCondition
                        buffer.WriteLong .VariableCompare
                            
                        buffer.WriteLong .chkSwitch
                        buffer.WriteLong .SwitchIndex
                        buffer.WriteLong .SwitchCompare
                        
                        buffer.WriteLong .chkHasItem
                        buffer.WriteLong .HasItemIndex
                        buffer.WriteLong .HasItemAmount
                            
                        buffer.WriteLong .chkSelfSwitch
                        buffer.WriteLong .SelfSwitchIndex
                        buffer.WriteLong .SelfSwitchCompare
                            
                        buffer.WriteLong .GraphicType
                        buffer.WriteLong .Graphic
                        buffer.WriteLong .GraphicX
                        buffer.WriteLong .GraphicY
                        buffer.WriteLong .GraphicX2
                        buffer.WriteLong .GraphicY2
                        
                        buffer.WriteLong .MoveType
                        buffer.WriteLong .MoveSpeed
                        buffer.WriteLong .MoveFreq
                        buffer.WriteLong .MoveRouteCount
                        
                        buffer.WriteLong .IgnoreMoveRoute
                        buffer.WriteLong .RepeatMoveRoute
                            
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                buffer.WriteLong .MoveRoute(y).index
                                buffer.WriteLong .MoveRoute(y).Data1
                                buffer.WriteLong .MoveRoute(y).Data2
                                buffer.WriteLong .MoveRoute(y).Data3
                                buffer.WriteLong .MoveRoute(y).Data4
                                buffer.WriteLong .MoveRoute(y).Data5
                                buffer.WriteLong .MoveRoute(y).Data6
                            Next
                        End If
                            
                        buffer.WriteLong .WalkAnim
                        buffer.WriteLong .DirFix
                        buffer.WriteLong .WalkThrough
                        buffer.WriteLong .ShowName
                        buffer.WriteLong .Trigger
                        buffer.WriteLong .CommandListCount
                        
                        buffer.WriteLong .Position
                    End With
                        
                    If Map(mapnum).Events(I).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map(mapnum).Events(I).Pages(x).CommandListCount
                            buffer.WriteLong Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount
                            buffer.WriteLong Map(mapnum).Events(I).Pages(x).CommandList(y).ParentList
                            If Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount > 0 Then
                                For z = 1 To Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(I).Pages(x).CommandList(y).Commands(z)
                                        buffer.WriteLong .index
                                        buffer.WriteString .Text1
                                        buffer.WriteString .Text2
                                        buffer.WriteString .Text3
                                        buffer.WriteString .Text4
                                        buffer.WriteString .Text5
                                        buffer.WriteLong .Data1
                                        buffer.WriteLong .Data2
                                        buffer.WriteLong .Data3
                                        buffer.WriteLong .Data4
                                        buffer.WriteLong .Data5
                                        buffer.WriteLong .Data6
                                        buffer.WriteLong .ConditionalBranch.CommandList
                                        buffer.WriteLong .ConditionalBranch.Condition
                                        buffer.WriteLong .ConditionalBranch.Data1
                                        buffer.WriteLong .ConditionalBranch.Data2
                                        buffer.WriteLong .ConditionalBranch.Data3
                                        buffer.WriteLong .ConditionalBranch.ElseCommandList
                                        buffer.WriteLong .MoveRouteCount
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                buffer.WriteLong .MoveRoute(w).index
                                                buffer.WriteLong .MoveRoute(w).Data1
                                                buffer.WriteLong .MoveRoute(w).Data2
                                                buffer.WriteLong .MoveRoute(w).Data3
                                                buffer.WriteLong .MoveRoute(w).Data4
                                                buffer.WriteLong .MoveRoute(w).Data5
                                                buffer.WriteLong .MoveRoute(w).Data6
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    'End Event Data
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
    Call SendSwitchesAndVariables(index)
End Sub

Function ParseEventText(ByVal index As Long, ByVal txt As String) As String
Dim I As Long, x As Long, newtxt As String, parsestring As String, z As Long
    txt = Replace(txt, "/name", Trim$(Player(index).Name))
    txt = Replace(txt, "/p", Trim$(Player(index).Name))
    Do While InStr(1, txt, "/v") > 0
        x = InStr(1, txt, "/v")
        If x > 0 Then
            I = 0
            Do Until IsNumeric(Mid(txt, x + 2 + I, 1)) = False
                I = I + 1
            Loop
            newtxt = Mid(txt, 1, x - 1)
            parsestring = Mid(txt, x + 2, I)
            z = Player(index).Variables(Val(parsestring))
            newtxt = newtxt & CStr(z)
            newtxt = newtxt & Mid(txt, x + 2 + I, Len(txt) - (x + I))
            txt = newtxt
        End If
    Loop
    ParseEventText = txt
End Function

Function FindEventLabel(ByVal Label As String, mapnum As Long, eventID As Long, pageID As Long, CurSlot As Long, CurList As Long, ListLeftOff() As Long)
Dim Stack() As Long, StackCount As Long, tmpCurSlot As Long, tmpCurList As Long, CurrentListOption() As Long
Dim removeEventProcess As Boolean, tmpListLeftOff() As Long, restartlist As Boolean, w As Long
    'Store the Old data, just in case
    tmpCurSlot = CurSlot
    tmpCurList = CurList
    tmpListLeftOff = ListLeftOff
    
    ReDim ListLeftOff(Map(mapnum).Events(eventID).Pages(pageID).CommandListCount)
    ReDim CurrentListOption(Map(mapnum).Events(eventID).Pages(pageID).CommandListCount)
    CurList = 1
    CurSlot = 1
    
    Do Until removeEventProcess = True
        If ListLeftOff(CurList) > 0 Then
            CurSlot = ListLeftOff(CurList)
            ListLeftOff(CurList) = 0
        End If
        If CurList > Map(mapnum).Events(eventID).Pages(pageID).CommandListCount Then
            'Get rid of this event, it is bad
            removeEventProcess = True
        End If
        
        If CurSlot > Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).CommandCount Then
            If CurList = 1 Then
                removeEventProcess = True
            Else
                CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).ParentList
                CurSlot = 1
                restartlist = True
            End If
        End If
        
        If restartlist = False Then
            If removeEventProcess = False Then
                'If we are still here, then we are good to process shit :D
                Select Case Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).index
                    Case EventType.evShowChoices
                        If Len(Trim$(Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text2)) > 0 Then
                            w = 1
                            If Len(Trim$(Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text3)) > 0 Then
                                w = 2
                                If Len(Trim$(Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text4)) > 0 Then
                                    w = 3
                                    If Len(Trim$(Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text5)) > 0 Then
                                        w = 4
                                    End If
                                End If
                            End If
                        End If
                        If w > 0 Then
                            If CurrentListOption(CurList) < w Then
                                CurrentListOption(CurList) = CurrentListOption(CurList) + 1
                                'Process
                                ListLeftOff(CurList) = CurSlot
                                Select Case CurrentListOption(CurList)
                                    Case 1
                                        CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data1
                                    Case 2
                                        CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data2
                                    Case 3
                                        CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data3
                                    Case 4
                                        CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data4
                                End Select
                                CurSlot = 0
                            Else
                                CurrentListOption(CurList) = 0
                                'continue on
                            End If
                        End If
                        w = 0
                    Case EventType.evCondition
                        If CurrentListOption(CurList) = 0 Then
                            CurrentListOption(CurList) = 1
                            ListLeftOff(CurList) = CurSlot
                            CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).ConditionalBranch.CommandList
                            CurSlot = 0
                        ElseIf CurrentListOption(CurList) = 1 Then
                            CurrentListOption(CurList) = 2
                            ListLeftOff(CurList) = CurSlot
                            CurList = Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).ConditionalBranch.ElseCommandList
                            CurSlot = 0
                        ElseIf CurrentListOption(CurList) = 2 Then
                            CurrentListOption(CurList) = 0
                        End If
                    Case EventType.evLabel
                        'Do nothing, just a label
                        If Trim$(Map(mapnum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text1) = Trim$(Label) Then
                            Exit Function
                        End If
                End Select
                CurSlot = CurSlot + 1
            End If
        End If
        restartlist = False
    Loop
    
    ListLeftOff = tmpListLeftOff
    CurList = tmpCurList
    CurSlot = tmpCurSlot
End Function
