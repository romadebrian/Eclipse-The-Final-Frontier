Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CHighlightItem) = GetAddress(AddressOf HandleHighlightItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CSetName) = GetAddress(AddressOf HandleSetName)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CProjecTileAttack) = GetAddress(AddressOf HandleProjecTileAttack)
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CPlayerVisibility) = GetAddress(AddressOf HandlePlayerVisibility)
    HandleDataSub(CHealPlayer) = GetAddress(AddressOf HandleHealPlayer)
    HandleDataSub(CKillPlayer) = GetAddress(AddressOf HandleKillPlayer)
    HandleDataSub(CSayGuild) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CGuildCommand) = GetAddress(AddressOf HandleGuildCommands)
    HandleDataSub(CSaveGuild) = GetAddress(AddressOf HandleGuildSave)
    HandleDataSub(CCharEditorCommand) = GetAddress(AddressOf HandleCharEditorCommand)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    HandleDataSub(COpenMyBank) = GetAddress(AddressOf HandleOpenMyBank)
    HandleDataSub(CWalkthrough) = GetAddress(AddressOf HandleToggleWalkthrough)
    HandleDataSub(CFollowPlayer) = GetAddress(AddressOf HandleStartFollowingPlayer)
    HandleDataSub(CClickPos) = GetAddress(AddressOf HandleBeFriend)
    HandleDataSub(CDeleteFriend) = GetAddress(AddressOf HandleDeleteFriend)
    HandleDataSub(CUpdateFList) = GetAddress(AddressOf HandleUpdateFriendsList)
    HandleDataSub(CFriendAccept) = GetAddress(AddressOf HandleAcceptFriend)
    HandleDataSub(CFriendDecline) = GetAddress(AddressOf HandleDeclineFriend)
    HandleDataSub(CPrivateMsg) = GetAddress(AddressOf HandlePrivateMsg)
    HandleDataSub(CRequestFriendData) = GetAddress(AddressOf HandleRequestFriendData)
    HandleDataSub(CRequestEditCombo) = GetAddress(AddressOf HandleRequestEditCombos)
    HandleDataSub(CRequestCombos) = GetAddress(AddressOf HandleRequestCombos)
    HandleDataSub(CSaveCombo) = GetAddress(AddressOf HandleSaveCombo)
    HandleDataSub(CInvHidden) = GetAddress(AddressOf HandleInvHidden)
    
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Sub HandleInvHidden(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
    For i = 1 To MAX_INV
        If Player(Index).Inv(i).Selected = 1 Then
            Player(Index).Inv(i).Selected = 0
            SendHighlight Index, i
            Exit Sub
        End If
    Next i
End Sub

Sub HandleUpdateFriendsList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call UpdateFriendsList(Index)
End Sub

Sub HandleDeleteFriend(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim fName As String, i As Long
Dim Parse() As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    fName = Trim$(Buffer.ReadString)
    Parse() = Split(fName, " ")
    fName = Parse(0)
    i = FindPlayer(fName)
    
    'Is there a name in the Variable?
    If Not Len(fName) > 0 Then Exit Sub
    
    ' Name's good, remove name from the list of both players
    Call RemoveFriend(Index, fName)
    
    ' Tell the players
    Call PlayerMsg(Index, "You are no longer friends with " & fName, BrightRed)
    Call PlayerMsg(i, "You are no longer friends with " & GetPlayerName(Index), BrightRed)
    
    ' Send the data
    SendDataTo Index, PlayerFriends(Index)
    SendDataTo i, PlayerFriends(i)
    
    Set Buffer = Nothing
End Sub

Sub HandleStartFollowingPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim WhoToFollow As Long
Dim i As Long
    Dim Buffer As New clsBuffer
    Buffer.WriteBytes Data()
    WhoToFollow = Buffer.ReadLong
    
    ' Make sure we're not following anyone else
    For i = 1 To MAX_PLAYERS
        If Player(i).Follower = Index Then
            Player(i).Follower = 0
            Call SendPlayerData(i)
            Exit For
        End If
    Next i
    
    If FollowerIsNearMe(Index, WhoToFollow, False) Then
        Player(WhoToFollow).Follower = Index
        Call PlayerMsg(Index, "You are now following: " & GetPlayerName(WhoToFollow), BrightBlue)
    Else
        Call PlayerMsg(Index, "You must be next to a player to follow them.", Red)
    End If
    Set Buffer = Nothing
End Sub

Public Sub HandleToggleWalkthrough(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim WalkThrough As Boolean
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then Exit Sub
    WalkThrough = Player(Index).WalkThrough
    Player(Index).WalkThrough = Not WalkThrough
    If WalkThrough Then Call PlayerMsg(Index, "Walkthrough Deactivated.", White)
    If Not WalkThrough Then Call PlayerMsg(Index, "Walkthrough Activated.", White)
    Call SendPlayerData(Index)
End Sub

Private Sub HandleOpenMyBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendBank Index
    TempPlayer(Index).InBank = True
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not isPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(Index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index
                Else
                    ' send new char shit
                    If Not isPlaying(Index) Then
                        Call SendNewCharClasses(Index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not isPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(Index, Name)

            If LenB(Trim$(Player(Index).Name)) > 0 Then
                Call DeleteName(Player(Index).Name)
            End If

            Call ClearPlayer(Index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not isPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(Index, Name)
            ClearBank Index
            LoadBank Index, Name
            ' check skill stats
            Call CheckSkills(Index)
            
            ' Check if character data has been created
            If LenB(Trim$(Player(Index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar Index
            Else
                ' send new char shit
                If Not isPlaying(Index) Then
                    Call SendNewCharClasses(Index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    If Not isPlaying(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar Index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Msg, White)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePrivateMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long, OrigName As String
    Dim Continue As Boolean
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    OrigName = Buffer.ReadString
    MsgTo = FindPlayer(OrigName)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        ' Make sure the two are friends.
        Continue = False
        For i = 1 To GetPlayerFriends(Index)
            If LCase$(GetPlayerFriendName(Index, i)) = LCase$(OrigName) Then
                Continue = True
            End If
        Next
            
        If Not Continue Then
            Call PlayerMsg(Index, "Only friends can Private Message each other.", BrightRed)
            Call PlayerMsg(Index, "To send a friend request, target the player by clicking on them and press the letter 'B' on your keyboard", White)
            Exit Sub
        End If
            
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, "[PM] " & GetPlayerName(Index) & ": '" & Msg & "'", TellColor)
            Call PlayerMsg(Index, "[PM] " & GetPlayerName(MsgTo) & ": '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them from moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prevent player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    ' If following someone, stop
    'For I = 1 To MAX_PLAYERS
    '    If Player(Index).Follower = I Then
    '        Player(Index).Follower = 0
    '        Exit For
    '    End If
    'Next I
    
    Call PlayerMove(Index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem Index, invNum
    
    ' send highlight item
    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighlightItem
    Buffer.WriteLong invNum
    Buffer.WriteLong Player(Index).Inv(invNum).Selected
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleHighlightItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long, i As Long, tempNum As Long, aiiSelected As Boolean
Dim Sel1 As Boolean, Sel2 As Boolean, II As Long
Dim Sel1_Index As Long, Sel2_Index As Long
Dim reSet As Boolean
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong
    aiiSelected = False
    
    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then Exit Sub
    
        Call CheckHighlight(Index, invNum)
    
    Set Buffer = Nothing
    SendHighlight Index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    
    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack Index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    CheckResource Index, x, y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(Index, PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerRawStat(Index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(Index, Stats.Endurance, GetPlayerRawStat(Index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(Index, Stats.Intelligence, GetPlayerRawStat(Index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(Index, Stats.Agility, GetPlayerRawStat(Index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(Index, Stats.Willpower, GetPlayerRawStat(Index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData Index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString 'Parse(1)
    Set Buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
    Exit Sub
    End If
    
    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    i = FindPlayer(Buffer.ReadString)
    Set Buffer = Nothing
    
    Call SetPlayerSprite(i, n)
    Call SendPlayerData(i)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim mapnum As Long
    Dim x As Long
    Dim y As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Index)
    i = Map(mapnum).Revision + 1
    Call ClearMap(mapnum)
    
    Map(mapnum).Name = Buffer.ReadString
    Map(mapnum).Music = Buffer.ReadString
    Map(mapnum).BGS = Buffer.ReadString
    Map(mapnum).Revision = i
    Map(mapnum).Moral = Buffer.ReadByte
    Map(mapnum).Up = Buffer.ReadLong
    Map(mapnum).Down = Buffer.ReadLong
    Map(mapnum).Left = Buffer.ReadLong
    Map(mapnum).Right = Buffer.ReadLong
    Map(mapnum).BootMap = Buffer.ReadLong
    Map(mapnum).BootX = Buffer.ReadByte
    Map(mapnum).BootY = Buffer.ReadByte
    
    Map(mapnum).Weather = Buffer.ReadLong
    Map(mapnum).WeatherIntensity = Buffer.ReadLong
    
    Map(mapnum).Fog = Buffer.ReadLong
    Map(mapnum).FogSpeed = Buffer.ReadLong
    Map(mapnum).FogOpacity = Buffer.ReadLong
    
    Map(mapnum).Red = Buffer.ReadLong
    Map(mapnum).Green = Buffer.ReadLong
    Map(mapnum).Blue = Buffer.ReadLong
    Map(mapnum).Alpha = Buffer.ReadLong
    
    Map(mapnum).MaxX = Buffer.ReadByte
    Map(mapnum).MaxY = Buffer.ReadByte
    
    Map(mapnum).DropItemsOnDeath = Buffer.ReadByte
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Layer(i).x = Buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(i).y = Buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(i).Tileset = Buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Autotile(z) = Buffer.ReadLong
            Next
            Map(mapnum).Tile(x, y).Type = Buffer.ReadByte
            Map(mapnum).Tile(x, y).Data1 = Buffer.ReadLong
            Map(mapnum).Tile(x, y).Data2 = Buffer.ReadLong
            Map(mapnum).Tile(x, y).Data3 = Buffer.ReadLong
            Map(mapnum).Tile(x, y).Data4 = Buffer.ReadString
            Map(mapnum).Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(mapnum).NPC(x) = Buffer.ReadLong
        Map(mapnum).NpcSpawnType(x) = Buffer.ReadLong
        Call ClearMapNpc(x, mapnum)
    Next
    
    'Event Data!
    Map(mapnum).EventCount = Buffer.ReadLong
        
    If Map(mapnum).EventCount > 0 Then
        ReDim Map(mapnum).Events(0 To Map(mapnum).EventCount)
        For i = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(i)
                .Name = Buffer.ReadString
                .Global = Buffer.ReadLong
                .x = Buffer.ReadLong
                .y = Buffer.ReadLong
                .PageCount = Buffer.ReadLong
            End With
            If Map(mapnum).Events(i).PageCount > 0 Then
                ReDim Map(mapnum).Events(i).Pages(0 To Map(mapnum).Events(i).PageCount)
                For x = 1 To Map(mapnum).Events(i).PageCount
                    With Map(mapnum).Events(i).Pages(x)
                        .chkVariable = Buffer.ReadLong
                        .VariableIndex = Buffer.ReadLong
                        .VariableCondition = Buffer.ReadLong
                        .VariableCompare = Buffer.ReadLong
                            
                        .chkSwitch = Buffer.ReadLong
                        .SwitchIndex = Buffer.ReadLong
                        .SwitchCompare = Buffer.ReadLong
                            
                        .chkHasItem = Buffer.ReadLong
                        .HasItemIndex = Buffer.ReadLong
                        .HasItemAmount = Buffer.ReadLong
                            
                        .chkSelfSwitch = Buffer.ReadLong
                        .SelfSwitchIndex = Buffer.ReadLong
                        .SelfSwitchCompare = Buffer.ReadLong
                            
                        .GraphicType = Buffer.ReadLong
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .GraphicX2 = Buffer.ReadLong
                        .GraphicY2 = Buffer.ReadLong
                            
                        .MoveType = Buffer.ReadLong
                        .MoveSpeed = Buffer.ReadLong
                        .MoveFreq = Buffer.ReadLong
                            
                        .MoveRouteCount = Buffer.ReadLong
                        
                        .IgnoreMoveRoute = Buffer.ReadLong
                        .RepeatMoveRoute = Buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(mapnum).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                            For y = 1 To .MoveRouteCount
                                .MoveRoute(y).Index = Buffer.ReadLong
                                .MoveRoute(y).Data1 = Buffer.ReadLong
                                .MoveRoute(y).Data2 = Buffer.ReadLong
                                .MoveRoute(y).Data3 = Buffer.ReadLong
                                .MoveRoute(y).Data4 = Buffer.ReadLong
                                .MoveRoute(y).Data5 = Buffer.ReadLong
                                .MoveRoute(y).Data6 = Buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = Buffer.ReadLong
                        .DirFix = Buffer.ReadLong
                        .WalkThrough = Buffer.ReadLong
                        .ShowName = Buffer.ReadLong
                        .Trigger = Buffer.ReadLong
                        .CommandListCount = Buffer.ReadLong
                            
                        .Position = Buffer.ReadLong
                    End With
                        
                    If Map(mapnum).Events(i).Pages(x).CommandListCount > 0 Then
                        ReDim Map(mapnum).Events(i).Pages(x).CommandList(0 To Map(mapnum).Events(i).Pages(x).CommandListCount)
                        For y = 1 To Map(mapnum).Events(i).Pages(x).CommandListCount
                            Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount = Buffer.ReadLong
                            Map(mapnum).Events(i).Pages(x).CommandList(y).ParentList = Buffer.ReadLong
                            If Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                ReDim Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount)
                                For z = 1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        .Index = Buffer.ReadLong
                                        .Text1 = Buffer.ReadString
                                        .Text2 = Buffer.ReadString
                                        .Text3 = Buffer.ReadString
                                        .Text4 = Buffer.ReadString
                                        .Text5 = Buffer.ReadString
                                        .Data1 = Buffer.ReadLong
                                        .Data2 = Buffer.ReadLong
                                        .Data3 = Buffer.ReadLong
                                        .Data4 = Buffer.ReadLong
                                        .Data5 = Buffer.ReadLong
                                        .Data6 = Buffer.ReadLong
                                        .ConditionalBranch.CommandList = Buffer.ReadLong
                                        .ConditionalBranch.Condition = Buffer.ReadLong
                                        .ConditionalBranch.Data1 = Buffer.ReadLong
                                        .ConditionalBranch.Data2 = Buffer.ReadLong
                                        .ConditionalBranch.Data3 = Buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = Buffer.ReadLong
                                        .MoveRouteCount = Buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).Index = Buffer.ReadLong
                                                .MoveRoute(w).Data1 = Buffer.ReadLong
                                                .MoveRoute(w).Data2 = Buffer.ReadLong
                                                .MoveRoute(w).Data3 = Buffer.ReadLong
                                                .MoveRoute(w).Data4 = Buffer.ReadLong
                                                .MoveRoute(w).Data5 = Buffer.ReadLong
                                                .MoveRoute(w).Data6 = Buffer.ReadLong
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

    Call SendMapNpcsToMap(mapnum)
    Call SpawnMapNpcs(mapnum)
    Call SpawnGlobalEvents(mapnum)
    
    For i = 1 To Player_HighIndex
        If Player(i).Map = mapnum Then
            SpawnMapEventsFor i, mapnum
        End If
    Next

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(mapnum)
    Call MapCache_Create(mapnum)
    Call ClearTempTile(mapnum)
    Call CacheResources(mapnum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If isPlaying(i) And GetPlayerMap(i) = mapnum Then
            Call PlayerWarp(i, mapnum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    Call CacheMapBlocks(mapnum)

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SpawnMapEventsFor(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(Index, invNum) < 1 Or GetPlayerInvItemNum(Index, invNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(Index, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(Index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, invNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(Index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    SendMapEventData (Index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleSaveCombo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ComboSize As Long
    Dim ComboData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ComboSize = LenB(Combo(n))
    ReDim ComboData(ComboSize - 1)
    ComboData = Buffer.ReadBytes(ComboSize)
    CopyMemory ByVal VarPtr(Combo(n)), ByVal VarPtr(ComboData(0)), ComboSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateComboToAll(n)
    Call SaveCombo(n)
    Call AddLog(GetPlayerName(Index) & " saved combo #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(Index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    npcNum = Buffer.ReadLong

    ' Prevent hacking
    If npcNum < 0 Or npcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(npcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(npcNum)
    Call SaveNpc(npcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & npcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = Buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    spellnum = Buffer.ReadLong

    ' Prevent hacking
    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(spellnum)
    Call SaveSpell(spellnum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & spellnum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Set name packet ::
' :::::::::::::::::::::::
Sub HandleSetName(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The name
    i = Buffer.ReadString 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check if player is on
    If n > 0 Then

        'check to see if same level access is trying to change another access of the very same level and boot them if they are.
        If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
            Call PlayerMsg(Index, "Invalid access level.", Red)
            Exit Sub
        End If
            
        Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s name too " & i & ".", ADMIN_LOG)
        Call SetPlayerName(n, i)
        Call SendPlayerData(n)
            
        If GetPlayerAccess(n) <= 0 Then
            Call PlayerMsg(n, "Your Name has been changed.", White)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong 'CLng(Parse(1))
    y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(Index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If isPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If Not GetPlayerVisible(i) = 1 Then
                    If GetPlayerX(i) = x Then
                        If GetPlayerY(i) = y Then
                        ' Change target
                            If TempPlayer(Index).targetType = TARGET_TYPE_PLAYER And TempPlayer(Index).target = i Then
                                TempPlayer(Index).target = 0
                                TempPlayer(Index).targetType = TARGET_TYPE_NONE
                                ' send target to player
                                SendTarget Index
                            Else
                                TempPlayer(Index).target = i
                                TempPlayer(Index).targetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                SendTarget Index
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).NPC(i).num > 0 Then
            If MapNpc(GetPlayerMap(Index)).NPC(i).x = x Then
                If MapNpc(GetPlayerMap(Index)).NPC(i).y = y Then
                    If TempPlayer(Index).target = i And TempPlayer(Index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).target = 0
                        TempPlayer(Index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget Index
                    Else
                        ' Change target
                        TempPlayer(Index).target = i
                        TempPlayer(Index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget Index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Check for Spawn Tile
    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_PLAYERSPAWN Then
        If GetPlayerX(Index) = x Or GetPlayerX(Index) + 1 = x Or GetPlayerX(Index) - 1 = x Then ' Player is to west or east or on same X of spawn tile
            If GetPlayerY(Index) = y Or GetPlayerY(Index) + 1 = y Or GetPlayerY(Index) - 1 = y Then ' Player is to south of north or on same Y of spawn tile
                SetPlayerSpawn Index, GetPlayerMap(Index), x, y
                PlayerMsg Index, "You spawn point has been reset!", Yellow
            End If
        End If
    End If
End Sub

' :::::::::::::::::::
' : Location Packet :
' :::::::::::::::::::
Sub HandleBeFriend(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim pI As Long
    Dim i As Long, II As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    pI = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    
        'make sure the friend's system is activated
        If Not frmServer.chkFriendSystem.Value = vbChecked Then Exit Sub

        If isPlaying(pI) Then
            
            ' If already friends, exit out
            For II = 1 To Player(Index).Friends.Count
                If GetPlayerFriendName(Index, II) = GetPlayerName(pI) Then Exit Sub
                'If GetPlayerFriendName(pI, II) = GetPlayerName(index) Then Exit Sub
            Next II
                            
            ' If player has max amount of friends, exit out
            If GetPlayerFriends(Index) + 1 > MAX_FRIENDS Then
                Call PlayerMsg(Index, "Your Friends List is full.", BrightRed)
                Exit Sub
            End If
                            
            ' If clicked player has max amount of friends, exit out
            If GetPlayerFriends(pI) + 1 > MAX_FRIENDS Then
                Call PlayerMsg(Index, GetPlayerName(pI) & "'s Friends List is full.", BrightRed)
                Exit Sub
            End If
            
            ' Make sure player hasn't reached friend request limit.
            If GetPlayerFriendRequests(Index) + 1 > MAX_REQUESTS Then
                Call PlayerMsg(Index, "You have sent too many requests with no response. Please wait 5 minutes.", BrightRed)
                Exit Sub
            End If
                            
            ' We're good, ask other player for friendship permission.
            Call SetPlayerFriendRequests(Index, 1)
            Call AskForFriendshipFrom(pI, GetPlayerName(Index))
            Call PlayerMsg(Index, "Friend Request Sent. Awaiting reply.", Orange) ' or maybe yellow
                            
        End If
End Sub

Sub HandleAcceptFriend(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tempStr As String
Dim pI As Long, i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    tempStr = Buffer.ReadString
    
    If Not Len(tempStr) > 0 Then
        Call PlayerMsg(Index, "Could not reply to friend request. Please try again later.", BrightRed)
        Exit Sub
    End If
    
    pI = FindPlayer(tempStr)
    
    ' Go ahead and just tell the player we're good to go.
    Call PlayerMsg(pI, GetPlayerName(Index) & " has accepted your friend request.", BrightGreen)

    ' We have permission, let's make these two buds.
    Call SetPlayerFriends(pI, 1)
            
    ' Update and tell the other player
    Call SetPlayerFriendName(pI, GetPlayerFriends(pI), GetPlayerName(Index))
    Call PlayerMsg(pI, "You are now friends with " & GetPlayerName(Index), Cyan)
    
    ' Subtract a request point
    If GetPlayerFriendRequests(pI) > 0 Then Call SetPlayerFriendRequests(pI, -1)
    
    'make sure we're not doubling up friends
    For i = 1 To GetPlayerFriends(Index)
        If GetPlayerFriendName(Index, i) = GetPlayerName(pI) Then GoTo SkipThatShit
    Next i
    
    ' Update and tell yourself
    Call SetPlayerFriendName(Index, GetPlayerFriends(Index), GetPlayerName(pI))
    Call PlayerMsg(Index, "You are now friends with " & GetPlayerName(pI), Cyan)
    Call SetPlayerFriends(Index, 1)
                            
SkipThatShit:
    ' Send new data to both players
    Call SendDataTo(Index, PlayerFriends(Index))
    Call SendDataTo(pI, PlayerFriends(pI))
    
    Set Buffer = Nothing
End Sub

Sub HandleRequestFriendData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tempStr As String, pI As Long, i As Long
Dim pData(1 To 6) As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    tempStr = Buffer.ReadString
    Set Buffer = Nothing
    
    'Make sure we have a name
    If Not Len(tempStr) > 0 Then Exit Sub
    pI = FindPlayer(tempStr)
    
    'Make sure we have an index
    If pI < 1 Or pI > MAX_PLAYERS Then Exit Sub
    
    
    'Start setting data
    pData(1) = GetPlayerLevel(pI)
    pData(2) = GetPlayerStat(pI, Strength)
    pData(3) = GetPlayerStat(pI, Endurance)
    pData(4) = GetPlayerStat(pI, Intelligence)
    pData(5) = GetPlayerStat(pI, Agility)
    pData(6) = GetPlayerStat(pI, Willpower)
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SFriendData
    For i = 1 To UBound(pData)
        Buffer.WriteLong pData(i)
    Next i
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleDeclineFriend(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tempStr As String
Dim pI As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    tempStr = Buffer.ReadString
    
    If Not Len(tempStr) > 0 Then
        Call PlayerMsg(Index, "Could not reply to friend request. Please try again later.", BrightRed)
        Exit Sub
    End If
    
    pI = FindPlayer(tempStr)
    ' Simply tell the player the request was declined.
    Call PlayerMsg(pI, GetPlayerName(Index) & " has declined your friend request.", BrightRed)
    
    ' Subtract a request point (On second thought, no. lol)
    'If GetPlayerFriendRequests(index) > 0 Then Call SetPlayerFriendRequests(index, -1)
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots Index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem Index, Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems Index
End Sub

Sub HandleRequestCombos(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendCombos Index
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources Index
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops Index
End Sub

Sub HandleRequestEditCombos(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SComboEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Set Buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim thePlr As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Exit Sub

    thePlr = FindPlayer(Buffer.ReadString)

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Exit Sub

    SetPlayerExp thePlr, GetPlayerNextLevel(thePlr)
    CheckPlayerLevelUp thePlr
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellslot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellslot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellslot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(Index).Spell(spellslot) = 0
    SendPlayerSpells Index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(Index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemAmount = HasItem(Index, .costitem)
        If itemAmount = 0 Or itemAmount < .costvalue Then
            PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem Index, .costitem, .costvalue
        GiveInvItem Index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim itemnum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invSlot) < 1 Or GetPlayerInvItemNum(Index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemnum = GetPlayerInvItemNum(Index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    price = Item(itemnum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, itemnum, 1
    GiveInvItem Index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots Index, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    TakeBankItem Index, BankSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    GiveBankItem Index, invSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank Index
    SavePlayer Index
    
    TempPlayer(Index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX Index, x
        SetPlayerY Index, y
        SendPlayerXYToMap Index
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(Index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(Index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(Index).x
    sY = Player(Index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    If TempPlayer(Index).InTrade > 0 Then
        TempPlayer(Index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(Index).TradeRequest
        ' let them know they're trading
        PlayerMsg Index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
        PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
        ' clear the tradeRequest server-side
        TempPlayer(Index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they're trading with each other
        TempPlayer(Index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = Index
        ' clear out their trade offers
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade Index, tradeTarget
        SendTrade tradeTarget, Index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate Index, 0
        SendTradeUpdate Index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
    PlayerMsg Index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemnum As Long
    
    TempPlayer(Index).AcceptTrade = True
    
    tradeTarget = TempPlayer(Index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus Index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(Index).TradeOffer(i).num > 0 Then
                itemnum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).num).num
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem(i).num = itemnum
                    tmpTradeItem(i).Value = TempPlayer(Index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).num > 0 Then
                itemnum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).num = itemnum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num, tmpTradeItem2(i).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).num > 0 Then
                ' give away!
                GiveInvItem Index, tmpTradeItem2(i).num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).num, tmpTradeItem(i).Value, False
            End If
        Next
    
        SendInventory Index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, "Trade completed.", BrightGreen
        PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
            
    End If
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade
    
    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, "You declined the trade.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim itemnum As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(Index, invSlot)
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invSlot) Then
        Exit Sub
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).num = invSlot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).Value = TempPlayer(Index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).Value > GetPlayerInvItemValue(Index, invSlot) Then
                    TempPlayer(Index).TradeOffer(i).Value = GetPlayerInvItemValue(Index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).num = invSlot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).num = invSlot
    TempPlayer(Index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(Index).Hotbar(hotbarNum).Slot = 0
            Player(Index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(Index).Inv(Slot).num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Inv(Slot).num
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(Index).Spell(Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Spell(Slot)
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case Player(Index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).num > 0 Then
                    If Player(Index).Inv(i).num = Player(Index).Hotbar(Slot).Slot Then
                        If Item(Player(Index).Inv(i).num).Type = ITEM_TYPE_CONSUME Then
                            If Not Item(Player(Index).Inv(i).num).Stackable = 1 Then
                                Player(Index).Hotbar(Slot).Slot = 0
                                Player(Index).Hotbar(Slot).sType = 0
                            End If
                            SendHotbar Index
                        End If
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i) > 0 Then
                    If Player(Index).Spell(i) = Player(Index).Hotbar(Slot).Slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(Index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(Index).target = Index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(Index).target) Or Not isPlaying(TempPlayer(Index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(Index).target) <> GetPlayerMap(Index) Then Exit Sub
    
    ' init the request
    Party_Invite Index, TempPlayer(Index).target
End Sub

Sub HandleAcceptParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(Index).partyInvite, Index
End Sub

Sub HandleDeclineParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(Index).partyInvite, Index
End Sub

Sub HandlePartyLeave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave Index
End Sub

Sub HandleEventChatReply(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim eventID As Long, pageID As Long, reply As Long, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    eventID = Buffer.ReadLong
    pageID = Buffer.ReadLong
    reply = Buffer.ReadLong
    
    If TempPlayer(Index).EventProcessingCount > 0 Then
        For i = 1 To TempPlayer(Index).EventProcessingCount
            If TempPlayer(Index).EventProcessing(i).eventID = eventID And TempPlayer(Index).EventProcessing(i).pageID = pageID Then
                If TempPlayer(Index).EventProcessing(i).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Index = EventType.evShowText Then
                            TempPlayer(Index).EventProcessing(i).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data1
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 2
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data2
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 3
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data3
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 4
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data4
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(Index).EventProcessing(i).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    
    
    
    Set Buffer = Nothing
End Sub

Sub HandleEvent(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long, begineventprocessing As Boolean, z As Long, Buffer As clsBuffer

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    i = Buffer.ReadLong
    Set Buffer = Nothing
    
    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
            If TempPlayer(Index).EventMap.EventPages(z).eventID = i Then
                i = z
                begineventprocessing = True
                Exit For
            End If
        Next
    End If
    
    If begineventprocessing = True Then
        If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
            'Process this event, it is action button and everything checks out.
            TempPlayer(Index).EventProcessingCount = TempPlayer(Index).EventProcessingCount + 1
            ReDim Preserve TempPlayer(Index).EventProcessing(TempPlayer(Index).EventProcessingCount)
            With TempPlayer(Index).EventProcessing(TempPlayer(Index).EventProcessingCount)
                .ActionTimer = GetTickCount
                .CurList = 1
                .CurSlot = 1
                .eventID = TempPlayer(Index).EventMap.EventPages(i).eventID
                .pageID = TempPlayer(Index).EventMap.EventPages(i).pageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount)
            End With
            'Call CheckTasks(index, QUEST_TYPE_GOGETFROMEVENT, TempPlayer(index).EventMap.EventPages(i).eventID)
        End If
        begineventprocessing = False
    End If
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (Index)
End Sub

Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set Buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

Sub HandlePlayerVisibility(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    If Not Player(Index).Visible = 0 Then
        Player(Index).Visible = 0
    Else
        Player(Index).Visible = 1
    End If
    
    Call SendPlayerData(Index)
End Sub

' ::::::::::::::::::::::::
' :: Heal Player packet ::
' ::::::::::::::::::::::::
Sub HandleHealPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then Exit Sub

    ' The index
    n = FindPlayer(Buffer.ReadString)
    Set Buffer = Nothing

    ' Check if player is on
    If n > 0 Then
        Call SetPlayerVital(n, Vitals.HP, GetPlayerMaxVital(n, Vitals.HP))
        Call SetPlayerVital(n, Vitals.MP, GetPlayerMaxVital(n, Vitals.MP))
        Call SendVital(n, Vitals.HP)
        Call SendVital(n, Vitals.MP)
        Call PlayerMsg(n, "You have been healed by " & GetPlayerName(Index) & "!", BrightBlue)
        Call AddLog(GetPlayerName(Index) & " healed" & GetPlayerName(n) & ".", ADMIN_LOG)
    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Kill Player packet ::
' ::::::::::::::::::::::::
Sub HandleKillPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then Exit Sub

    ' The index
    n = FindPlayer(Buffer.ReadString)
    Set Buffer = Nothing

    ' Check if player is on
    If n > 0 Then
        Call SetPlayerVital(n, Vitals.HP, 0)
        Call SendVital(n, Vitals.HP)
        Call OnDeath(n)
        Call PlayerMsg(n, "You have been killed by " & GetPlayerName(Index) & "!", BrightRed)
        Call AddLog(GetPlayerName(Index) & " killed" & GetPlayerName(n) & ".", ADMIN_LOG)
    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Client Character Editor ::
' :::::::::::::::::::::::::::::
Sub SendCharEditorRequest(ByVal i As Long, ByVal command As Byte, ByVal num As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SCharEditorRequest
    
    Select Case command
        Case 1:
            Buffer.WriteByte command
            Buffer.WriteLong GetPlayerLevel(i)
            Buffer.WriteLong GetPlayerExp(i)
            Buffer.WriteLong GetPlayerPOINTS(i)
            Buffer.WriteLong GetPlayerStat(i, Endurance)
            Buffer.WriteLong GetPlayerStat(i, Strength)
            Buffer.WriteLong GetPlayerStat(i, Intelligence)
            Buffer.WriteLong GetPlayerStat(i, Agility)
            Buffer.WriteLong GetPlayerStat(i, Willpower)
            Buffer.WriteByte GetPlayerCombatLevel(i, num)
            Buffer.WriteLong GetPlayerCombatExp(i, num)
            Buffer.WriteLong GetPlayerInvItemNum(i, num)
            Buffer.WriteLong GetPlayerInvItemValue(i, num)
            Buffer.WriteLong GetPlayerBankItemNum(i, num)
            Buffer.WriteLong GetPlayerBankItemValue(i, num)
            Buffer.WriteLong GetPlayerLevel(i)
        Case 2:
            Buffer.WriteByte command
            Buffer.WriteByte GetPlayerCombatLevel(i, num)
            Buffer.WriteLong GetPlayerCombatExp(i, num)
        Case 3:
            Buffer.WriteByte command
            Buffer.WriteLong GetPlayerInvItemNum(i, num)
            Buffer.WriteLong GetPlayerInvItemValue(i, num)
        Case 4:
            Buffer.WriteByte command
            Buffer.WriteLong GetPlayerBankItemNum(i, num)
            Buffer.WriteLong GetPlayerBankItemValue(i, num)
    End Select
    
    SendDataTo i, Buffer.ToArray
    Set Buffer = Nothing
End Sub

Sub HandleCharEditorCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, n As Long, command As Byte, plExp As Long, plPts As Long, pStr As Long, pEnd As Long, pInt As Long, pAgi As Long, pWill As Long
    Dim lvl As Long, invNum As Long, itmNum As Long, itmQty As Long, bnkNum As Long, bankNum As Byte, bankQty As Long
    Dim comType As Byte, comLvl As Byte, comExp As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Exit Sub

    command = Buffer.ReadByte
    
    Select Case command
        Case 1
            i = FindPlayer(Buffer.ReadString)
            If Not i = 0 Then
                SendCharEditorRequest i, 1, 1
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
        Case 2
            i = FindPlayer(Buffer.ReadString)
            lvl = Buffer.ReadLong
            plExp = Buffer.ReadLong
            plPts = Buffer.ReadLong
            If GetPlayerLevel(i) < lvl Then
                SetPlayerPOINTS i, GetPlayerPOINTS(i) + (3 * (lvl - GetPlayerLevel(i)))
                SetPlayerLevel i, lvl
            Else
                SetPlayerLevel i, lvl
                SetPlayerExp i, plExp
                SetPlayerPOINTS i, plPts
            End If
            pEnd = Buffer.ReadLong
            pStr = Buffer.ReadLong
            pInt = Buffer.ReadLong
            pAgi = Buffer.ReadLong
            pWill = Buffer.ReadLong
            If pEnd > 100 Then pEnd = 100
            If pStr > 100 Then pStr = 100
            If pInt > 100 Then pInt = 100
            If pAgi > 100 Then pAgi = 100
            If pWill > 100 Then pWill = 100
            SetPlayerStat i, Endurance, pEnd
            SetPlayerStat i, Strength, pStr
            SetPlayerStat i, Intelligence, pInt
            SetPlayerStat i, Agility, pAgi
            SetPlayerStat i, Willpower, pWill
                invNum = Buffer.ReadLong
                itmNum = Buffer.ReadLong
                itmQty = Buffer.ReadLong
            SetPlayerInvItemNum i, invNum, itmNum
            SetPlayerInvItemValue i, invNum, itmQty
                bnkNum = Buffer.ReadLong
                bankNum = Buffer.ReadLong
                bankQty = Buffer.ReadLong
            SetPlayerBankItemNum i, bnkNum, bankNum
            SetPlayerBankItemValue i, bnkNum, bankQty
            SendInventoryUpdate i, invNum
            SendEXP i
            CheckPlayerLevelUp i
            SaveBank i
            SavePlayer i
            SendPlayerData i
                SendCharEditorRequest i, 1, 1
        Case 3
            i = FindPlayer(Buffer.ReadString)
            comType = Buffer.ReadByte
            If comType > MAX_COMBAT Then
                Call PlayerMsg(Index, "Number to high, combat skills only go to " & MAX_COMBAT, AlertColor)
                Exit Sub
            End If
        
            If Not i = 0 Then
                If Not comType = 0 Then
                    SendCharEditorRequest i, 2, comType
                Else
                    Call PlayerMsg(Index, "combat skill must be greater then 0!", AlertColor)
                End If
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
        Case 4
            i = FindPlayer(Buffer.ReadString)
            comType = Buffer.ReadByte
            comLvl = Buffer.ReadByte
            comExp = Buffer.ReadLong
            If comType > MAX_COMBAT Then
                Call PlayerMsg(Index, "Number to high, combat skills only go to " & MAX_COMBAT, AlertColor)
                Exit Sub
            End If
        
            If Not i = 0 Then
                If Not comType = 0 Then
                    SetPlayerCombatLevel i, comLvl, comType
                    SetPlayerCombatExp i, comType, comExp
                    SendPlayerData i
                    SavePlayer i
                    SendCharEditorRequest i, 2, comType
                Else
                    Call PlayerMsg(Index, "combat skill must be greater then 0!", AlertColor)
                End If
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
        Case 5
            i = FindPlayer(Buffer.ReadString)
            n = Buffer.ReadLong
            If n > MAX_INV Then
                Call PlayerMsg(Index, "Number to high, inventory only goes to " & MAX_INV, AlertColor)
                Exit Sub
            End If
        
            If Not i = 0 Then
                If Not n = 0 Then
                    SendCharEditorRequest i, 3, n
                Else
                    Call PlayerMsg(Index, "Item Num must be greater then 0!", AlertColor)
                End If
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
        Case 6
            i = FindPlayer(Buffer.ReadString)
            invNum = Buffer.ReadLong
            itmNum = Buffer.ReadLong
            itmQty = Buffer.ReadLong
            If invNum > MAX_INV Then
                Call PlayerMsg(Index, "Number to high, inventory only goes to " & MAX_INV, AlertColor)
                Exit Sub
            End If
        
            If Not i = 0 Then
                If Not invNum = 0 Then
                    SetPlayerInvItemNum i, invNum, itmNum
                    SetPlayerInvItemValue i, invNum, itmQty
                    SendInventoryUpdate i, invNum
                    SendPlayerData i
                    SavePlayer i
                    SendCharEditorRequest i, 3, invNum
                Else
                    Call PlayerMsg(Index, "Item Num must be greater then 0!", AlertColor)
                End If
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
        Case 7
            i = FindPlayer(Buffer.ReadString)
            n = Buffer.ReadLong
            If n > MAX_BANK Then
                Call PlayerMsg(Index, "Number to high, bank only goes to " & MAX_BANK, AlertColor)
                Exit Sub
            End If
        
            If Not i = 0 Then
                If Not n = 0 Then
                    SendCharEditorRequest i, 4, n
                Else
                    Call PlayerMsg(Index, "Item Num must be greater then 0!", AlertColor)
                End If
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
        Case 8
            i = FindPlayer(Buffer.ReadString)
            bnkNum = Buffer.ReadLong
            bankNum = Buffer.ReadLong
            bankQty = Buffer.ReadLong
            If n > MAX_BANK Then
                Call PlayerMsg(Index, "Number to high, bank only goes to " & MAX_BANK, AlertColor)
                Exit Sub
            End If
        
            If Not i = 0 Then
                If Not bnkNum = 0 Then
                    SetPlayerBankItemNum i, bnkNum, bankNum
                    SetPlayerBankItemValue i, bnkNum, bankQty
                    SaveBank i
                    SavePlayer i
                    SendPlayerData i
                    SendCharEditorRequest i, 4, bankNum
                Else
                    Call PlayerMsg(Index, "Item Num must be greater then 0!", AlertColor)
                End If
            Else
                Call PlayerMsg(Index, "Player Not Found!", AlertColor)
            End If
    End Select
End Sub
Sub HandleRequestEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    
    n = Buffer.ReadLong 'CLng(Parse(1))
    
    If n < 0 Or n > MAX_QUESTS Then
    Exit Sub
    End If
    
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(Index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
End Sub

Sub HandlePlayerHandleQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim QuestNum As Long, Order As Long, i As Long, n As Long
Dim RemoveStartItems As Boolean

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong
    'prevent error, but tell me about it QUICKCHANGE
    If QuestNum < 1 Then
        Call PlayerMsg(Index, "Could not retrieve quest data.", Red)
        Exit Sub
    End If
    Order = Buffer.ReadLong '1 = accept quest, 2 = cancel quest

    If Order = 1 Then
        RemoveStartItems = False
        'Alatar v1.2
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                    PlayerMsg Index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                    RemoveStartItems = True
                    Exit For
                Else
                    If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            If FindOpenInvSlot(Index, Quest(QuestNum).GiveItem(i).Item) = 0 Then
                                PlayerMsg Index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                                RemoveStartItems = True
                                Exit For
                            Else
                                GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1
                            End If
                        Next
                    End If
                End If
            End If
        Next

        If RemoveStartItems = False Then 'this means everything went ok
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED '1
            Player(Index).PlayerQuest(QuestNum).ActualTask = 1
            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
            PlayerMsg Index, "New quest accepted: " & Trim$(Quest(QuestNum).Name) & "!", BrightGreen
        End If
    
    ElseIf Order = 2 Then
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED '2
        Player(Index).PlayerQuest(QuestNum).ActualTask = 1
        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
        RemoveStartItems = True 'avoid exploits
        PlayerMsg Index, Trim$(Quest(QuestNum).Name) & " has been canceled!", BrightGreen
    End If

    If RemoveStartItems = True Then
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If HasItem(Index, Quest(QuestNum).GiveItem(i).Item) > 0 Then
                    If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If


    SavePlayer Index
    SendPlayerData Index
    SendPlayerQuests Index
    
    Set Buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests Index
End Sub

Private Sub HandleProjecTileAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim curProjecTile As Long, i As Long, CurEquipment As Long

    ' prevent subscript
    If Index > MAX_PLAYERS Or Index < 1 Then Exit Sub
    
    ' get the players current equipment
    CurEquipment = GetPlayerEquipment(Index, Weapon)
    
    ' check if they've got equipment
    If CurEquipment < 1 Or CurEquipment > MAX_ITEMS Then Exit Sub
    
    ' set the curprojectile
    For i = 1 To MAX_PLAYER_PROJECTILES
        If TempPlayer(Index).ProjecTile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile Index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
        End If
    Next
    
    ' check for subscript
    If curProjecTile < 1 Then Exit Sub
    
    ' populate the data in the player rec
    With TempPlayer(Index).ProjecTile(curProjecTile)
        .Damage = Item(CurEquipment).ProjecTile.Damage
        .Direction = GetPlayerDir(Index)
        .Pic = Item(CurEquipment).ProjecTile.Pic
        .Range = Item(CurEquipment).ProjecTile.Range
        .Speed = Item(CurEquipment).ProjecTile.Speed
        .x = GetPlayerX(Index)
        .y = GetPlayerY(Index)
    End With
                
    ' trololol, they have no more projectile space left
    If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' update the projectile on the map
    SendProjectileToMap Index, curProjecTile
    
End Sub
