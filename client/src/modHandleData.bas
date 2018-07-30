Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SHandleProjectile) = GetAddress(AddressOf HandleProjectile)
    HandleDataSub(SSendGuild) = GetAddress(AddressOf HandleSendGuild)
    HandleDataSub(SAdminGuild) = GetAddress(AddressOf HandleAdminGuild)
    HandleDataSub(SCharEditorRequest) = GetAddress(AddressOf HandleCharEditorRequest)
    HandleDataSub(SPlayerCombatEXP) = GetAddress(AddressOf HandlePlayerCombatEXP)
    HandleDataSub(SFollowPlayer) = GetAddress(AddressOf HandleFollowPlayer)
    HandleDataSub(SUpdateSkill) = GetAddress(AddressOf HandleUpdateSkill)
    HandleDataSub(SUpdateFList) = GetAddress(AddressOf HandleUpdateFriendsList)
    HandleDataSub(SFriendRequest) = GetAddress(AddressOf HandleFriendRequest)
    HandleDataSub(SFriendData) = GetAddress(AddressOf HandleFriendData)
    HandleDataSub(SFriends) = GetAddress(AddressOf HandlePlayerFriends)
    HandleDataSub(SClearFList) = GetAddress(AddressOf HandleClearFList)
    HandleDataSub(SHighlightItem) = GetAddress(AddressOf HandleHighlightItem)
    HandleDataSub(SUpdateCombo) = GetAddress(AddressOf HandleUpdateCombo)
    HandleDataSub(SComboEditor) = GetAddress(AddressOf HandleComboEditor)
    HandleDataSub(SOpenBook) = GetAddress(AddressOf HandleOpenBook)
    HandleDataSub(SGUIBars) = GetAddress(AddressOf HandleGUIBars)
    
    
    'Events
    HandleDataSub(SSpawnEvent) = GetAddress(AddressOf HandleSpawnEventPage)
    HandleDataSub(SEventMove) = GetAddress(AddressOf HandleEventMove)
    HandleDataSub(SEventDir) = GetAddress(AddressOf HandleEventDir)
    HandleDataSub(SEventChat) = GetAddress(AddressOf HandleEventChat)
    
    HandleDataSub(SEventStart) = GetAddress(AddressOf HandleEventStart)
    HandleDataSub(SEventEnd) = GetAddress(AddressOf HandleEventEnd)
    
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SStopSound) = GetAddress(AddressOf HandleStopSound)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    
    HandleDataSub(SMapEventData) = GetAddress(AddressOf HandleMapEventData)
    
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    
    HandleDataSub(SSpecialEffect) = GetAddress(AddressOf HandleSpecialEffect)
    
    HandleDataSub(SFlash) = GetAddress(AddressOf HandleFlash)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleGUIBars(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    OldGuiBars = CBool(Val(buffer.ReadLong))
    
    Set buffer = Nothing
End Sub

Sub HandleOpenBook(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' open/close the motha fuckin book
    GUIWindow(GUI_BOOK).Visible = Not GUIWindow(GUI_BOOK).Visible
    If Not GUIWindow(GUI_BOOK).Visible Then OpeningBook = True
End Sub

Sub HandleHighlightItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim slotNum As Long, vValue As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    slotNum = buffer.ReadLong
    vValue = buffer.ReadLong
    If slotNum < 1 Then Exit Sub
    If vValue < 0 Or vValue > 1 Then Exit Sub
    
    PlayerInv(slotNum).Selected = vValue
    
    Set buffer = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError "HandleHighlightItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClearFList(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.lstFriends.Clear
    Exit Sub
ErrorHandler:
    HandleError "HandleClearFList", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleFriendRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, tempname As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    FriendRequestSender = buffer.ReadString
    FriendRequestVisible = True
    
    'Now we open gui with friend requests name and wait for the reply.
    GUIWindow(GUI_FRIENDREQUEST).Visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateFriendsList", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleUpdateFriendsList(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tempname As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    tempname = buffer.ReadString
    frmMain.lstFriends.AddItem tempname
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateFriendsList", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleFollowPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Dir As Byte
Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
        buffer.WriteBytes data()
        Dir = buffer.ReadByte
        Call CheckMovement(True, Dir)
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleFollowPlayer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    Show_Login False
    frmMenu.picCharacter.Visible = False
    Show_Register False
    frmMenu.lblNews.Visible = True
    
    Msg = buffer.ReadString 'Parse(1)
    
    Set buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Options.Game_Name)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' save options
    Options.savePass = frmMenu.chkPass.value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' player high index
    Player_HighIndex = buffer.ReadLong
    
    Set buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("Receiving game data...")
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim Z As Long, x As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For I = 1 To Max_Classes

        With Class(I)
            .name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For x = 0 To Z
                .MaleSprite(x) = buffer.ReadLong
            Next
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For x = 0 To Z
                .FemaleSprite(x) = buffer.ReadLong
            Next
            
            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMenu.picCharacter.Visible = True
    frmMenu.picCredits.Visible = False
    Show_Login False
    Show_Register False
    frmLoad.Visible = False
    frmMenu.cmbClass.Clear
    For I = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(I).name)
    Next

    frmMenu.cmbClass.ListIndex = 0
    n = frmMenu.cmbClass.ListIndex + 1
    
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim Z As Long, x As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For I = 1 To Max_Classes

        With Class(I)
            .name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For x = 0 To Z
                .MaleSprite(x) = buffer.ReadLong
            Next
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For x = 0 To Z
                .FemaleSprite(x) = buffer.ReadLong
            Next
                            
            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InGame = True
    Call GameInit
    Call GameLoop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1

    For I = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, I, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, I, buffer.ReadLong)
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).Visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))
    ' changes, clear drop menu
        sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).Visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).Visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)
    
    'If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
    '    'frmMain.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
    '    frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    '    ' hp bar
    '    frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    'End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)
    
    'If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
    '    'frmMain.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
    '    frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    '    ' mp bar
    '    frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPRBar_Width)) * SPRBar_Width
    'End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    'For i = 1 To Stats.Stat_Count - 1
    '    SetPlayerStat Index, i, buffer.ReadLong
    '    frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i)
    'Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong
    
    'frmMain.lblEXP.Caption = GetPlayerExp(Index) & "/" & TNL
    ' mp bar
    'frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerFriends(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Player(MyIndex).Friends.count = buffer.ReadLong
    If Player(MyIndex).Friends.count > MAX_FRIENDS Then Player(MyIndex).Friends.count = MAX_FRIENDS
    
    If Player(MyIndex).Friends.count > 0 Then
        For I = 1 To Player(MyIndex).Friends.count
            Player(MyIndex).Friends.NameOfFriend(I) = buffer.ReadString
        Next I
    Else
        For I = 1 To MAX_FRIENDS
            Player(MyIndex).Friends.NameOfFriend(I) = vbNullString
        Next I
    End If
    
    Call UpdateFriendsList
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerFriends", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long, x As Long, Z As Long, II As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    Call SetPlayerName(I, buffer.ReadString)
    Call SetPlayerLevel(I, buffer.ReadLong)
    Call SetPlayerPOINTS(I, buffer.ReadLong)
    Call SetPlayerSprite(I, buffer.ReadLong)
    Call SetPlayerMap(I, buffer.ReadLong)
    Call SetPlayerX(I, buffer.ReadLong)
    Call SetPlayerY(I, buffer.ReadLong)
    Call SetPlayerDir(I, buffer.ReadLong)
    Call SetPlayerAccess(I, buffer.ReadLong)
    Call SetPlayerPK(I, buffer.ReadLong)
    Call SetPlayerClass(I, buffer.ReadLong)
    Call SetPlayerVisible(I, buffer.ReadLong)
    'set walkthrough
    Player(MyIndex).Walkthrough = CBool(buffer.ReadLong)
    Player(I).Follower = buffer.ReadLong
    MAX_SKILLS = buffer.ReadLong
    ' Compensate for new array length
    If MAX_SKILLS > 0 Then ReDim Skill(1 To MAX_SKILLS)
    If MAX_SKILLS > 0 Then ReDim Player(I).Skills(1 To MAX_SKILLS)
    
    If MAX_SKILLS > 0 Then
        For II = 1 To MAX_SKILLS
            Skill(II).name = buffer.ReadString
            Skill(II).MaxLvl = buffer.ReadLong
            Player(I).Skills(II).Level = buffer.ReadLong
            Player(I).Skills(II).EXP = buffer.ReadLong
            Player(I).Skills(II).EXP_Needed = buffer.ReadLong
        Next II
    End If
    
    For x = 1 To Stats.Stat_Count - 1
        SetPlayerStat I, x, buffer.ReadLong
    Next
    
    For x = 1 To MAX_COMBAT
        Player(I).Combat(x).Level = buffer.ReadByte
        Player(I).Combat(x).EXP = buffer.ReadLong
        CombatTNL(x) = buffer.ReadLong
    Next
    
    If buffer.ReadByte = 1 Then
        Player(I).GuildName = buffer.ReadString
        Player(I).GuildTag = buffer.ReadString
        Player(I).GuildColor = buffer.ReadInteger
    Else
        Player(I).GuildName = vbNullString
        Player(I).GuildTag = vbNullString
        Player(I).GuildColor = 0
    End If

    ' Check if the player is the client player
    If I = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If

    ' Make sure they aren't walking
    Player(I).Moving = 0
    Player(I).xOffset = 0
    Player(I).yOffset = 0
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim x As Long
Dim y As Long
Dim Z As Long
Dim Dir As Long
Dim n As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    Dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(I, x)
    Call SetPlayerY(I, y)
    Call SetPlayerDir(I, Dir)
    Player(I).xOffset = 0
    Player(I).yOffset = 0
    Player(I).Moving = n

    Select Case GetPlayerDir(I)
        Case DIR_UP
            Player(I).yOffset = PIC_Y
        Case DIR_DOWN
            Player(I).yOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(I).xOffset = PIC_X
        Case DIR_RIGHT
            Player(I).xOffset = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim x As Long
Dim y As Long
Dim Dir As Long
Dim Movement As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    Dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .x = x
        .y = y
        MiniMapNPC(MapNpcNum).x = x * 4
        MiniMapNPC(MapNpcNum).y = y * 4
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerDir(I, Dir)

    With Player(I)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    Dir = buffer.ReadLong

    With MapNpc(I)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim Dir As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    x = buffer.ReadLong
    y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, x)
    Call SetPlayerY(MyIndex, y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim Dir As Long
Dim buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, x)
    Call SetPlayerY(thePlayer, y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    ' Set player to attacking
    Player(I).Attacking = 1
    Player(I).AttackTimer = GetTickCount
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    ' Set player to attacking
    MapNpc(I).Attacking = 1
    MapNpc(I).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim I As Long
Dim NeedMap As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Erase all players except self
    For I = 1 To MAX_PLAYERS
        If I <> MyIndex Then
            Call SetPlayerMap(I, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For I = 1 To MAX_BYTE
        Blood(I).x = 0
        Blood(I).y = 0
        Blood(I).Sprite = 0
        Blood(I).timer = 0
    Next
    
    Map.CurrentEvents = 0
    ReDim Map.MapEvents(0)
    
    ' Get map num
    x = buffer.ReadLong
    ' Get revision
    y = buffer.ReadLong

    If FileExist(MAP_PATH & "map" & x & MAP_EXT, False) Then
        Call LoadMap(x)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
            CacheNewMapSounds
            initAutotiles
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    GettingMap = True
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim x As Long
Dim y As Long
Dim I As Long, Z As Long, w As Long
Dim buffer As clsBuffer
Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()

    MapNum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.BGS = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    
    Map.Fog = buffer.ReadLong
    Map.FogSpeed = buffer.ReadLong
    Map.FogOpacity = buffer.ReadLong
    
    Map.Red = buffer.ReadLong
    Map.Green = buffer.ReadLong
    Map.Blue = buffer.ReadLong
    Map.alpha = buffer.ReadLong
    
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    
    Map.DropItemsOnDeath = buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                Map.Tile(x, y).layer(I).x = buffer.ReadLong
                Map.Tile(x, y).layer(I).y = buffer.ReadLong
                Map.Tile(x, y).layer(I).Tileset = buffer.ReadLong
            Next
            For Z = 1 To MapLayer.Layer_Count - 1
                Map.Tile(x, y).Autotile(Z) = buffer.ReadLong
            Next
            Map.Tile(x, y).Type = buffer.ReadByte
            Map.Tile(x, y).data1 = buffer.ReadLong
            Map.Tile(x, y).Data2 = buffer.ReadLong
            Map.Tile(x, y).Data3 = buffer.ReadLong
            Map.Tile(x, y).Data4 = buffer.ReadString
            Map.Tile(x, y).DirBlock = buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map.NPC(x) = buffer.ReadLong
        Map.NpcSpawnType(x) = buffer.ReadLong
        n = n + 1
    Next

    ClearTempTile
    initAutotiles
    
    Set buffer = Nothing
    
    ' Save the map
    Call SaveMap(MapNum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    CacheNewMapSounds

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For I = 1 To MAX_MAP_ITEMS
        With MapItem(I)
            .PlayerName = buffer.ReadString
            .num = buffer.ReadLong
            .value = buffer.ReadLong
            .x = buffer.ReadLong
            .y = buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For I = 1 To MAX_MAP_NPCS
        With MapNpc(I)
            .num = buffer.ReadLong
            .x = buffer.ReadLong
            .y = buffer.ReadLong
            MiniMapNPC(I).x = MapNpc(I).x * 4
            MiniMapNPC(I).y = MapNpc(I).y * 4
            .Dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
            .HPSetTo = buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
Dim I As Long
Dim MusicFile As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' clear the action msgs
    For I = 1 To MAX_BYTE
        ClearActionMsg (I)
    Next I
    Action_HighIndex = 1
    
    ' load tilesets we need
    LoadTilesets
            
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    
    ' re-position the map name
    Call UpdateDrawMapName
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For I = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(I).num > 0 Then
            Npc_HighIndex = I
            Exit For
        End If
    Next
    
    For I = 1 To MAX_BYTE
        ClearAnimInstance (I)
    Next
    
    initAutotiles
    
    CurrentWeather = Map.Weather
    CurrentWeatherIntensity = Map.WeatherIntensity
    CurrentFog = Map.Fog
    CurrentFogSpeed = Map.FogSpeed
    CurrentFogOpacity = Map.FogOpacity
    CurrentTintR = Map.Red
    CurrentTintG = Map.Green
    CurrentTintB = Map.Blue
    CurrentTintA = Map.alpha

    GettingMap = False
    CanMoveNow = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFriendData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To 6
        PlayerInfoValue(I) = buffer.ReadLong
    Next I
    
    'We have the data, now Show/Hide GUI
    GUIWindow(GUI_PLAYERINFO).Visible = True
    GUIWindow(GUI_FRIENDS).Visible = False
    frmMain.lstFriends.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleFriendData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapItem(n)
        .PlayerName = buffer.ReadString
        .num = buffer.ReadLong
        .value = buffer.ReadLong
        .x = buffer.ReadLong
        .y = buffer.ReadLong
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ITEMS
            .lstIndex.AddItem I & ": " & Trim$(Item(I).name)
        Next
        
        .cmbSkill.Clear
        For I = 1 To MAX_SKILLS
            .cmbSkill.AddItem Trim$(Skill(I).name)
        Next I

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleComboEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Combos
        Editor = EDITOR_COMBO
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_COMBO
            If Combo(I).Item_1 > 0 And Combo(I).Item_2 > 0 Then
                If Combo(I).Item_1 > 1 Or Combo(I).Item_2 > 1 Then
                    .lstIndex.AddItem I & ": " & Trim$(Item(Combo(I).Item_1).name) & " + " & Trim$(Item(Combo(I).Item_2).name)
                End If
            Else
                .lstIndex.AddItem I & ": "
            End If
        Next
        
        .cmbSkill.Clear
        .cmbGSkill.Clear
        .cmbSkill.AddItem "None"
        .cmbGSkill.AddItem "None"
        For I = 1 To MAX_SKILLS
            .cmbSkill.AddItem Trim$(Skill(I).name)
            .cmbGSkill.AddItem Trim$(Skill(I).name)
        Next I
        .cmbSkill.ListIndex = 0
        .cmbGSkill.ListIndex = 0
        
        .cmbItems1.Clear
        .cmbItems2.Clear
        .cmbItems1.AddItem "None"
        .cmbItems2.AddItem "None"
        For I = 1 To MAX_ITEMS
            .cmbItems1.AddItem I & ": " & Trim$(Item(I).name)
            .cmbItems2.AddItem I & ": " & Trim$(Item(I).name)
        Next I
        .cmbItems1.ListIndex = 0
        .cmbItems2.ListIndex = 0

        .Show
        .lstIndex.ListIndex = 0
        ComboEditorInit
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleComboEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleAnimationEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem I & ": " & Trim$(Animation(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).Visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateCombo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ComboSize As Long
Dim ComboData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the item
    ComboSize = LenB(Combo(n))
    ReDim ComboData(ComboSize - 1)
    ComboData = buffer.ReadBytes(ComboSize)
    CopyMemory ByVal VarPtr(Combo(n)), ByVal VarPtr(ComboData(0)), ComboSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateCombo", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSkill(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim SkillSize As Long
Dim SkillData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the skill
    SkillSize = LenB(Skill(n))
    ReDim SkillData(SkillSize - 1)
    SkillData = buffer.ReadBytes(SkillSize)
    CopyMemory ByVal VarPtr(Skill(n)), ByVal VarPtr(SkillData(0)), SkillSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateSkill", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long, I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapNpc(n)
        .num = buffer.ReadLong
        .x = buffer.ReadLong
        .y = buffer.ReadLong
        .Dir = buffer.ReadLong
        .HPSetTo = buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For I = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(I).num > 0 Then
            Npc_HighIndex = I
            Exit For
        End If
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    Call ClearMapNpc(n)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_NPCS
            .lstIndex.AddItem I & ": " & Trim$(NPC(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_RESOURCES
            .lstIndex.AddItem I & ": " & Trim$(Resource(I).name)
        Next
        
        .cmbSkill.Clear
        .cmbSkillReq.Clear
        .cmbSkillReq.AddItem "None"
        For I = 1 To MAX_SKILLS
            .cmbSkill.AddItem Trim$(Skill(I).name)
            .cmbSkillReq.AddItem Trim$(Skill(I).name)
        Next I

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim x As Long
Dim y As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    x = buffer.ReadLong
    y = buffer.ReadLong
    n = buffer.ReadByte
    TempTile(x, y).DoorOpen = n
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SHOPS
            .lstIndex.AddItem I & ": " & Trim$(Shop(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopnum = buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SPELLS
            .lstIndex.AddItem I & ": " & Trim$(Spell(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    spellnum = buffer.ReadLong
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    
    ' Update the spells on the pic
    Set buffer = New clsBuffer
    buffer.WriteLong CSpells
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(I) = buffer.ReadLong
    Next
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For I = 0 To Resource_Index
            MapResource(I).ResourceState = buffer.ReadByte
            MapResource(I).x = buffer.ReadLong
            MapResource(I).y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    Call DrawPing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    x = buffer.ReadLong
    y = buffer.ReadLong
    With TempTile(x, y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleDoorAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, Message As String, Color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Message = buffer.ReadString
    Color = buffer.ReadLong
    tmpType = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong

    Set buffer = Nothing
    
    CreateActionMsg Message, Color, tmpType, x, y
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, Sprite As Long, I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong

    Set buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For I = 1 To MAX_BYTE
        If Blood(I).x = x And Blood(I).y = y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .x = x
        .y = y
        .Sprite = Sprite
        .timer = GetTickCount
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .x = buffer.ReadLong
        .y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).x, AnimInstance(AnimationIndex).y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MapNpcNum = buffer.ReadLong
    For I = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(I) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Slot = buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Long
Dim name As String
Dim Message As String
Dim colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = buffer.ReadString
    Header = buffer.ReadString
    saycolour = buffer.ReadLong
    
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                colour = White
            Case 1
                colour = DarkGrey
            Case 2
                colour = Cyan
            Case 3
                colour = BrightGreen
            Case 4
                colour = Yellow
        End Select
    Else
        colour = BrightRed
    End If

    AddText Header & name & ": " & Message, colour
        
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopnum = buffer.ReadLong
    
    Set buffer = Nothing
    
    OpenShop shopnum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    StunDuration = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_BANK
        Bank.Item(I).num = buffer.ReadLong
        Bank.Item(I).value = buffer.ReadLong
    Next
    
    InBank = True
    GUIWindow(GUI_BANK).Visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    InTrade = buffer.ReadLong
    GUIWindow(GUI_TRADE).Visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InTrade = 0
    GUIWindow(GUI_TRADE).Visible = False
    TradeStatus = vbNullString
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim dataType As Byte
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    dataType = buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For I = 1 To MAX_INV
            TradeYourOffer(I).num = buffer.ReadLong
            TradeYourOffer(I).value = buffer.ReadLong
        Next
        YourWorth = buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For I = 1 To MAX_INV
            TradeTheirOffer(I).num = buffer.ReadLong
            TradeTheirOffer(I).value = buffer.ReadLong
        Next
        TheirWorth = buffer.ReadLong & "g"
    End If
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim status As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    status = buffer.ReadByte
    
    Set buffer = Nothing
    
    Select Case status
        Case 0 ' clear
            TradeStatus = vbNullString
        Case 1 ' they've accepted
            TradeStatus = "Other player has accepted."
        Case 2 ' you've accepted
            TradeStatus = "Waiting for other player to accept."
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    For I = 1 To MAX_HOTBAR
        Hotbar(I).Slot = buffer.ReadLong
        Hotbar(I).sType = buffer.ReadByte
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim FS As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Player_HighIndex = buffer.ReadLong
    FS = buffer.ReadLong
    If FS <> 1 Then
        frmMain.BorderStyle = 2
    Else
        frmMain.BorderStyle = 1
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    
    PlayMapSound x, y, entityType, entityNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    theName = buffer.ReadString
    
    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    theName = buffer.ReadString
    
    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, I As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    inParty = buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For I = 1 To MAX_PARTY_MEMBERS
        Party.Member(I) = buffer.ReadLong
    Next
    Party.MemberCount = buffer.ReadLong
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim buffer As clsBuffer, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' which player?
    playerNum = buffer.ReadLong
    ' set vitals
    For I = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(I) = buffer.ReadLong
        Player(playerNum).Vital(I) = buffer.ReadLong
    Next
        Player(playerNum).name = buffer.ReadString
        SetPlayerName playerNum, Trim$(Player(playerNum).name)
    
    ' find the party number
    For I = 1 To MAX_PARTY_MEMBERS
        If Party.Member(I) = playerNum Then
            partyIndex = I
        End If
    Next
        
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnEventPage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long, I As Long, Z As Long, x As Long, y As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    id = buffer.ReadLong
    
    If id > Map.CurrentEvents Then
        Map.CurrentEvents = id
        ReDim Preserve Map.MapEvents(Map.CurrentEvents)
    End If

    With Map.MapEvents(id)
        .name = buffer.ReadString
        .Dir = buffer.ReadLong
        .ShowDir = .Dir
        .GraphicNum = buffer.ReadLong
        .GraphicType = buffer.ReadLong
        .GraphicX = buffer.ReadLong
        .GraphicX2 = buffer.ReadLong
        .GraphicY = buffer.ReadLong
        .GraphicY2 = buffer.ReadLong
        .MovementSpeed = buffer.ReadLong
        .Moving = 0
        .x = buffer.ReadLong
        .y = buffer.ReadLong
        .xOffset = 0
        .yOffset = 0
        .Position = buffer.ReadLong
        .Visible = buffer.ReadLong
        .WalkAnim = buffer.ReadLong
        .DirFix = buffer.ReadLong
        .Walkthrough = buffer.ReadLong
        .ShowName = buffer.ReadLong
    End With
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpawnEventPage", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long
Dim x As Long
Dim y As Long
Dim Dir As Long, ShowDir As Long
Dim Movement As Long, MovementSpeed As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    id = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    Dir = buffer.ReadLong
    ShowDir = buffer.ReadLong
    MovementSpeed = buffer.ReadLong
    
    If id > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(id)
        .x = x
        .y = y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 1
        .ShowDir = ShowDir
        .MovementSpeed = MovementSpeed
        

        Select Case Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    Dir = buffer.ReadLong
    
    If I > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(I)
        .Dir = Dir
        .ShowDir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventChat(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim Dir As Byte
Dim buffer As clsBuffer
Dim choices As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    EventReplyID = buffer.ReadLong
    EventReplyPage = buffer.ReadLong
    chatText = buffer.ReadString
    choices = buffer.ReadLong
    
    InEvent = True
    inChat = True
    
    If choices = 0 Then
        chatOnlyContinue = True
    Else
        chatOnlyContinue = False
        For I = 1 To choices
            chatOpt(I) = buffer.ReadString
        Next
    End If
    
    GUIWindow(GUI_EVENTCHAT).Visible = True
    AnotherChat = buffer.ReadLong
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventChat", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventStart(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InEvent = True

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventStart", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventEnd(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InEvent = False

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventEnd", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    str = buffer.ReadString
    
    StopMusic
    PlayMusic str
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    str = buffer.ReadString

    PlaySound str, -1, -1
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'Need to learn how to fadeout :P
    'do later... way later.. like, after release, maybe never
    StopMusic
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStopSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 0 To UBound(Sounds()) - 1
        StopSound (I)
    Next
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_SWITCHES
        Switches(I) = buffer.ReadString
    Next
    
    For I = 1 To MAX_VARIABLES
        Variables(I) = buffer.ReadString
    Next
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapEventData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, I As Long, x As Long, y As Long, Z As Long, w As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    'Event Data!
    Map.EventCount = buffer.ReadLong
        
    If Map.EventCount > 0 Then
        ReDim Map.Events(0 To Map.EventCount)
        For I = 1 To Map.EventCount
            With Map.Events(I)
                .name = buffer.ReadString
                .Global = buffer.ReadLong
                .x = buffer.ReadLong
                .y = buffer.ReadLong
                .pageCount = buffer.ReadLong
            End With
            If Map.Events(I).pageCount > 0 Then
                ReDim Map.Events(I).Pages(0 To Map.Events(I).pageCount)
                For x = 1 To Map.Events(I).pageCount
                    With Map.Events(I).Pages(x)
                        .chkVariable = buffer.ReadLong
                        .VariableIndex = buffer.ReadLong
                        .VariableCondition = buffer.ReadLong
                        .VariableCompare = buffer.ReadLong
                            
                        .chkSwitch = buffer.ReadLong
                        .SwitchIndex = buffer.ReadLong
                        .SwitchCompare = buffer.ReadLong
                            
                        .chkHasItem = buffer.ReadLong
                        .HasItemIndex = buffer.ReadLong
                        .HasItemAmount = buffer.ReadLong
                            
                        .chkSelfSwitch = buffer.ReadLong
                        .SelfSwitchIndex = buffer.ReadLong
                        .SelfSwitchCompare = buffer.ReadLong
                            
                        .GraphicType = buffer.ReadLong
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .GraphicX2 = buffer.ReadLong
                        .GraphicY2 = buffer.ReadLong
                            
                        .MoveType = buffer.ReadLong
                        .MoveSpeed = buffer.ReadLong
                        .MoveFreq = buffer.ReadLong
                            
                        .MoveRouteCount = buffer.ReadLong
                        
                        .IgnoreMoveRoute = buffer.ReadLong
                        .RepeatMoveRoute = buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map.Events(I).Pages(x).MoveRoute(0 To .MoveRouteCount)
                            For y = 1 To .MoveRouteCount
                                .MoveRoute(y).Index = buffer.ReadLong
                                .MoveRoute(y).data1 = buffer.ReadLong
                                .MoveRoute(y).Data2 = buffer.ReadLong
                                .MoveRoute(y).Data3 = buffer.ReadLong
                                .MoveRoute(y).Data4 = buffer.ReadLong
                                .MoveRoute(y).Data5 = buffer.ReadLong
                                .MoveRoute(y).Data6 = buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = buffer.ReadLong
                        .DirFix = buffer.ReadLong
                        .Walkthrough = buffer.ReadLong
                        .ShowName = buffer.ReadLong
                        .Trigger = buffer.ReadLong
                        .CommandListCount = buffer.ReadLong
                            
                        .Position = buffer.ReadLong
                    End With
                        
                    If Map.Events(I).Pages(x).CommandListCount > 0 Then
                        ReDim Map.Events(I).Pages(x).CommandList(0 To Map.Events(I).Pages(x).CommandListCount)
                        For y = 1 To Map.Events(I).Pages(x).CommandListCount
                            Map.Events(I).Pages(x).CommandList(y).CommandCount = buffer.ReadLong
                            Map.Events(I).Pages(x).CommandList(y).ParentList = buffer.ReadLong
                            If Map.Events(I).Pages(x).CommandList(y).CommandCount > 0 Then
                                ReDim Map.Events(I).Pages(x).CommandList(y).Commands(1 To Map.Events(I).Pages(x).CommandList(y).CommandCount)
                                For Z = 1 To Map.Events(I).Pages(x).CommandList(y).CommandCount
                                    With Map.Events(I).Pages(x).CommandList(y).Commands(Z)
                                        .Index = buffer.ReadLong
                                        .Text1 = buffer.ReadString
                                        .Text2 = buffer.ReadString
                                        .Text3 = buffer.ReadString
                                        .Text4 = buffer.ReadString
                                        .Text5 = buffer.ReadString
                                        .data1 = buffer.ReadLong
                                        .Data2 = buffer.ReadLong
                                        .Data3 = buffer.ReadLong
                                        .Data4 = buffer.ReadLong
                                        .Data5 = buffer.ReadLong
                                        .Data6 = buffer.ReadLong
                                        .ConditionalBranch.CommandList = buffer.ReadLong
                                        .ConditionalBranch.Condition = buffer.ReadLong
                                        .ConditionalBranch.data1 = buffer.ReadLong
                                        .ConditionalBranch.Data2 = buffer.ReadLong
                                        .ConditionalBranch.Data3 = buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = buffer.ReadLong
                                        .MoveRouteCount = buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).Index = buffer.ReadLong
                                                .MoveRoute(w).data1 = buffer.ReadLong
                                                .MoveRoute(w).Data2 = buffer.ReadLong
                                                .MoveRoute(w).Data3 = buffer.ReadLong
                                                .MoveRoute(w).Data4 = buffer.ReadLong
                                                .MoveRoute(w).Data5 = buffer.ReadLong
                                                .MoveRoute(w).Data6 = buffer.ReadLong
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
    
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapEventData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, targetType As Long, target As Long, Message As String, colour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    target = buffer.ReadLong
    targetType = buffer.ReadLong
    Message = buffer.ReadString
    colour = buffer.ReadLong
    
    AddChatBubble target, targetType, Message, colour
    Set buffer = Nothing
ErrorHandler:
    HandleError "HandleChatBubble", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpecialEffect(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, effectType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    effectType = buffer.ReadLong
    
    Select Case effectType
        Case EFFECT_TYPE_FADEIN
            FadeType = 1
            FadeAmount = 0
        Case EFFECT_TYPE_FADEOUT
            FadeType = 0
            FadeAmount = 255
        Case EFFECT_TYPE_FLASH
            FlashTimer = GetTickCount + 150
        Case EFFECT_TYPE_FOG
            CurrentFog = buffer.ReadLong
            CurrentFogSpeed = buffer.ReadLong
            CurrentFogOpacity = buffer.ReadLong
        Case EFFECT_TYPE_WEATHER
            CurrentWeather = buffer.ReadLong
            CurrentWeatherIntensity = buffer.ReadLong
        Case EFFECT_TYPE_TINT
            CurrentTintR = buffer.ReadLong
            CurrentTintG = buffer.ReadLong
            CurrentTintB = buffer.ReadLong
            CurrentTintA = buffer.ReadLong
    End Select
    Set buffer = Nothing
ErrorHandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFlash(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, target As Long, n As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    target = buffer.ReadLong
    n = buffer.ReadByte
    If n = 1 Then
        MapNpc(target).StartFlash = GetTickCount + 200
    Else
        Player(target).StartFlash = GetTickCount + 200
    End If
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleFlash", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleCharEditorRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Command As Byte

    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Command = buffer.ReadByte
    Select Case Command
        Case 1: ' Get Player Info
            frmEditor_Character.txtELvl.text = buffer.ReadLong
            frmEditor_Character.txtEExp.text = buffer.ReadLong
            frmEditor_Character.txtEPts.text = buffer.ReadLong
            frmEditor_Character.txtEEnd.text = buffer.ReadLong
            frmEditor_Character.txtEStr.text = buffer.ReadLong
            frmEditor_Character.txtEInt.text = buffer.ReadLong
            frmEditor_Character.txtEAgi.text = buffer.ReadLong
            frmEditor_Character.txtEWill.text = buffer.ReadLong
            frmEditor_Character.txtESkillNum.text = 1
            frmEditor_Character.txtESkillLvl.text = buffer.ReadByte
            frmEditor_Character.txtESkillExp.text = buffer.ReadLong
            frmEditor_Character.txtEInvNum.text = 1
            frmEditor_Character.txtEItemNum.text = buffer.ReadLong
            frmEditor_Character.txtEItemQty.text = buffer.ReadLong
            frmEditor_Character.txtEBankNum.text = 1
            frmEditor_Character.txtEBItemNum.text = buffer.ReadLong
            frmEditor_Character.txtEBItemQty.text = buffer.ReadLong
        Case 2: ' Get Combat Info
            frmEditor_Character.txtESkillLvl.text = buffer.ReadByte
            frmEditor_Character.txtESkillExp.text = buffer.ReadLong
        Case 3: ' Get Inventory Info
            frmEditor_Character.txtEItemNum.text = buffer.ReadLong
            frmEditor_Character.txtEItemQty.text = buffer.ReadLong
        Case 4: ' Get Bank Info
            frmEditor_Character.txtEBItemNum.text = buffer.ReadLong
            frmEditor_Character.txtEBItemQty.text = buffer.ReadLong
    End Select
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCharEditorRequest", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleQuestEditor(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long

    With frmEditor_Quest
        Editor = EDITOR_TASKS
        .lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_QUESTS
            .lstIndex.AddItem I & ": " & Trim$(Quest(I).name)
        Next
        
        ' Update skill combo
        For I = 1 To MAX_SKILLS
            frmEditor_Quest.cmbSkill.AddItem Trim$(Skill(I).name)
        Next I
        
        .Show
        .lstIndex.ListIndex = 0
        frmEditor_Quest.cmbSkill.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Private Sub HandleUpdateQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim QuestNum As Long
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte
    On Error GoTo QuestErrHandler
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    ClearQuest QuestNum
    
    CopyMemory ByVal VarPtr(Quest(QuestNum)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
    Exit Sub
QuestErrHandler:
    MsgBox "modHandleData_HandleUpdateQuest CopyMemory error.", vbOKOnly, "Error"
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_QUESTS
        Player(MyIndex).PlayerQuest(I).status = buffer.ReadLong
        Player(MyIndex).PlayerQuest(I).ActualTask = buffer.ReadLong
        Player(MyIndex).PlayerQuest(I).CurrentCount = buffer.ReadLong
    Next
    
    RefreshQuestLog

Set buffer = Nothing
End Sub

Private Sub HandleQuestMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim I As Long, QuestNum As Long, QuestNumForStart As Long
    Dim Message As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong
    Message = Trim$(buffer.ReadString)
    QuestNumForStart = buffer.ReadLong
    
    QuestName = Trim$(Quest(QuestNum).name)
    QuestSay = Message
    QuestSubtitle = "Info:"
    inChat = True
    GUIWindow(GUI_QUESTDIALOGUE).Visible = True
    
    If QuestNumForStart > 0 And QuestNumForStart <= MAX_QUESTS Then
        QuestAcceptVisible = True
        QuestAcceptTag = QuestNumForStart
    End If
        
    Set buffer = Nothing
End Sub

Public Sub HandlePlayerCombatEXP(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim c As Long

    If Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
     
    For c = 1 To MAX_COMBAT
        Player(MyIndex).Combat(c).Level = buffer.ReadByte
        Player(MyIndex).Combat(c).EXP = buffer.ReadLong
        CombatTNL(c) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerCombatEXP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PlayerProjectile As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' create a new instance of the buffer
    Set buffer = New clsBuffer
    
    ' read bytes from data()
    buffer.WriteBytes data()
    
    ' recieve projectile number
    PlayerProjectile = buffer.ReadLong
    Index = buffer.ReadLong
    
    ' populate the values
    With Player(Index).ProjecTile(PlayerProjectile)
    
        ' set the direction
        .Direction = buffer.ReadLong
        
        ' set the direction to support file format
        Select Case .Direction
            Case DIR_DOWN
                .Direction = 0
            Case DIR_UP
                .Direction = 1
            Case DIR_RIGHT
                .Direction = 2
            Case DIR_LEFT
                .Direction = 3
        End Select
        
        ' set the pic
        .Pic = buffer.ReadLong
        ' set the coordinates
        .x = GetPlayerX(Index)
        .y = GetPlayerY(Index)
        ' get the range
        .Range = buffer.ReadLong
        ' get the damge
        .Damage = buffer.ReadLong
        ' get the speed
        .speed = buffer.ReadLong
        
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleProjectile", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
