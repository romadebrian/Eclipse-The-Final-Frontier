Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public mouseClicked As Boolean
Public mouseState As GUIType

Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    If Not chatOn Then
        If Not GUIWindow(GUI_QUESTDIALOGUE).Visible Then 'fix quest and space key mix up
            If GetKeyState(vbKeySpace) < 0 Then
                CheckMapGetItem
            End If
        End If
        'Move Up
        If GetKeyState(vbKeyW) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyD) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetKeyState(vbKeyS) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyA) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
        'Move Up
        If GetKeyState(vbKeyUp) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If

    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim name As String
Dim I As Long
Dim n As Long
Dim Command() As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    chatText = MyText
    
    If GUIWindow(GUI_CURRENCY).Visible Then
        If (KeyAscii = vbKeyBack) Then
            If LenB(sDialogue) > 0 Then sDialogue = Mid$(sDialogue, 1, Len(sDialogue) - 1)
        End If
            
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
            sDialogue = sDialogue & ChrW$(KeyAscii)
        End If
    End If
        
    
    'GUI Window Toggles
    If GetKeyState(vbKeyV) < 0 And Not chatOn Then
        OpenGuiWindow 1
    End If
    If GetKeyState(vbKeyN) < 0 And Not chatOn Then
        OpenGuiWindow 1
    End If
    
    If GetKeyState(vbKeyE) < 0 And Not chatOn Then
        If GetPlayerAccess(MyIndex) >= ADMIN_MONITOR Then
            ' Weza gonna rotate through editors
            If Scroll_Editor < 10 Then
                Scroll_Editor = Scroll_Editor + 1
            Else
                Scroll_Editor = 1
            End If
            
            Call ScrollEditor
        End If
    End If
    If GetKeyState(vbKeyO) < 0 And Not chatOn Then
        OpenGuiWindow 2
    End If
    If GetKeyState(vbKeyC) < 0 And Not chatOn Then
        OpenGuiWindow 3
    End If
    If GetKeyState(vbKeyEscape) < 0 And Not chatOn Then
        OpenGuiWindow 4
    End If
    If GetKeyState(vbKeyT) < 0 And Not chatOn Then
        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            SendTradeRequest
        Else
            AddText "Invalid trade target.", BrightRed
        End If
    End If
    If GetKeyState(vbKeyP) < 0 And Not chatOn Then
        OpenGuiWindow 6
    End If
    If GetKeyState(vbKeyG) < 0 And Not chatOn Then
        OpenGuiWindow 7
    End If
    If GetKeyState(vbKeyQ) < 0 And Not chatOn Then
        OpenGuiWindow 8
    End If
    If GetKeyState(vbKeyK) < 0 And Not chatOn Then
        OpenGuiWindow 9
    End If
    If GetKeyState(vbKeyF) < 0 And Not chatOn Then
        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            Call SendFollowPlayer(myTarget)
        End If
    End If
    If GetKeyState(vbKeyB) < 0 And Not chatOn Then
        Dim SendClick As Boolean
        If Not myTargetType = TARGET_TYPE_PLAYER Or myTarget = MyIndex Then
            OpenGuiWindow 10
        ElseIf myTarget > 0 And myTarget <> MyIndex Then
            'For adding friends
            SendClick = True
            For I = 1 To Player(MyIndex).Friends.count
                If Trim$(Player(MyIndex).Friends.NameOfFriend(I)) = GetPlayerName(myTarget) Then
                    OpenGuiWindow 10
                    SendClick = False
                    Exit For
                End If
            Next I
            
            If SendClick Then Call SendDblClickPos
        End If
    End If
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        chatOn = Not chatOn
                
        'Guild Message
        If Left$(chatText, 1) = ";" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
        
            If Len(chatText) > 0 Then
                Call GuildMsg(chatText)
            End If
        
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If
        
        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(chatText)

                If Mid$(chatText, I, 1) <> Space(1) Then
                    name = name & Mid$(chatText, I, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, I, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) > 0 Then 'had - i)
                'I don't even know what this next line was meant for but it's useless
                'MyText = Mid$(chatText, i + 1, Len(chatText) - i)
                'MyText isn't used again other than to set it to nothing anyway.. people these days :p
                
                ' Send the message to the player
                Call PrivateMsg(chatText, name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /info, /who, /fps, /fpslock", HelpColor)
                
                Case "/guild"
                    If UBound(Command) < 1 Then
                        OpenGuiWindow 7
                        GoTo Continue
                    End If
                
                Select Case Command(1)
                    Case "help"
                        Call AddText("Guild Commands:", HelpColor)
                        Call AddText("Make Guild: /guild make (GuildName) (GuildTag)", HelpColor)
                        Call AddText("To transfer founder status use /guild founder (name)", HelpColor)
                        Call AddText("Invite to Guild: /guild invite (name)", HelpColor)
                        Call AddText("Leave Guild: /guild leave", HelpColor)
                        Call AddText("Open Guild Admin: /guild admin", HelpColor)
                        Call AddText("Guild kick: /guild kick (name)", HelpColor)
                        Call AddText("Guild disband: /guild disband yes", HelpColor)
                        Call AddText("View Guild: /guild view (online/all/offline)", HelpColor)
                        Call AddText("^Default is online, example: /guild view would display all online users.", HelpColor)
                        Call AddText("You can talk in guild chat with: ;Message ", HelpColor)
                    Case "make"
                        If UBound(Command) = 3 Then
                            Call GuildMake(Command(2), Command(3))
                        Else
                            Call AddText("Must have a name, use format /guild make (name)", BrightRed)
                        End If
                    
                    Case "invite"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(2, Command(2))
                        Else
                            Call AddText("Must select user, use format /guild invite (name)", BrightRed)
                        End If
                    
                    Case "leave"
                        Call GuildCommand(3, "")
                    
                    Case "admin"
                        Call GuildCommand(4, "")
                    
                    Case "view"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(5, Command(2))
                        Else
                            Call GuildCommand(5, "")
                        End If
                    
                    Case "accept"
                        Call GuildCommand(6, "")
                    
                    Case "decline"
                        Call GuildCommand(7, "")
                    
                    Case "founder"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(8, Command(2))
                        Else
                            Call AddText("Must select user, use format /guild founder (name)", BrightRed)
                        End If
                    Case "kick"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(9, Command(2))
                        Else
                            Call AddText("Must select user, use format /guild kick (name)", BrightRed)
                        End If
                    Case "disband"
                        If UBound(Command) = 2 Then
                            If LCase(Command(2)) = LCase("yes") Then
                                Call GuildCommand(10, "")
                            Else
                                Call AddText("Type like /guild disband yes (This is to help prevent an accident!)", BrightRed)
                            End If
                        Else
                            Call AddText("Type like /guild disband yes (This is to help prevent an accident!)", BrightRed)
                    End If
                
                End Select
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlayerInfoRequest
                    buffer.WriteString Command(1)
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set buffer = New clsBuffer
                    buffer.WriteLong CGetStats
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    frmMain.picAdmin.Visible = Not frmMain.picAdmin.Visible
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    'Walkthrough toggle
                Case "/walkthrough"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    SendWalkthrough
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    BLoc = Not BLoc
                    ' Spawn item
                Case "/spawnitem"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    If UBound(Command) < 2 Then
                        AddText "Usage: /spawnitem (ItemNum) (ItemValue)", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Trim$(Command(1))) Then
                        AddText "ItemNum must be numeric", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Trim$(Command(2))) Then
                        AddText "ItemValue must be numeric", AlertColor
                        GoTo Continue
                    End If
                    
                    SendSpawnItem CLng(Command(1)), CLng(Command(2))
                    ' Bank
                Case "/bank"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    SendOpenBankCommand
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1)), GetPlayerName(MyIndex)
                    'visibility toggle
                    Case "/visible"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    
                    SendVisibility
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' Killing a player
                Case "/kill"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                        If UBound(Command) < 1 Then
                            AddText "Usage: /kill (name)", AlertColor
                            GoTo Continue
                        End If

                    SendKillPlayer Command(1)
                    ' Level up player
                Case "/level"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                        If UBound(Command) < 1 Then
                            AddText "Usage: /level (name)", AlertColor
                            GoTo Continue
                        End If

                    SendRequestLevelUp Command(1)
                    ' // Developer Admin Commands //
                    ' Character Editor
                Case "/editchar"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    
                    frmEditor_Character.Visible = True
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditItem
                    ' Editing combo request
                Case "/editcombo"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditCombo
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditSpell
                Case "/editquest"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    SendRequestEditQuest
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(MyText)
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    If Not chatOn Then Exit Sub
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub HandleMouseMove(ByVal x As Long, ByVal y As Long, ByVal Button As Long)
Dim I As Long
    ' Set the global cursor position
    
    GlobalX = x
    GlobalY = y
    GlobalX_Map = (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (x >= GUIWindow(I).x And x <= GUIWindow(I).x + GUIWindow(I).Width) And (y >= GUIWindow(I).y And y <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS ', GUI_MENU
                            ' Put nothing here and we can click through them!
                        Case GUI_INVENTORY, GUI_SPELLS, GUI_CHARACTER, GUI_PARTY, GUI_OPTIONS, GUI_GUILD, GUI_QUESTLOG, GUI_COMBAT
                            ' Moveable GUI if right clicked
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' Handle the events
    CurX = TileView.Left + ((x + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, x, y)
        End If
    End If
    
    If mouseClicked Then
        If mouseState = GUI_INVENTORY Then
            GUIWindow(GUI_INVENTORY).x = GlobalX
            GUIWindow(GUI_INVENTORY).y = GlobalY
        ElseIf mouseState = GUI_SPELLS Then
            GUIWindow(GUI_SPELLS).x = GlobalX
            GUIWindow(GUI_SPELLS).y = GlobalY
        ElseIf mouseState = GUI_CHARACTER Then
            GUIWindow(GUI_CHARACTER).x = GlobalX
            GUIWindow(GUI_CHARACTER).y = GlobalY
        ElseIf mouseState = GUI_PARTY Then
            GUIWindow(GUI_PARTY).x = GlobalX
            GUIWindow(GUI_PARTY).y = GlobalY
        ElseIf mouseState = GUI_OPTIONS Then
            GUIWindow(GUI_OPTIONS).x = GlobalX
            GUIWindow(GUI_OPTIONS).y = GlobalY
        ElseIf mouseState = GUI_GUILD Then
            GUIWindow(GUI_GUILD).x = GlobalX
            GUIWindow(GUI_GUILD).y = GlobalY
        ElseIf mouseState = GUI_QUESTLOG Then
            GUIWindow(GUI_QUESTLOG).x = GlobalX
            GUIWindow(GUI_QUESTLOG).y = GlobalY
            frmMain.lstQuestLog.Left = (GUIWindow(GUI_QUESTLOG).x + (GUIWindow(GUI_QUESTLOG).Width / 2)) - (frmMain.lstQuestLog.Width / 2)
            frmMain.lstQuestLog.Top = GUIWindow(GUI_QUESTLOG).y + 10
        ElseIf mouseState = GUI_COMBAT Then
            GUIWindow(GUI_COMBAT).x = GlobalX
            GUIWindow(GUI_COMBAT).y = GlobalY
        ElseIf mouseState = GUI_FRIENDS Then
            GUIWindow(GUI_FRIENDS).x = GlobalX
            GUIWindow(GUI_FRIENDS).y = GlobalY
            frmMain.lstFriends.Left = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width / 2)) - (frmMain.lstFriends.Width / 2)
            frmMain.lstFriends.Top = GUIWindow(GUI_FRIENDS).y + 10
        End If
    End If
    
End Sub
Public Sub HandleMouseDown(ByVal Button As Long)
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).x And GlobalX <= GUIWindow(I).x + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).y And GlobalY <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_CHAT
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            mouseState = GUI_INVENTORY
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            mouseState = GUI_SPELLS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_MENU
                            If Options.Buttons = 1 Then
                                Menu_MouseDown Button
                            End If
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            mouseState = GUI_CHARACTER
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_CURRENCY
                            Currency_MouseDown
                            Exit Sub
                        Case GUI_DIALOGUE
                            Dialogue_MouseDown
                            Exit Sub
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            mouseState = GUI_PARTY
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_OPTIONS
                            Options_MouseDown
                            mouseState = GUI_OPTIONS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_TRADE
                            Trade_MouseDown
                            Exit Sub
                        Case GUI_EVENTCHAT
                            Chat_MouseDown
                            Exit Sub
                        Case GUI_GUILD
                            Guild_MouseDown
                            mouseState = GUI_GUILD
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_QUESTLOG
                            QuestLog_MouseDown
                            mouseState = GUI_QUESTLOG
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_QUESTDIALOGUE
                            QuestDialogue_MouseDown
                            Exit Sub
                        Case GUI_COMBAT
                            Combat_MouseDown
                            mouseState = GUI_COMBAT
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_FRIENDS
                            Buddies_MouseDown
                            mouseState = GUI_FRIENDS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_FRIENDREQUEST
                            FriendRequest_MouseDown
                            mouseState = GUI_FRIENDREQUEST
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_PLAYERINFO
                            PlayerInfo_MouseDown
                            mouseState = GUI_PLAYERINFO
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_BOOK
                            Book_MouseDown
                            mouseState = GUI_BOOK
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_BARS
                            Bars_MouseDown Button
                            mouseState = GUI_BARS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        'QUICKCHANGE
                        'Case Else
                        '    Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If Not inChat Then
            ChatScroll_MouseDown
        End If
    End If
    
    ' Handle events
    If InMapEditor Then
        Call MapEditorMouseDown(Button, GlobalX, GlobalY, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            Call PlayerSearch(CurX, CurY)
            'FindTarget
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If
    If frmEditor_Events.Visible Then frmEditor_Events.SetFocus
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).x And GlobalX <= GUIWindow(I).x + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).y And GlobalY <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_CHAT
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                            mouseClicked = False
                        Case GUI_SPELLS
                            Spells_MouseUp
                            mouseClicked = False
                        Case GUI_MENU
                            If Options.Buttons = 1 Then
                                Menu_MouseUp
                            End If
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                            mouseClicked = False
                        Case GUI_CURRENCY
                            Currency_MouseUp
                        Case GUI_DIALOGUE
                            Dialogue_MouseUp
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                            mouseClicked = False
                        Case GUI_OPTIONS
                            Options_MouseUp
                            mouseClicked = False
                        Case GUI_TRADE
                            Trade_MouseUp
                        Case GUI_EVENTCHAT
                            Chat_MouseUp
                        Case GUI_GUILD
                            Guild_MouseUp
                            mouseClicked = False
                        Case GUI_QUESTLOG
                            QuestLog_MouseUp
                            mouseClicked = False
                        Case GUI_QUESTDIALOGUE
                            QuestDialogue_MouseUp
                        Case GUI_COMBAT
                            Combat_MouseUp
                            mouseClicked = False
                        Case GUI_FRIENDS
                            Buddies_MouseUp
                            mouseClicked = False
                        Case GUI_FRIENDREQUEST
                            FriendRequest_MouseUp
                            mouseClicked = False
                        Case GUI_PLAYERINFO
                            PlayerInfo_MouseUp
                            mouseClicked = False
                        Case GUI_BARS
                            Bars_MouseUp
                            mouseClicked = False
                    End Select
                End If
            End If
        Next
    End If

    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub

Public Sub HandleSingleClick()
    If Not InMapEditor And Not hideGUI Then
        If (GlobalX >= GUIWindow(GUI_INVENTORY).x And GlobalX <= GUIWindow(GUI_INVENTORY).x + GUIWindow(GUI_INVENTORY).Width) And (GlobalY >= GUIWindow(GUI_INVENTORY).y And GlobalY <= GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_INVENTORY).Height) Then
            If GUIWindow(GUI_INVENTORY).Visible Then
                Inventory_SingleClick
            End If
        End If
    End If
End Sub

Public Sub HandleDoubleClick()
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).x And GlobalX <= GUIWindow(I).x + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).y And GlobalY <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                            Exit Sub
                        Case GUI_SHOP
                            Shop_DoubleClick
                            Exit Sub
                        Case GUI_BANK
                            Bank_DoubleClick
                            Exit Sub
                        Case GUI_TRADE
                            Trade_DoubleClick
                            Exit Sub
                        'Case Else
                        '    Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal Index As Long)
Dim buffer As clsBuffer
    If Index = 1 Then
        GUIWindow(GUI_INVENTORY).Visible = Not GUIWindow(GUI_INVENTORY).Visible
        Call InvHidden
    Else
        GUIWindow(GUI_INVENTORY).Visible = False
        Call InvHidden
    End If
    
    If Index = 2 Then
        GUIWindow(GUI_SPELLS).Visible = Not GUIWindow(GUI_SPELLS).Visible
        ' Update the spells on the pic
        Set buffer = New clsBuffer
        buffer.WriteLong CSpells
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        GUIWindow(GUI_SPELLS).Visible = False
    End If
    
    If Index = 3 Then
        GUIWindow(GUI_CHARACTER).Visible = Not GUIWindow(GUI_CHARACTER).Visible
    Else
        GUIWindow(GUI_CHARACTER).Visible = False
    End If
    
    If Index = 4 Then
        GUIWindow(GUI_OPTIONS).Visible = Not GUIWindow(GUI_OPTIONS).Visible
    Else
        GUIWindow(GUI_OPTIONS).Visible = False
    End If
    
    If Index = 6 Then
        GUIWindow(GUI_PARTY).Visible = Not GUIWindow(GUI_PARTY).Visible
    Else
        GUIWindow(GUI_PARTY).Visible = False
    End If
    
    If Index = 7 Then
        If Not Player(MyIndex).GuildName = vbNullString Then
            GUIWindow(GUI_GUILD).Visible = Not GUIWindow(GUI_GUILD).Visible
        End If
    Else
        GUIWindow(GUI_GUILD).Visible = False
    End If
    
    If Index = 8 Then
        GUIWindow(GUI_QUESTLOG).Visible = Not GUIWindow(GUI_QUESTLOG).Visible
        frmMain.lstQuestLog.Visible = Not frmMain.lstQuestLog.Visible
        UpdateQuestLog
    Else
        GUIWindow(GUI_QUESTLOG).Visible = False
        frmMain.lstQuestLog.Visible = False
    End If
    
    If Index = 9 Then
        GUIWindow(GUI_COMBAT).Visible = Not GUIWindow(GUI_COMBAT).Visible
    Else
        GUIWindow(GUI_COMBAT).Visible = False
    End If
    
    If Index = 10 Then
        GUIWindow(GUI_FRIENDS).Visible = Not GUIWindow(GUI_FRIENDS).Visible
        frmMain.lstFriends.Visible = Not frmMain.lstFriends.Visible
        UpdateFriendsList
    Else
        GUIWindow(GUI_FRIENDS).Visible = False
        frmMain.lstFriends.Visible = False
    End If
    
End Sub

Public Sub Currency_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    x = GUIWindow(GUI_CURRENCY).x + 155
    y = GUIWindow(GUI_CURRENCY).y + 96
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        CurrencyAcceptState = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    x = GUIWindow(GUI_CURRENCY).x + 218
    y = GUIWindow(GUI_CURRENCY).y + 96
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        CurrencyCloseState = 2 ' clicked
    End If
End Sub
Public Sub Currency_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long, buffer As clsBuffer
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    x = GUIWindow(GUI_CURRENCY).x + 155
    y = GUIWindow(GUI_CURRENCY).y + 96
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        If CurrencyAcceptState = 2 Then
            ' do stuffs
            If IsNumeric(sDialogue) Then
                Select Case CurrencyMenu
                    Case 1 ' drop item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        SendDropItem tmpCurrencyItem, Val(sDialogue)
                    Case 2 ' deposit item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        DepositItem tmpCurrencyItem, Val(sDialogue)
                    Case 3 ' withdraw item
                        If Val(sDialogue) > GetBankItemValue(tmpCurrencyItem) Then sDialogue = GetBankItemValue(tmpCurrencyItem)
                        WithdrawItem tmpCurrencyItem, Val(sDialogue)
                    Case 4 ' offer trade item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        TradeItem tmpCurrencyItem, Val(sDialogue)
                End Select
            Else
                AddText "Please enter a valid amount.", BrightRed
                Exit Sub
            End If
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    x = GUIWindow(GUI_CURRENCY).x + 218
    y = GUIWindow(GUI_CURRENCY).y + 96
    ' check if we're on the button
    If (GlobalX >= x And GlobalX <= x + Buttons(12).Width) And (GlobalY >= y And GlobalY <= y + Buttons(12).Height) Then
        If CurrencyCloseState = 2 Then
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    CurrencyAcceptState = 0
    CurrencyCloseState = 0
    GUIWindow(GUI_CURRENCY).Visible = False
    inChat = False
    chatOn = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    sDialogue = vbNullString
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Dialogue_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 90
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            Dialogue_ButtonState(1) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            Dialogue_ButtonState(2) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            Dialogue_ButtonState(3) = 2 ' clicked
        End If
    End If
End Sub

Public Sub Dialogue_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        x = GUIWindow(GUI_CHAT).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_CHAT).y + 90
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If Dialogue_ButtonState(1) = 2 Then
                Dialogue_Button_MouseDown (2)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(1) = 0
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        x = GUIWindow(GUI_CHAT).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_CHAT).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If Dialogue_ButtonState(2) = 2 Then
                Dialogue_Button_MouseDown (1)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(2) = 0
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        x = GUIWindow(GUI_CHAT).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_CHAT).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If Dialogue_ButtonState(3) = 2 Then
                Dialogue_Button_MouseDown (3)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(3) = 0
    End If
End Sub

' scroll bar
Public Sub ChatScroll_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    ' find out which button we're clicking
    For I = 34 To 35
        x = GUIWindow(GUI_CHAT).x + Buttons(I).x
        y = GUIWindow(GUI_CHAT).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
            ' scroll the actual chat
            Select Case I
                Case 34 ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case 35 ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 23 To 23
        x = GUIWindow(GUI_SHOP).x + Buttons(I).x
        y = GUIWindow(GUI_SHOP).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 23
                        ' exit
                        Set buffer = New clsBuffer
                        buffer.WriteLong CCloseShop
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        GUIWindow(GUI_SHOP).Visible = False
                        InShop = 0
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Shop_MouseDown()
Dim I As Long, x As Long, y As Long

    ' find out which button we're clicking
    For I = 23 To 23
        x = GUIWindow(GUI_SHOP).x + Buttons(I).x
        y = GUIWindow(GUI_SHOP).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub
Public Sub Bank_DoubleClick()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum <> 0 Then
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetBankItemNum(bankNum)).Stackable > 0 Then
            CurrencyMenu = 3 ' withdraw
            CurrencyText = "How many do you want withdraw?"
            tmpCurrencyItem = bankNum
            sDialogue = vbNullString
            GUIWindow(GUI_CURRENCY).Visible = True
            inChat = True
            chatOn = True
            Exit Sub
        End If
        WithdrawItem bankNum, 0
        Exit Sub
    End If
End Sub
Public Sub Trade_DoubleClick()
Dim tradeNum As Long
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum <> 0 Then
        UntradeItem tradeNum
        Exit Sub
    End If
End Sub
Public Sub Trade_MouseDown()
Dim I As Long, x As Long, y As Long

    ' find out which button we're clicking
    For I = 40 To 41
        x = Buttons(I).x
        y = Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Trade_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 40 To 41
        x = Buttons(I).x
        y = Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 40
                        AcceptTrade
                    Case 41
                        DeclineTrade
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

' Party
Public Sub Party_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 24 To 25
        x = GUIWindow(GUI_PARTY).x + Buttons(I).x
        y = GUIWindow(GUI_PARTY).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 24 ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText "Invalid invitation target.", BrightRed
                        End If
                    Case 25 ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText "You are not in a party.", BrightRed
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 24 To 25
        x = GUIWindow(GUI_PARTY).x + Buttons(I).x
        y = GUIWindow(GUI_PARTY).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Guild
Public Sub Guild_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 43 To 44
        x = GUIWindow(GUI_GUILD).x + Buttons(I).x
        y = GUIWindow(GUI_GUILD).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 43 ' Scroll Up
                        If GuildScroll > 1 Then GuildScroll = GuildScroll - 1
                    Case 44 ' Scroll Down
                        If GuildScroll < MAX_GUILD_MEMBERS - 4 And Not GuildData.Guild_Members(GuildScroll + 1).User_Name = vbNullString Then GuildScroll = GuildScroll + 1
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Guild_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 43 To 44
        x = GUIWindow(GUI_GUILD).x + Buttons(I).x
        y = GUIWindow(GUI_GUILD).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Combat
Public Sub Combat_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 52 To 53
        x = GUIWindow(GUI_COMBAT).x + Buttons(I).x
        y = GUIWindow(GUI_COMBAT).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 52 ' Scroll Up
                        If CombatScroll > 0 Then CombatScroll = CombatScroll - 1
                    Case 53 ' Scroll Down
                        'It's "- 4" beacuse you the first visible 4 don't count towards the scroll value
                        'You could also do: If CombatScroll + 4 < MAX_COMBAT + MAX_SKILLS Then
                        If CombatScroll < MAX_COMBAT + MAX_SKILLS - 4 Then CombatScroll = CombatScroll + 1
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Combat_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 52 To 53
        x = GUIWindow(GUI_COMBAT).x + Buttons(I).x
        y = GUIWindow(GUI_COMBAT).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For I = 26 To 33
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 3 Then
                ' do stuffs
                Select Case I
                    Case 26 ' music on
                        Options.Music = 1
                        PlayMusic Trim$(Map.Music)
                        SaveOptions
                        Buttons(26).state = 2
                        Buttons(27).state = 0
                    Case 27 ' music off
                        Options.Music = 0
                        StopMusic
                        SaveOptions
                        Buttons(26).state = 0
                        Buttons(27).state = 2
                    Case 28 ' sound on
                        Options.sound = 1
                        SaveOptions
                        Buttons(28).state = 2
                        Buttons(29).state = 0
                    Case 29 ' sound off
                        Options.sound = 0
                        StopAllSounds
                        SaveOptions
                        Buttons(28).state = 0
                        Buttons(29).state = 2
                    Case 30 ' debug on
                        Options.Debug = 1
                        SaveOptions
                        Buttons(30).state = 2
                        Buttons(31).state = 0
                    Case 31 ' debug off
                        Options.Debug = 0
                        SaveOptions
                        Buttons(30).state = 0
                        Buttons(31).state = 2
                    Case 32 ' levels on
                        Options.Lvls = 1
                        SaveOptions
                        Buttons(32).state = 2
                        Buttons(33).state = 0
                    Case 33 ' levels off
                        Options.Lvls = 0
                        SaveOptions
                        Buttons(32).state = 0
                        Buttons(33).state = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    For I = 59 To 62
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 3 Then
                ' do stuffs
                Select Case I
                    Case 59 ' minimap on
                        Options.MiniMap = 1
                        SaveOptions
                        Buttons(59).state = 2
                        Buttons(60).state = 0
                    Case 60 ' minimap off
                        Options.MiniMap = 0
                        SaveOptions
                        Buttons(59).state = 0
                        Buttons(60).state = 2
                    Case 61 ' buttons on
                        Options.Buttons = 1
                        SaveOptions
                        Buttons(61).state = 2
                        Buttons(62).state = 0
                    Case 62 ' buttons off
                        Options.Buttons = 0
                        SaveOptions
                        Buttons(61).state = 0
                        Buttons(62).state = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next I
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 26 To 33
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 0 Then
                Buttons(I).state = 3 ' clicked
            End If
        End If
    Next
    
    For I = 59 To 62
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 0 Then
                Buttons(I).state = 3 ' clicked
            End If
        End If
    Next
End Sub

' Menu
Public Sub Menu_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 1 To 7
        x = GUIWindow(GUI_MENU).x + Buttons(I).x
        y = GUIWindow(GUI_MENU).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 1
                        ' open window
                        OpenGuiWindow 1
                    Case 2
                        ' open window
                        OpenGuiWindow 2
                    Case 3
                        ' open window
                        OpenGuiWindow 3
                    Case 4
                        ' open window
                        OpenGuiWindow 4
                    Case 5
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendTradeRequest
                        Else
                            AddText "Invalid trade target.", BrightRed
                        End If
                    Case 6
                        ' open window
                        OpenGuiWindow 6
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' guild button
    I = 42
    x = GUIWindow(GUI_MENU).x + Buttons(I).x
    y = GUIWindow(GUI_MENU).y + Buttons(I).y
    ' check if we're on the button
    If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
        If Buttons(I).state = 2 Then
            ' do stuffs
            OpenGuiWindow 7
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 1 To 6
        If Buttons(I).Visible Then
            x = GUIWindow(GUI_MENU).x + Buttons(I).x
            y = GUIWindow(GUI_MENU).y + Buttons(I).y
            ' check if we're on the button
            If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
                Buttons(I).state = 2 ' clicked
            End If
        End If
    Next
    
    ' guild button
    I = 42
    If Buttons(I).Visible Then
        x = GUIWindow(GUI_MENU).x + Buttons(I).x
        y = GUIWindow(GUI_MENU).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    End If
End Sub

' HP/MP/SP GUI
Public Sub Bars_MouseDown(ByVal Button As Long)
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 63 To 64
        x = GUIWindow(GUI_BARS).x + Buttons(I).x
        y = GUIWindow(GUI_BARS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'If Buttons(I).state = 2 Then
                Select Case I
                    Case 63 ' minimap
                        Buttons(I).state = 2 'clicked
                    Case 64 ' buttons
                        Buttons(I).state = 2 'clicked
                End Select
                ' play sound - NO. double beeping is driving me crazy lol
                'PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    'resetClickedButtons
End Sub

' HP/MP/SP GUI
Public Sub Bars_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 63 To 64
        x = GUIWindow(GUI_BARS).x + Buttons(I).x
        y = GUIWindow(GUI_BARS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'If Buttons(I).state = 2 Then
                Select Case I
                    Case 63 ' minimap
                        Options.MiniMap = Abs(Not CBool(Options.MiniMap))
                        SaveOptions
                        Buttons(I).state = 0 'normal
                        
                        If CBool(Options.MiniMap) = True Then
                            Buttons(59).state = 2 ' normal
                            Buttons(60).state = 3 ' clicked
                        Else
                            Buttons(59).state = 3 ' clicked
                            Buttons(60).state = 2 ' normal
                        End If
                    Case 64 ' buttons
                        Options.Buttons = Abs(Not CBool(Options.Buttons))
                        SaveOptions
                        Buttons(I).state = 0 'normal
                        
                        If CBool(Options.Buttons) = True Then
                            Buttons(61).state = 2 ' normal
                            Buttons(62).state = 3 ' clicked
                        Else
                            Buttons(61).state = 3 ' clicked
                            Buttons(62).state = 2 ' normal
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        mouseClicked = False
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long

    invNum = IsInvItem(GlobalX, GlobalY)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        CurrencyText = "How many do you want to drop?"
                        tmpCurrencyItem = invNum
                        sDialogue = vbNullString
                        GUIWindow(GUI_CURRENCY).Visible = True
                        inChat = True
                        chatOn = True
                    End If
                Else
                    Call SendDropItem(invNum, 0)
                End If
            End If
        End If
    End If
End Sub

Public Sub Inventory_SingleClick()
    Dim invNum As Long, value As Long, multiplier As Double, I As Long
    
    ' Not if we're in a shop
    If InShop > 0 Then Exit Sub
    ' Not if in bank
    If InBank Then Exit Sub
    ' Not if in trade
    If InTrade > 0 Then Exit Sub
    
    'Check if selected
        invNum = IsInvItem(GlobalX, GlobalY)
        If invNum > 0 Then
            Call SendCheckHighlightItem(invNum)
            Exit Sub
        End If
    
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, value As Long, multiplier As Double, I As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX, GlobalY)

    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 2 ' deposit
                CurrencyText = "How many do you want to deposit?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).Visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For I = 1 To MAX_INV
                If TradeYourOffer(I).num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).Stackable > 0 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyText = "How many do you want to trade?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).Visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            If PlayerSpells(spellnum) > 0 Then
                Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum)).name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub
' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub

Public Sub Dialogue_Button_MouseDown(Index As Integer)
    ' call the handler
    dialogueHandler Index
    GUIWindow(GUI_DIALOGUE).Visible = False
    inChat = False
    dialogueIndex = 0
End Sub

Public Sub Character_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        x = GUIWindow(GUI_CHARACTER).x + Buttons(I).x
        y = GUIWindow(GUI_CHARACTER).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        x = GUIWindow(GUI_CHARACTER).x + Buttons(I).x
        y = GUIWindow(GUI_CHARACTER).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (I - 15)
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    Next
End Sub

' Npc Chat
Public Sub Chat_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long

If chatOnlyContinue = False Then
    For I = 1 To 4
        If Len(Trim$(chatOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(I)) & "]")
            x = GUIWindow(GUI_EVENTCHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_EVENTCHAT).y + 115 - ((I - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                chatOptState(I) = 2 ' clicked
            End If
        End If
    Next
Else
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    x = GUIWindow(GUI_EVENTCHAT).x + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    y = GUIWindow(GUI_EVENTCHAT).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        chatContinueState = 2 ' clicked
    End If
End If

End Sub
Public Sub Chat_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long

If chatOnlyContinue = False Then
    For I = 1 To 4
        If Len(Trim$(chatOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(I)) & "]")
            x = GUIWindow(GUI_EVENTCHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_EVENTCHAT).y + 115 - ((I - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                ' are we clicked?
                If chatOptState(I) = 2 Then
                    SendChatOption I
                    ' play sound
                    PlaySound Sound_ButtonClick, -1, -1
                End If
            End If
        End If
    Next
    
    For I = 1 To 4
        chatOptState(I) = 0 ' normal
    Next
Else
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    x = GUIWindow(GUI_EVENTCHAT).x + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    y = GUIWindow(GUI_EVENTCHAT).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        ' are we clicked?
        If chatContinueState = 2 Then
            SendChatContinue
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    chatContinueState = 0
End If
End Sub
Public Sub HandleKeyUp(ByVal KeyCode As Long)
Dim I As Long

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                frmMain.picAdmin.Visible = Not frmMain.picAdmin.Visible
            End If
            
    End Select
    
    ' hotbar
    If Not chatOn Then
        For I = 1 To 9
            If KeyCode = 48 + I Then
                SendHotbarUse I
            End If
        Next
        If KeyCode = 48 Then ' 0
            SendHotbarUse 10
        ElseIf KeyCode = 189 Then ' -
            SendHotbarUse 11
        ElseIf KeyCode = 187 Then ' =
            SendHotbarUse 12
        End If
    End If
    
    ' handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If
End Sub
Public Sub QuestLog_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 45 To 51
        x = GUIWindow(GUI_QUESTLOG).x + Buttons(I).x
        y = GUIWindow(GUI_QUESTLOG).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Buddies_MouseDown()
Dim x As Long, y As Long, I As Long
    For I = 54 To 55
        x = GUIWindow(GUI_FRIENDS).x + Buttons(I).x
        y = GUIWindow(GUI_FRIENDS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next I
End Sub
Public Sub FriendRequest_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If FriendRequestVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_FRIENDREQUEST).y + 80
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            FriendRequestAcceptState = 2 ' clicked
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Georgia, "[Decline]")
    x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_FRIENDREQUEST).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        FriendRequestDeclineState = 2 ' clicked
    End If
End Sub
Public Sub FriendRequest_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If Not FriendRequestVisible Then Exit Sub
    
    Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
    x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_FRIENDREQUEST).y + 80
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        FriendRequestAcceptState = 0 ' clicked
        ' Send accept packet to server
        Call SendAcceptFriend(FriendRequestSender)
        GUIWindow(GUI_FRIENDREQUEST).Visible = False
        FriendRequestVisible = False
        FriendRequestSender = vbNullString
    End If
    
    Width = EngineGetTextWidth(Font_Georgia, "[Decline]")
    x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_FRIENDREQUEST).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        FriendRequestDeclineState = 0 ' clicked
        ' Send decline packet to server
        Call SendDeclineFriend(FriendRequestSender)
        GUIWindow(GUI_FRIENDREQUEST).Visible = False
        FriendRequestVisible = False
        FriendRequestSender = vbNullString
    End If
End Sub
Public Sub Book_MouseDown()
Dim x As Long, y As Long, I As Long
Dim Parse() As String
    For I = 56 To 58
        x = GUIWindow(GUI_BOOK).x + Buttons(I).x
        y = GUIWindow(GUI_BOOK).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 0 ' normal
            
            Select Case I
                ' Left button
                Case 56
                    Book_PageLeft = True
                    Exit Sub
                ' Right button
                Case 57
                    Book_PageRight = True
                    Exit Sub
                ' X button
                Case 58
                    GUIWindow(GUI_BOOK).Visible = False
                    OpeningBook = True
                    Exit Sub
            End Select
        End If
    Next I
End Sub
Public Sub PlayerInfo_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Georgia, "[X]")
    x = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width - 25))
    y = GUIWindow(GUI_FRIENDS).y + 10
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        PlayerInfoX = 2 ' clicked
    End If
End Sub
Public Sub PlayerInfo_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Georgia, "[X]")
    x = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width - 25))
    y = GUIWindow(GUI_FRIENDS).y + 10
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        PlayerInfoX = 0 ' clicked
        GUIWindow(GUI_PLAYERINFO).Visible = False
        GUIWindow(GUI_FRIENDS).Visible = True
        frmMain.lstFriends.Visible = True
    End If
End Sub
Public Sub Buddies_MouseUp()
Dim x As Long, y As Long, I As Long
Dim Parse() As String
    For I = 54 To 55
        x = GUIWindow(GUI_FRIENDS).x + Buttons(I).x
        y = GUIWindow(GUI_FRIENDS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 0 ' normal
            
            Select Case I
                ' Defriend button
                Case 54
                    'is something selected and if so, does it have text?
                    If frmMain.lstFriends.ListIndex < 0 Then Exit Sub
                    If Not Len(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex)) > 0 Then Exit Sub
                    
                    ' Delete friend
                    SendDeleteFriend frmMain.lstFriends.List(frmMain.lstFriends.ListIndex)
                    Exit Sub
                ' Message button
                Case 55
                    If frmMain.lstFriends.ListIndex < 0 Then Exit Sub ' Nothing selected
                    If Not Len(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex)) > 0 Then Exit Sub 'No name in selection
                    If InStr(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex), "Offline") > 0 Then Exit Sub ' Player is offline

                    chatOn = True
                    Parse() = Split(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex), " ")
                    MyText = "!" & Parse(0) & " "
                    RenderChatText = MyText
                    Exit Sub
            End Select
        End If
    Next I
End Sub
Public Sub QuestLog_MouseUp()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 45 To 51
        x = GUIWindow(GUI_QUESTLOG).x + Buttons(I).x
        y = GUIWindow(GUI_QUESTLOG).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' send the level up
            If Trim$(frmMain.lstQuestLog.text) = vbNullString Then Exit Sub
                LoadQuestlogBox (I - 44)
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    Next
End Sub
Public Sub QuestAccept_MouseDown()
    PlayerHandleQuest CLng(QuestAcceptTag), 1
    inChat = False
    GUIWindow(GUI_QUESTDIALOGUE).Visible = False
    QuestAcceptVisible = False
    QuestAcceptTag = vbNullString
    QuestSay = "-"
    RefreshQuestLog
End Sub
Public Sub QuestExtra_MouseDown()
    RunQuestDialogueExtraLabel
End Sub

Public Sub QuestClose_MouseDown()
    inChat = False
    GUIWindow(GUI_QUESTDIALOGUE).Visible = False
    
    QuestExtraVisible = False
    QuestAcceptVisible = False
    QuestAcceptTag = vbNullString
    QuestSay = "-"
End Sub
Public Sub QuestDialogue_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            QuestAcceptState = 2 ' clicked
        End If
    End If
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            QuestExtraState = 2 ' clicked
        End If
    End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_CHAT).y + 120
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        QuestCloseState = 2 ' clicked
    End If
End Sub

Public Sub QuestDialogue_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If QuestAcceptState = 2 Then
                QuestAccept_MouseDown
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        QuestAcceptState = 0
    End If
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If QuestExtraState = 2 Then
                QuestExtra_MouseDown
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        QuestExtraState = 0
    End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_CHAT).y + 120
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        If QuestCloseState = 2 Then
            QuestClose_MouseDown
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    QuestCloseState = 0
End Sub
