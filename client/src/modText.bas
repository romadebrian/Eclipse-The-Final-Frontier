Attribute VB_Name = "modText"
Option Explicit
' Stuffs
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As DX8TextureRec
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single
'Text buffer

Public Type ChatTextBuffer
    text As String
    Color As Long
End Type

'Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Const FVF_SIZE As Long = 28

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal Color As Long, Optional ByVal alpha As Long = 0, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim I As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    ' set the color
    alpha = 255 - alpha
    Color = dx8Colour(Color, alpha)
    
    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    
    'Set the texture
    Direct3D_Device.SetTexture 0, gTexture(UseFont.Texture.Texture).Texture
    'CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For I = 0 To UBound(TempStr)
        If Len(TempStr(I)) > 0 Then
            yOffset = I * UseFont.CharHeight
            count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(I), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(I))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_SIZE * 4)
                
                'Set up the verticies
                TempVA(0).x = x + count
                TempVA(0).y = y + yOffset
                TempVA(1).x = TempVA(1).x + x + count
                TempVA(1).y = TempVA(0).y
                TempVA(2).x = TempVA(0).x
                TempVA(2).y = TempVA(2).y + TempVA(0).y
                TempVA(3).x = TempVA(1).x
                TempVA(3).y = TempVA(2).y
                
                'Set the colors
                TempVA(0).Color = TempColor
                TempVA(1).Color = TempColor
                TempVA(2).Color = TempColor
                TempVA(3).Color = TempColor
                
                'Draw the verticies
                Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)))
                
                'Shift over the the position to render the next character
                count = count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
            Next j
        End If
    Next I
End Sub

Sub EngineInitFontTextures()
    ' FONT DEFAULT
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Default.Texture.Texture = NumTextures
    Font_Default.Texture.filepath = App.Path & FONT_PATH & "texdefault.png"
    LoadTexture Font_Default.Texture
    
    ' Georgia
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Georgia.Texture.Texture = NumTextures
    Font_Georgia.Texture.filepath = App.Path & FONT_PATH & "georgia.png"
    LoadTexture Font_Georgia.Texture
End Sub

Sub UnloadFontTextures()
    UnloadFont Font_Default
    UnloadFont Font_Georgia
End Sub
Sub UnloadFont(Font As CustomFont)
    Font.Texture.Texture = 0
End Sub


Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal filename As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single


    'Load the header information
    FileNum = FreeFile
    Open App.Path & FONT_PATH & filename For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).x = 0
            .Vertex(0).y = 0
            .Vertex(0).Z = 0
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).x = theFont.HeaderInfo.CellWidth
            .Vertex(1).y = 0
            .Vertex(1).Z = 0
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).y = theFont.HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).x = theFont.HeaderInfo.CellWidth
            .Vertex(3).y = theFont.HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
    Next LoopChar
End Sub

Sub EngineInitFontSettings()
    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"
End Sub
Public Function dx8Colour(ByVal colourNum As Long, ByVal alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(alpha, 98, 84, 52)
        Case 17 'Orange
            dx8Colour = D3DColorARGB(alpha, 255, 96, 0)
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI

End Function

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim name As String
Dim Text2X As Long
Dim Text2Y As Long
Dim GuildString As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                Color = Orange
            Case 1
                Color = White
            Case 2
                Color = Cyan
            Case 3
                Color = BrightGreen
            Case 4
                Color = Yellow
        End Select

    Else
        Color = BrightRed
    End If

    If Options.Lvls = 1 Then
        name = Trim$(Player(Index).name & " [" & Player(Index).Level & "]")
    Else
        name = Trim$(Player(Index).name)
    End If
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    GuildString = Player(Index).GuildTag
    Text2X = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(GuildString))) / 2)
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - 16
        Text2Y = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (Tex_Character(GetPlayerSprite(Index)).Height / 4) + 16
        Text2Y = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (Tex_Character(GetPlayerSprite(Index)).Height / 4) + 4
    End If

    ' Draw name
    RenderText Font_Default, name, TextX, TextY, Color, 0
    If Not Player(Index).GuildName = vbNullString Then
        'Call DrawText(TexthDC, Text2X, Text2Y, GuildString, Color)
        RenderText Font_Default, GuildString, Text2X, Text2Y, Player(Index).GuildColor, 0
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim I As Long
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim name As String
Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    npcNum = MapNpc(Index).num

    Select Case NPC(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            Color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            Color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            Color = Grey
        Case Else
            Color = BrightGreen
    End Select
    
    name = Trim$(NPC(npcNum).name)
    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    If NPC(npcNum).Sprite < 1 Or NPC(npcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4) + 16
    End If

    ' Draw name
    RenderText Font_Default, name, TextX, TextY, Color, 0
    

    For I = 1 To MAX_QUESTS
    'check if the npc is the next task to any quest: [!] symbol
        If Quest(I).name <> "" Then
            If Player(MyIndex).PlayerQuest(I).status = QUEST_STARTED Then
                If Quest(I).Task(Player(MyIndex).PlayerQuest(I).ActualTask).NPC = npcNum Then
                    name = "[!]"
                    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - getWidth(Font_Default, (Trim$(name)))
                    If NPC(npcNum).Sprite < 1 Or NPC(npcNum).Sprite > NumCharacters Then
                        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - 16
                    Else
                        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4)
                    End If
                    RenderText Font_Default, name, TextX, TextY, Yellow
                    Exit For
                End If
            End If
        
            'check if the npc is the starter to any quest: [?] symbol
            'can accept the quest as a new one?
            If Player(MyIndex).PlayerQuest(I).status = QUEST_NOT_STARTED Or Player(MyIndex).PlayerQuest(I).status = QUEST_COMPLETED_BUT Then
                'the npc gives this quest?
                If NPC(npcNum).QuestNum = I Then
                    name = "[?]"
                    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - getWidth(Font_Default, (Trim$(name)))
                    If NPC(npcNum).Sprite < 1 Or NPC(npcNum).Sprite > NumCharacters Then
                        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - 16
                    Else
                        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4)
                    End If
                    RenderText Font_Default, name, TextX, TextY, Yellow
                    Exit For
                End If
            End If
        End If
    Next
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tx As Long
    Dim ty As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Map.optAttribs.value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    With Map.Tile(x, y)
                        tx = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Font_Default, "B", tx, ty, BrightRed, 0
                            Case TILE_TYPE_WARP
                                RenderText Font_Default, "W", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_ITEM
                                RenderText Font_Default, "I", tx, ty, White, 0
                            Case TILE_TYPE_NPCAVOID
                                RenderText Font_Default, "N", tx, ty, White, 0
                            Case TILE_TYPE_KEY
                                RenderText Font_Default, "K", tx, ty, White, 0
                            Case TILE_TYPE_KEYOPEN
                                RenderText Font_Default, "O", tx, ty, White, 0
                            Case TILE_TYPE_RESOURCE
                                RenderText Font_Default, "B", tx, ty, Green, 0
                            Case TILE_TYPE_DOOR
                                RenderText Font_Default, "D", tx, ty, Brown, 0
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Font_Default, "S", tx, ty, Yellow, 0
                            Case TILE_TYPE_SHOP
                                RenderText Font_Default, "S", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_BANK
                                RenderText Font_Default, "B", tx, ty, Blue, 0
                            Case TILE_TYPE_HEAL
                                RenderText Font_Default, "H", tx, ty, BrightGreen, 0
                            Case TILE_TYPE_TRAP
                                RenderText Font_Default, "T", tx, ty, BrightRed, 0
                            Case TILE_TYPE_SLIDE
                                RenderText Font_Default, "S", tx, ty, BrightCyan, 0
                            Case TILE_TYPE_SOUND
                                RenderText Font_Default, "S", tx, ty, Orange, 0
                            Case TILE_TYPE_PLAYERSPAWN
                                RenderText Font_Default, "PS", tx, ty, Pink, 0
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawActionMsg(ByVal Index As Long)
    Dim x As Long, y As Long, I As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For I = MAX_BYTE To 1 Step -1
                If ActionMsg(I).Type = ACTIONMSG_SCREEN Then
                    If I <> Index Then
                        ClearActionMsg Index
                        Index = I
                    End If
                End If
            Next
            x = (frmMain.ScaleWidth \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        RenderText Font_Default, ActionMsg(Index).Message, x, y, ActionMsg(Index).Color, 0
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(Font As CustomFont, ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    getWidth = EngineGetTextWidth(Font, text)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub DrawEventName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    If InMapEditor Then Exit Sub

    Color = White

    name = Trim$(Map.MapEvents(Index).name)
    
    ' calc pos
    TextX = ConvertMapX(Map.MapEvents(Index).x * PIC_X) + Map.MapEvents(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    If Map.MapEvents(Index).GraphicType = 0 Then
        TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
    ElseIf Map.MapEvents(Index).GraphicType = 1 Then
        If Map.MapEvents(Index).GraphicNum < 1 Or Map.MapEvents(Index).GraphicNum > NumCharacters Then
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
        Else
            ' Determine location for text
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - (Tex_Character(Map.MapEvents(Index).GraphicNum).Height / 4) + 16
        End If
    ElseIf Map.MapEvents(Index).GraphicType = 2 Then
        If Map.MapEvents(Index).GraphicY2 > 0 Then
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - ((Map.MapEvents(Index).GraphicY2 - Map.MapEvents(Index).GraphicY) * 32) + 16
        Else
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - 32 + 16
        End If
    End If

    ' Draw name
    RenderText Font_Default, name, TextX, TextY, Color, 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String, x As Long, y As Long, I As Long, MaxWidth As Long, x2 As Long, y2 As Long, colour As Long
    
    With chatBubble(Index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' it's on our map - get co-ords
                If Player(.target).GuildName = vbNullString Then
                    x = ConvertMapX((Player(.target).x * 32) + Player(.target).xOffset) + 16
                    y = ConvertMapY((Player(.target).y * 32) + Player(.target).yOffset) - 40
                Else
                    x = ConvertMapX((Player(.target).x * 32) + Player(.target).xOffset) + 16
                    y = ConvertMapY((Player(.target).y * 32) + Player(.target).yOffset) - 55
                End If
            End If
        ElseIf .targetType = TARGET_TYPE_NPC Then
            ' it's on our map - get co-ords
            x = ConvertMapX((MapNpc(.target).x * 32) + MapNpc(.target).xOffset) + 16
            y = ConvertMapY((MapNpc(.target).y * 32) + MapNpc(.target).yOffset) - 40
        ElseIf .targetType = TARGET_TYPE_EVENT Then
            x = ConvertMapX((Map.MapEvents(.target).x * 32) + Map.MapEvents(.target).xOffset) + 16
            y = ConvertMapY((Map.MapEvents(.target).y * 32) + Map.MapEvents(.target).yOffset) - 40
        End If
        
        ' word wrap the text
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
        ' find max width
        For I = 1 To UBound(theArray)
            If EngineGetTextWidth(Font_Default, theArray(I)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(I))
        Next
                
        ' calculate the new position
        x2 = x - (MaxWidth \ 2)
        y2 = y - (UBound(theArray) * 12)
                
        ' render bubble - top left
        RenderTexture Tex_GUI(25), x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_GUI(25), x2 + MaxWidth, y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_GUI(25), x2, y2 - 5, 10, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_GUI(25), x2 - 9, y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture Tex_GUI(25), x2 + MaxWidth, y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture Tex_GUI(25), x2, y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_GUI(25), x2 + (MaxWidth \ 2) + 6, y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_GUI(25), x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_GUI(25), x2 + MaxWidth, y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_GUI(25), x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_GUI(25), x - 5, y, 58, 19, 11, 11, 11, 11
                
        ' render each line centralised
        For I = 1 To UBound(theArray)
            RenderText Font_Georgia, theArray(I), x - (EngineGetTextWidth(Font_Default, theArray(I)) / 2), y2, DarkBrown
            y2 = y2 + 12
        Next
        ' check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

' Chat Box
Public Sub RenderChatTextBuffer()
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim I As Long

    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    Direct3D_Device.SetTexture 0, gTexture(Font_Default.Texture.Texture).Texture

    If ChatArrayUbound > 0 Then
        Direct3D_Device.SetStreamSource 0, ChatVBS, FVF_SIZE
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        Direct3D_Device.SetStreamSource 0, ChatVB, FVF_SIZE
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
    
End Sub

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim v As Single
Dim x As Single
Dim y As Single
Dim y2 As Single
Dim I As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    ' set the offset of each line
    yOffset = 14

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    
    Chunk = ChatScroll
    
    'Get the number of characters in all the visible buffer
    Size = 0
    
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        Size = Size + Len(ChatTextBuffer(LoopC).text)
    Next
    
    Size = Size - j
    ChatArrayUbound = Size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    'Set the base position
    x = GUIWindow(GUI_CHAT).x + ChatOffsetX
    y = GUIWindow(GUI_CHAT).y + ChatOffsetY

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        'Set the temp color
        TempColor = ChatTextBuffer(LoopC).Color
        
        'Set the Y position to be used
        y2 = y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        
        'Loop through each line if there are line breaks (vbCrLf)
        count = 0   'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).text) <> 0 Then  'Dont bother with empty strings
            
            'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).text)
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).text, j, 1))
                
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                v = Row * Font_Default.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + count
                    .y = (y2)
                    .TU = u
                    .TV = v
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + count
                    .y = (y2) + Font_Default.HeaderInfo.CellHeight
                    .TU = u
                    .TV = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + count + Font_Default.HeaderInfo.CellWidth
                    .y = (y2) + Font_Default.HeaderInfo.CellHeight
                    .TU = u + Font_Default.ColFactor
                    .TV = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                
                'Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * Pos)) = ChatVA(0 + (6 * Pos)) 'Top-left corner
                
                ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + count + Font_Default.HeaderInfo.CellWidth
                    .y = (y2)
                    .TU = u + Font_Default.ColFactor
                    .TV = v
                    .RHW = 1
                End With

                ChatVA(5 + (6 * Pos)) = ChatVA(2 + (6 * Pos))

                'Update the character we are on
                Pos = Pos + 1

                'Shift over the the position to render the next character
                count = count + Font_Default.HeaderInfo.CharWidth(Ascii)
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).Color
                End If
            Next
        End If
    Next LoopC
        
    If Not Direct3D_Device Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = Direct3D_Device.CreateVertexBuffer(FVF_SIZE * Pos * 6, 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_SIZE * Pos * 6, 0, ChatVAS(0)
        Set ChatVB = Direct3D_Device.CreateVertexBuffer(FVF_SIZE * Pos * 6, 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_SIZE * Pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
    
End Sub

Public Sub AddText(ByVal text As String, ByVal tColor As Long, Optional ByVal alpha As Long = 255)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim I As Long
Dim B As Long
Dim Color As Long

    Color = dx8Colour(tColor, alpha)

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Loop through all the characters
        For I = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), I, 1)
                Case " ": lastSpace = I
                Case "_": lastSpace = I
                Case "-": lastSpace = I
            End Select
            
            'Add up the size
            Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), I, 1)))
            
            'Check for too large of a size
            If Size > ChatWidth Then
                
                'Check if the last space was too far back
                If I - lastSpace > 10 Then
                
                    'Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, (I - 1) - B)), Color
                    B = I - 1
                    Size = 0
                Else
                    'Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)), Color
                    B = lastSpace + 1
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, I - lastSpace))
                End If
            End If
            
            'This handles the remainder
            If I = Len(TempSplit(TSLoop)) Then
                If B <> I Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), B, I), Color
            End If
        Next I
    Next TSLoop
    
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_Default.RowPitch = 0 Then Exit Sub
    
    If ChatScroll > 8 Then ChatScroll = ChatScroll + 1

    'Update the array
    UpdateChatArray
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal text As String, ByVal Color As Long)
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).text = text
    ChatTextBuffer(1).Color = Color
    
    ' set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
End Sub
Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, I As Long, Size As Long, lastSpace As Long, B As Long
    
    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If
    
    ' default values
    B = 1
    lastSpace = 1
    Size = 0
    
    For I = 1 To Len(text)
        ' if it's a space, store it
        Select Case Mid$(text, I, 1)
            Case " ": lastSpace = I
            Case "_": lastSpace = I
            Case "-": lastSpace = I
        End Select
        
        'Add up the size
        Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(text, I, 1)))
        
        'Check for too large of a size
        If Size > MaxLineLen Then
            'Check if the last space was too far back
            If I - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (I - 1) - B))
                B = I - 1
                Size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = EngineGetTextWidth(Font_Default, Mid$(text, lastSpace, I - lastSpace))
            End If
        End If
        
        ' Remainder
        If I = Len(text) Then
            If B <> I Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, I)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim I As Long
Dim B As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For I = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), I, 1)
                    Case " ": lastSpace = I
                    Case "_": lastSpace = I
                    Case "-": lastSpace = I
                End Select
    
                'Add up the size
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), I, 1)))
 
                'Check for too large of a size
                If Size > MaxLineLen Then
                    'Check if the last space was too far back
                    If I - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (I - 1) - B)) & vbNewLine
                        B = I - 1
                        Size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, I - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If I = Len(TempSplit(TSLoop)) Then
                    If B <> I Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, I)
                    End If
                End If
            Next I
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function
 

Public Sub UpdateShowChatText()
Dim CHATOFFSET As Long, I As Long, x As Long

    CHATOFFSET = 52
    
    If EngineGetTextWidth(Font_Default, MyText) > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
        For I = Len(MyText) To 1 Step -1
            x = x + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(MyText, I, 1)))
            If x > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
                RenderChatText = Right$(MyText, Len(MyText) - I + 1)
                Exit For
            End If
        Next
    Else
        RenderChatText = MyText
    End If
End Sub

