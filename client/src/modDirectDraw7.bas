Attribute VB_Name = "modGraphics"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Private Direct3DX As D3DX8

'The 2D (Transformed and Lit) vertex format.
Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    RHW As Single
    Color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

Public ScreenWidth As Long
Public ScreenHeight As Long

'Graphic Textures
Public Tex_GUI() As DX8TextureRec
Public Tex_Buttons() As DX8TextureRec
Public Tex_Buttons_h() As DX8TextureRec
Public Tex_Buttons_c() As DX8TextureRec
Public Tex_Item() As DX8TextureRec ' arrays
Public Tex_Item_S() As DX8TextureRec
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Projectile() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Blood As DX8TextureRec ' singes
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Shadow As DX8TextureRec
Public Tex_MainMenu As DX8TextureRec

' Number of graphic files
Public NumGUIs As Long
Public NumButtons As Long
Public NumItems_S As Long
Public NumButtons_c As Long
Public NumButtons_h As Long
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumProjectiles As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    Loaded As Boolean
    UnloadTimer As Long
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set DirectX8 = Nothing
    Set Direct3D = Nothing
    Set Direct3DX = Nothing

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    ScreenWidth = 800
    ScreenHeight = 600
    
    'Buggy version of fullscreen
    'ScreenWidth = FormatNumber(frmMain.Width / 15.1125, 0)
    'ScreenHeight = FormatNumber(frmMain.Height / 15.7, 0)
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = ScreenWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = ScreenHeight 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hwnd 'Use frmMain as the device window.
    
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    ' Initialise the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "InitDX8", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TryCreateDirectX8Device() As Boolean
Dim I As Long

On Error GoTo nexti

    For I = 1 To 4
        Select Case I
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(value As Long) As Long
Dim I As Long
    Do While 2 ^ I < value
        I = I + 1
    Loop
    GetNearestPOT = 2 ^ I
End Function
Public Sub SetTexture(ByRef TextureRec As DX8TextureRec)
If TextureRec.Texture > NumTextures Then TextureRec.Texture = NumTextures
If TextureRec.Texture < 0 Then TextureRec.Texture = 0

If Not TextureRec.Texture = 0 Then
    If Not gTexture(TextureRec.Texture).Loaded Then
        Call LoadTexture(TextureRec)
    End If
End If

End Sub
Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, I As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            I = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, I, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(I)
            I = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, I)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (I)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    gTexture(TextureRec.Texture).Loaded = True
    gTexture(TextureRec.Texture).UnloadTimer = GetTickCount
    Exit Sub
ErrorHandler:
    HandleError "LoadTexture", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub LoadTextures()
Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call CheckGUIs
    Call CheckButtons
    Call CheckButtons_c
    Call CheckButtons_h
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckItems_S
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckProjectiles
    
    NumTextures = NumTextures + 10
    
    ReDim Preserve gTexture(NumTextures)
    Tex_Shadow.filepath = App.Path & "\data files\graphics\misc\shadow.png"
    Tex_Shadow.Texture = NumTextures - 9
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader.png"
    Tex_Fade.Texture = NumTextures - 8
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather.png"
    Tex_Weather.Texture = NumTextures - 7
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white.png"
    Tex_White.Texture = NumTextures - 6
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction.png"
    Tex_Direction.Texture = NumTextures - 5
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target.png"
    Tex_Target.Texture = NumTextures - 4
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc.png"
    Tex_Misc.Texture = NumTextures - 3
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood.png"
    Tex_Blood.Texture = NumTextures - 2
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select.png"
    Tex_Selection.Texture = NumTextures
    
    EngineInitFontTextures
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "LoadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures(Optional ByVal Complete As Boolean = False)
Dim I As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    If Complete = False Then
        For I = 1 To NumTextures
            If gTexture(I).UnloadTimer > GetTickCount + 150000 Then
                Set gTexture(I).Texture = Nothing
                ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
                gTexture(I).UnloadTimer = 0
                gTexture(I).Loaded = False
            End If
        Next
    Else
    
    For I = 1 To NumTextures
        Set gTexture(I).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
    Next
    
    ReDim gTexture(1)

    
    For I = 1 To NumTileSets
        Tex_Tileset(I).Texture = 0
    Next

    For I = 1 To numitems
        Tex_Item(I).Texture = 0
    Next

    For I = 1 To NumCharacters
        Tex_Character(I).Texture = 0
    Next
    
    For I = 1 To NumPaperdolls
        Tex_Paperdoll(I).Texture = 0
    Next
    
    For I = 1 To NumResources
        Tex_Resource(I).Texture = 0
    Next
    
    For I = 1 To NumAnimations
        Tex_Animation(I).Texture = 0
    Next
    
    For I = 1 To NumSpellIcons
        Tex_SpellIcon(I).Texture = 0
    Next
    
    For I = 1 To NumFaces
        Tex_Face(I).Texture = 0
    Next
    
    For I = 1 To NumProjectiles
        Tex_Projectile(I).Texture = 0
    Next
    
    For I = 1 To NumGUIs
        Tex_GUI(I).Texture = 0
    Next
    
    For I = 1 To NumButtons
        Tex_Buttons(I).Texture = 0
    Next
    
    For I = 1 To NumButtons_c
        Tex_Buttons_c(I).Texture = 0
    Next
    
    For I = 1 To NumButtons_c
        Tex_Item_S(I).Texture = 0
    Next
    
    For I = 1 To NumButtons_h
        Tex_Buttons_h(I).Texture = 0
    Next
    
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    Tex_Bars.Texture = 0
    Tex_White.Texture = 0
    Tex_Weather.Texture = 0
    Tex_Fade.Texture = 0
    Tex_Shadow.Texture = 0
    
    UnloadFontTextures
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UnloadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional Color As Long = -1)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    
    SetTexture TextureRec
    
    TextureNum = TextureRec.Texture
    
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, Color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, Color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, Color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, Color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, gTexture(TextureNum).Texture
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRECT As RECT, dRect As RECT)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    RenderTexture TextureRec, dRect.Left, dRect.Top, sRECT.Left, sRECT.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
Dim rec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' render grid
    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.Top + 32
    RenderTexture Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' render dir blobs
    For I = 1 To 4
        rec.Left = (I - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(x, y).DirBlock, CByte(I)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.Bottom = rec.Top + 8
        'render!
        RenderTexture Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(I), ConvertMapY(y * PIC_Y) + DirArrowY(I), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawDirection", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2)
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' clipping
    If y < 0 Then
        With sRECT
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With sRECT
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Target, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawTarget", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal target As Long, ByVal x As Long, ByVal y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2)

    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' clipping
    If y < 0 Then
        With sRECT
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With sRECT
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Target, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawHover", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapTile(ByVal x As Long, ByVal y As Long)
Dim rec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With Map.Tile(x, y)
        For I = MapLayer.Ground To MapLayer.Mask2
            If Autotile(x, y).layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .layer(I).x * 32, .layer(I).y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(x, y).layer(I).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile I, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile I, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile I, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile I, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
    
ErrorHandler:
    HandleError "DrawMapTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapFringeTile(ByVal x As Long, ByVal y As Long)
Dim rec As RECT
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With Map.Tile(x, y)
        For I = MapLayer.Fringe To MapLayer.Fringe2
            If Autotile(x, y).layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .layer(I).x * 32, .layer(I).y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(x, y).layer(I).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile I, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile I, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile I, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile I, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawMapFringeTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'load blood then
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        ' check if we should be seeing it
        If .timer + 20000 < GetTickCount Then Exit Sub
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawBlood", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal layer As Long)
Dim Sprite As Integer, sRECT As RECT, I As Long, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim x As Long, y As Long, lockindex As Long
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(layer)
    
    ' total width divided by frame count
    Width = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).width / frameCount
    Height = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).height
    
    With sRECT
        .Top = (Height * ((AnimInstance(Index).frameIndex(layer) - 1) \ AnimColumns))
        .Bottom = .Top + Height
        .Left = (Width * (((AnimInstance(Index).frameIndex(layer) - 1) Mod AnimColumns)))
        .Right = .Left + Width
    End With
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
    End If
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Animation(Sprite), x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
End Sub
Public Sub ScreenshotMap()
Dim x As Long, y As Long, I As Long, rec As RECT, drec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.picSSMap.Cls
    
    ' render the tiles
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            With Map.Tile(x, y)
                For I = MapLayer.Ground To MapLayer.Mask2
                    ' skip tile?
                    If (.layer(I).Tileset > 0 And .layer(I).Tileset <= NumTileSets) And (.layer(I).x > 0 Or .layer(I).y > 0) Then
                        ' sort out rec
                        rec.Top = .layer(I).y * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = .layer(I).x * PIC_X
                        rec.Right = rec.Left + PIC_X
                        
                        drec.Left = x * PIC_X
                        drec.Top = y * PIC_Y
                        drec.Right = drec.Left + (rec.Right - rec.Left)
                        drec.Bottom = drec.Top + (rec.Bottom - rec.Top)
                        ' render
                        RenderTextureByRects Tex_Tileset(.layer(I).Tileset), rec, drec
                    End If
                Next
            End With
        Next
    Next
    
    ' render the resources
    For y = 0 To Map.MaxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For I = 1 To Resource_Index
                        If MapResource(I).y = y Then
                            Call DrawMapResource(I, True)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' render the tiles
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            With Map.Tile(x, y)
                For I = MapLayer.Fringe To MapLayer.Fringe2
                    ' skip tile?
                    If (.layer(I).Tileset > 0 And .layer(I).Tileset <= NumTileSets) And (.layer(I).x > 0 Or .layer(I).y > 0) Then
                        ' sort out rec
                        rec.Top = .layer(I).y * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = .layer(I).x * PIC_X
                        rec.Right = rec.Left + PIC_X
                        
                        drec.Left = x * PIC_X
                        drec.Top = y * PIC_Y
                        drec.Right = drec.Left + (rec.Right - rec.Left)
                        drec.Bottom = drec.Top + (rec.Bottom - rec.Top)
                        ' render
                        RenderTextureByRects Tex_Tileset(.layer(I).Tileset), rec, drec
                    End If
                Next
            End With
        Next
    Next
    
    ' dump and save
    frmMain.picSSMap.Width = (Map.MaxX + 1) * 32
    frmMain.picSSMap.Height = (Map.MaxY + 1) * 32
    rec.Top = 0
    rec.Left = 0
    rec.Bottom = (Map.MaxX + 1) * 32
    rec.Right = (Map.MaxY + 1) * 32
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"
    
    ' let them know we did it
    AddText "Screenshot of map #" & GetPlayerMap(MyIndex) & " saved.", BrightGreen
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScreenshotMap", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As RECT
Dim x As Long, y As Long
Dim I As Long, alpha As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' make sure it's not out of map
    If MapResource(Resource_num).x > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).x, MapResource(Resource_num).y).data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    

    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
            If ConvertMapY(GetPlayerY(I)) < ConvertMapY(MapResource(Resource_num).y) And ConvertMapY(GetPlayerY(I)) > ConvertMapY(MapResource(Resource_num).y) - (Tex_Resource(Resource_sprite).Height) / 32 Then
                If ConvertMapX(GetPlayerX(I)) >= ConvertMapX(MapResource(Resource_num).x) - ((Tex_Resource(Resource_sprite).Width / 2) / 32) And ConvertMapX(GetPlayerX(I)) <= ConvertMapX(MapResource(Resource_num).x) + ((Tex_Resource(Resource_sprite).Width / 2) / 32) Then
                    alpha = 150
                Else
                    alpha = 255
                End If
            Else
                alpha = 255
            End If
        End If
    Next

    
    ' render it
    If Not screenShot Then
        Call DrawResource(Resource_sprite, alpha, x, y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, x, y, rec)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawMapResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal alpha As Long, ByVal dX As Long, dY As Long, rec As RECT)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    x = ConvertMapX(dX)
    y = ConvertMapY(dY)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Resource(Resource), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, alpha)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal x As Long, y As Long, rec As RECT)
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With rec
            .Left = .Left - x
        End With
        x = 0
    End If
    RenderTexture Tex_Resource(Resource), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScreenshotResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim I As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SetTexture Tex_Bars
    ' dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 4
    
    ' render health bars
    For I = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(I).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(I).Vital(Vitals.HP) > 0 And MapNpc(I).Vital(Vitals.HP) < MapNpc(I).HPSetTo Then
                ' lock to npc
                tmpX = MapNpc(I).x * PIC_X + MapNpc(I).xOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(I).y * PIC_Y + MapNpc(I).yOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(I).Vital(Vitals.HP) / sWidth) / (MapNpc(I).HPSetTo / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                
                ' draw the bar proper
                With sRECT
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            
            ' draw the bar proper
            With sRECT
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRECT
            .Top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
       
        ' draw the bar proper
        With sRECT
            .Top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For I = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(I)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).xOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).yOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        Next
    End If
                    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawBars", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayer(ByVal Index As Long)
Dim anim As Byte, I As Long, x As Long, y As Long
Dim Sprite As Long, spritetop As Long
Dim rec As RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).speed
    Else
        attackspeed = 1000
    End If

    If VXFRAME = False Then
        ' Reset frame
        If Player(Index).Step = 3 Then
            anim = 0
        ElseIf Player(Index).Step = 1 Then
            anim = 2
        End If
    Else
        anim = 1
    End If
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            If VXFRAME = False Then
                anim = 3
            Else
                anim = 2
            End If
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset > 8) Then anim = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).yOffset < -8) Then anim = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).xOffset > 8) Then anim = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).xOffset < -8) Then anim = Player(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        If VXFRAME = False Then
            .Left = anim * (Tex_Character(Sprite).Width / 4)
            .Right = .Left + (Tex_Character(Sprite).Width / 4)
        Else
            .Left = anim * (Tex_Character(Sprite).Width / 3)
            .Right = .Left + (Tex_Character(Sprite).Width / 3)
        End If
    End With

    ' Calculate the X
    If VXFRAME = False Then
        x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    Else
        x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    End If
    
    ' render player shadow
    RenderTexture Tex_Shadow, ConvertMapX(x), ConvertMapY(y + 18), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    ' render the actual sprite
    If GetTickCount > Player(Index).StartFlash Then
        Call DrawSprite(Sprite, x, y, rec)
        Player(Index).StartFlash = 0
    Else
        Call DrawSprite(Sprite, x, y, rec, True)
    End If
    
    ' check for paperdolling
    For I = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(I)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll > 0 Then
                Call DrawPaperdoll(x, y, Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll, anim, spritetop)
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawPlayer", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
Dim anim As Byte, I As Long, x As Long, y As Long, Sprite As Long, spritetop As Long
Dim rec As RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    
    Sprite = NPC(MapNpc(MapNpcNum).num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    attackspeed = 1000

    ' Reset frame
    anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            If VXFRAME = False Then
                anim = 3
            Else
                anim = 2
            End If
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < -8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < -8) Then anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        If VXFRAME = False Then
            .Left = anim * (Tex_Character(Sprite).Width / 4)
            .Right = .Left + (Tex_Character(Sprite).Width / 4)
        Else
            .Left = anim * (Tex_Character(Sprite).Width / 3)
            .Right = .Left + (Tex_Character(Sprite).Width / 3)
        End If
    End With

    ' Calculate the X
    If VXFRAME = False Then
        x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    Else
        x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset
    End If
    
    ' render player shadow
    RenderTexture Tex_Shadow, ConvertMapX(x), ConvertMapY(y + 18), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    ' render the actual sprite
    If GetTickCount > MapNpc(MapNpcNum).StartFlash Then
        Call DrawSprite(Sprite, x, y, rec)
        MapNpc(MapNpcNum).StartFlash = 0
    Else
        Call DrawSprite(Sprite, x, y, rec, True)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal anim As Long, ByVal spritetop As Long)
Dim rec As RECT
Dim x As Long, y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    With rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        If VXFRAME = False Then
            .Left = anim * (Tex_Paperdoll(Sprite).Width / 4)
            .Right = .Left + (Tex_Paperdoll(Sprite).Width / 4)
        Else
            .Left = anim * (Tex_Paperdoll(Sprite).Width / 3)
            .Right = .Left + (Tex_Paperdoll(Sprite).Width / 3)
        End If
    End With
    
    ' clipping
    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With rec
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Paperdoll(Sprite), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As RECT, Optional Flash As Boolean = False)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    If Flash = True Then
        RenderTexture Tex_Character(Sprite), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 150)
    Else
        RenderTexture Tex_Character(Sprite), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, Color As Long, x As Long, y As Long, RenderState As Long

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    Color = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)

    RenderState = 0
    ' render state
    Select Case RenderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For x = 0 To ((Map.MaxX * 32) / 256) + 1
        For y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((x * 256) + fogOffsetX), ConvertMapY((y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Color
        Next
    Next
    
    ' reset render state
    If RenderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
Dim Color As Long
    Color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, Color
End Sub

Public Sub DrawWeather()
Dim Color As Long, I As Long, SpriteLeft As Long
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(I).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(I).x), ConvertMapY(WeatherParticle(I).y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Sub DrawAnimatedInvItems()
Dim I As Long
Dim itemNum As Long, ItemPic As Long
Dim x As Long, y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For I = 1 To MAX_MAP_ITEMS

        If MapItem(I).num > 0 Then
            ItemPic = Item(MapItem(I).num).Pic

            If ItemPic < 1 Or ItemPic > numitems Then Exit Sub
            MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(I).Frame < MaxFrames - 1 Then
                MapItem(I).Frame = MapItem(I).Frame + 1
            Else
                MapItem(I).Frame = 1
            End If
        End If

    Next

    For I = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, I)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic

            If ItemPic > 0 And ItemPic <= numitems Then
                If Tex_Item(ItemPic).Width > 64 Then
                    MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(I) < MaxFrames - 1 Then
                        InvItemFrame(I) = InvItemFrame(I) + 1
                    Else
                        InvItemFrame(I) = 1
                    End If

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = (Tex_Item(ItemPic).Width / 2) + (InvItemFrame(I) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' We'll now re-Draw the item, and place the currency value over it again :P
                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                        y = rec_pos.Top + 22
                        x = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, I))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        RenderText Font_Default, ConvertCurrency(Amount), x, y, Yellow, 0
                    End If
                End If
            End If
        End If

    Next

    'frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawAnimatedInvItems", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long
Dim sRECT As RECT
Dim dRect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmEditor_Map.scrlPictureX.value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.value * PIC_Y
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRECT.Left = frmEditor_Map.scrlPictureX.value * PIC_X
    sRECT.Top = frmEditor_Map.scrlPictureY.value * PIC_Y
    sRECT.Right = sRECT.Left + Width
    sRECT.Bottom = sRECT.Top + Height
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    RenderTextureByRects Tex_Tileset(Tileset), sRECT, dRect
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    
    With destRect
        .x1 = (EditorTileX * 32) - sRECT.Left
        .x2 = (EditorTileWidth * 32) + .x1
        .y1 = (EditorTileY * 32) - sRECT.Top
        .y2 = (EditorTileHeight * 32) + .y1
    End With
    
    DrawSelectionBox destRect
        
    With srcRect
        .x1 = 0
        .x2 = Width
        .y1 = 0
        .y2 = Height
    End With
                    
    With destRect
        .x1 = 0
        .x2 = frmEditor_Map.picBack.ScaleWidth
        .y1 = 0
        .y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    
    'Now render the selection tiles and we are done!
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picBack.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorMap_DrawTileset", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawSelectionBox(dRect As D3DRECT)
Dim Width As Long, Height As Long, x As Long, y As Long
    Width = dRect.x2 - dRect.x1
    Height = dRect.y2 - dRect.y1
    x = dRect.x1
    y = dRect.y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, x, y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Selection, x + 2, y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Selection, x, y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Selection, x, y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Selection, x + 2, y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawTileOutline()
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Map.optBlock.value Then Exit Sub

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawTileOutline", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterDrawSprite()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    SetTexture Tex_Character(Sprite)
    
    If VXFRAME = False Then
        Width = Tex_Character(Sprite).Width / 4
    Else
        Width = Tex_Character(Sprite).Width / 3
    End If
    
    Height = Tex_Character(Sprite).Height / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    
    With srcRect
        .x1 = 0
        .x2 = Width
        .y1 = 0
        .y2 = Height
    End With
                    
    With destRect
        .x1 = 0
        .x2 = Width
        .y1 = 0
        .y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMenu.picSprite.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "NewCharacterDrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawMapItem()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    itemNum = Item(frmEditor_Map.scrlMapItem.value).Pic

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemNum), sRECT, dRect
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapItem.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorMap_DrawMapItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawKey()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    itemNum = Item(frmEditor_Map.scrlMapKey.value).Pic

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    RenderTextureByRects Tex_Item(itemNum), sRECT, dRect
    
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapKey.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawItem()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    itemNum = frmEditor_Item.scrlPic.value

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemNum), sRECT, dRect
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picItem.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    'frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRECT.Left = 0
    If VXFRAME = False Then
        sRECT.Right = Tex_Paperdoll(Sprite).Width / 4
    Else
        sRECT.Right = Tex_Paperdoll(Sprite).Width / 3
    End If
    ' same for destination as source
    dRect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRECT, dRect
                    
    With destRect
        .x1 = 0
        If VXFRAME = False Then
            .x2 = Tex_Paperdoll(Sprite).Width / 4
        Else
            .x2 = Tex_Paperdoll(Sprite).Width / 3
        End If
        .y1 = 0
        .y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picPaperdoll.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorItem_DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
Dim iconnum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    iconnum = frmEditor_Spell.scrlIcon.value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(iconnum), sRECT, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picSprite.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorSpell_DrawIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_DrawAnim()
Dim I As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim sX As Long, sY As Long, sRECT As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    sRECT.Top = 0
    sRECT.Bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For I = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(I).value
        
        If Animationnum <= 0 Or Animationnum > NumAnimations Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(I)
            FrameCount = frmEditor_Animation.scrlFrameCount(I)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(I) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(I) >= FrameCount Then
                    AnimEditorFrame(I) = 1
                Else
                    AnimEditorFrame(I) = AnimEditorFrame(I) + 1
                End If
                AnimEditorTimer(I) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(I).value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    Height = 192

                    sY = (Height * ((AnimEditorFrame(I) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(I) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call Direct3D_Device.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call Direct3D_Device.BeginScene
                    
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Animation(Animationnum), 0, 0, sX, sY, Width, Height, Width, Height
                    
                    ' Finish Rendering
                    Call Direct3D_Device.EndScene
                    Call Direct3D_Device.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(I).hwnd, ByVal 0)
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorAnim_DrawAnim", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
Dim Sprite As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Sprite = frmEditor_NPC.scrlSprite.value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    dRect.Top = 0
    dRect.Bottom = SIZE_Y
    dRect.Left = 0
    dRect.Right = SIZE_X
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    
    With destRect
        .x1 = 0
        .x2 = SIZE_X
        .y1 = 0
        .y2 = SIZE_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picSprite.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorNpc_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawSprite()
Dim Sprite As Long
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        With srcRect
            .x1 = 0
            .x2 = Tex_Resource(Sprite).Width
            .y1 = 0
            .y2 = Tex_Resource(Sprite).Height
        End With
        
        With destRect
            .x1 = 0
            .x2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .y1 = 0
            .y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hwnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        
        With destRect
            .x1 = 0
            .x2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .y1 = 0
            .y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .x1 = 0
            .x2 = Tex_Resource(Sprite).Width
            .y1 = 0
            .y2 = Tex_Resource(Sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hwnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub Render_Graphics()
Dim x As Long
Dim y As Long
Dim I As Long
Dim rec As RECT
Dim rec_pos As RECT, srcRect As D3DRECT
Static Hidden As Boolean, Shown As Boolean
    
    ' If debug mode, handle error then exit out
   If Options.Debug Then On Error GoTo ErrorHandler
    
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera

    ' unload any textures we need to unload
    UnloadTextures
   Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
        
        Direct3D_Device.BeginScene
            ' blit lower tiles
            If NumTileSets > 0 Then
                For x = TileView.Left To TileView.Right
                    For y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(x, y) Then
                            Call DrawMapTile(x, y)
                        End If
                    Next
                Next
            End If
        
            ' render the decals
            For I = 1 To MAX_BYTE
                Call DrawBlood(I)
            Next
        
            ' Blit out the items
            If numitems > 0 Then
                For I = 1 To MAX_MAP_ITEMS
                    If MapItem(I).num > 0 Then
                        Call DrawItem(I)
                    End If
                Next
            End If
            
            If Map.CurrentEvents > 0 Then
                For I = 1 To Map.CurrentEvents
                    If Map.MapEvents(I).Position = 0 Then
                        DrawEvent I
                    End If
                Next
            End If
            
            ' draw animations
            If NumAnimations > 0 Then
                For I = 1 To MAX_BYTE
                    If AnimInstance(I).Used(0) Then
                        DrawAnimation I, 0
                    End If
                Next
            End If
            
            'LEFTOFF - Might need an if-statement
            ' draw projectiles for each player
            For I = 1 To Player_HighIndex
                For x = 1 To MAX_PLAYER_PROJECTILES
                    If Player(I).ProjecTile(x).Pic > 0 Then
                        DrawProjectile I, x
                    End If
                Next
            Next
        
            ' Y-based render. Renders Players, Npcs, and Resources based on Y-axis.
            For y = 0 To Map.MaxY
                If NumCharacters > 0 Then
                
                    If Map.CurrentEvents > 0 Then
                        For I = 1 To Map.CurrentEvents
                            If Map.MapEvents(I).Position = 1 Then
                                If y = Map.MapEvents(I).y Then
                                    DrawEvent I
                                End If
                            End If
                        Next
                    End If
                    
                    ' Players
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                            If Player(I).y = y And (Not GetPlayerVisible(I) = 1 Or I = MyIndex) Then
                                Call DrawPlayer(I)
                            End If
                        End If
                    Next
                    
                    
                
                    ' Npcs
                    For I = 1 To Npc_HighIndex
                        If MapNpc(I).y = y Then
                            Call DrawNpc(I)
                        End If
                    Next
                End If
                
                ' Resources
                If NumResources > 0 Then
                    If Resources_Init Then
                        If Resource_Index > 0 Then
                            For I = 1 To Resource_Index
                                If MapResource(I).y = y Then
                                    Call DrawMapResource(I)
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            
            ' animations
            If NumAnimations > 0 Then
                For I = 1 To MAX_BYTE
                    If AnimInstance(I).Used(1) Then
                        DrawAnimation I, 1
                    End If
                Next
            End If
        
            ' blit out upper tiles
            If NumTileSets > 0 Then
                For x = TileView.Left To TileView.Right
                    For y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(x, y) Then
                            Call DrawMapFringeTile(x, y)
                        End If
                    Next
                Next
            End If
            
            If Map.CurrentEvents > 0 Then
                For I = 1 To Map.CurrentEvents
                    If Map.MapEvents(I).Position = 2 Then
                        DrawEvent I
                    End If
                Next
            End If
            
            DrawWeather
            DrawFog
            DrawTint
            
            ' blit out a square at mouse cursor
            If InMapEditor Then
                If frmEditor_Map.optBlock.value = True Then
                    For x = TileView.Left To TileView.Right
                        For y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(x, y) Then
                                Call DrawDirection(x, y)
                            End If
                        Next
                    Next
                End If
                Call DrawTileOutline
            End If
            
            ' Render the bars
            DrawBars
            
            ' Draw the target icon
            If myTarget > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    DrawTarget (Player(myTarget).x * 32) + Player(myTarget).xOffset, (Player(myTarget).y * 32) + Player(myTarget).yOffset
                ElseIf myTargetType = TARGET_TYPE_NPC Then
                    DrawTarget (MapNpc(myTarget).x * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).y * 32) + MapNpc(myTarget).yOffset
                End If
            End If
            
            ' Draw the hover icon
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    If Player(I).Map = Player(MyIndex).Map Then
                        If CurX = Player(I).x And CurY = Player(I).y Then
                            If myTargetType = TARGET_TYPE_PLAYER And myTarget = I Or GetPlayerVisible(I) = 1 Then
                                ' dont render lol
                            Else
                                DrawHover TARGET_TYPE_PLAYER, I, (Player(I).x * 32) + Player(I).xOffset, (Player(I).y * 32) + Player(I).yOffset
                            End If
                        End If
                    End If
                End If
            Next
            For I = 1 To Npc_HighIndex
                If MapNpc(I).num > 0 Then
                    If CurX = MapNpc(I).x And CurY = MapNpc(I).y Then
                        If myTargetType = TARGET_TYPE_NPC And myTarget = I Then
                            ' dont render lol
                        Else
                            DrawHover TARGET_TYPE_NPC, I, (MapNpc(I).x * 32) + MapNpc(I).xOffset, (MapNpc(I).y * 32) + MapNpc(I).yOffset
                        End If
                    End If
                End If
            Next
            
            If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
            
            ' Get rec
            With rec
                .Top = Camera.Top
                .Bottom = .Top + ScreenY
                .Left = Camera.Left
                .Right = .Left + ScreenX
            End With
                
            ' rec_pos
            With rec_pos
                .Bottom = ScreenY
                .Right = ScreenX
            End With
                
            With srcRect
                .x1 = 0
                .x2 = frmMain.ScaleWidth
                .y1 = 0
                .y2 = frmMain.ScaleHeight
            End With
            
            If BFPS Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS), 12, 100, Yellow, 0
            End If
            
            ' draw cursor, player X and Y locations
            If BLoc Then
                RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 12, 114, Yellow, 0
                RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 12, 128, Yellow, 0
                RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 12, 142, Yellow, 0
            End If
            
            ' draw player names
            For I = 1 To Player_HighIndex
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) And (Not GetPlayerVisible(I) = 1 Or I = MyIndex) Then
                    Call DrawPlayerName(I)
                End If
            Next
            
            For I = 1 To Map.CurrentEvents
                If Map.MapEvents(I).Visible = 1 Then
                    If Map.MapEvents(I).ShowName = 1 Then
                        DrawEventName (I)
                    End If
                End If
            Next
            
            ' draw npc names
            For I = 1 To Npc_HighIndex
                If MapNpc(I).num > 0 Then
                    Call DrawNpcName(I)
                End If
            Next
            
                ' draw the messages
            For I = 1 To MAX_BYTE
                If chatBubble(I).active Then
                    DrawChatBubble I
                End If
            Next
            
            For I = 1 To Action_HighIndex
                Call DrawActionMsg(I)
            Next I
            
            ' Render the MiniMap / if not in map editor
            If Not InMapEditor Then
                If Options.MiniMap Then DrawMiniMap
                DrawGUI
                Hidden = False
                If Not Shown Then
                    ShowGame
                    Shown = True
                End If
            Else
                Shown = False
                If Not Hidden Then
                    HideGame
                    ShowGame
                    Hidden = True
                End If
            End If
            
            RenderText Font_Default, Map.name, DrawMapNameX, DrawMapNameY, DrawMapNameColor
            If InMapEditor And frmEditor_Map.optEvent.value = True Then DrawEvents
            If InMapEditor Then Call DrawMapAttributes
            
            If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
            If FlashTimer > GetTickCount Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, -1
        Direct3D_Device.EndScene
        
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If

    ' Error handler
    Exit Sub
    
ErrorHandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If Options.Debug = 1 Then
            HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Unrecoverable DX8 error."
        DestroyGame
    End If
End Sub

Sub HandleDeviceLost()
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures True
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
    
   LoadTextures
   
End Sub

Private Function DirectX_ReInit() As Boolean

    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 800 ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 600 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hwnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ConvertMapX = x - (TileView.Left * PIC_X) - Camera.Left
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ConvertMapX", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ConvertMapY = y - (TileView.Top * PIC_Y) - Camera.Top
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ConvertMapY", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    InViewPort = False

    If x < TileView.Left Then Exit Function
    If y < TileView.Top Then Exit Function
    If x > TileView.Right Then Exit Function
    If y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "InViewPort", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsValidMapPoint = False

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > Map.MaxX Then Exit Function
    If y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsValidMapPoint", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim x As Long
Dim y As Long
Dim I As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(x, y).layer(I).Tileset > 0 And Map.Tile(x, y).layer(I).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(x, y).layer(I).Tileset) = True
                End If
            Next
        Next
    Next
    
    For I = 1 To NumTileSets
        If tilesetInUse(I) Then
        
        Else
            ' unload tileset
            'Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            'Set Tex_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub DrawEvents()
Dim sRECT As RECT
Dim Width As Long, Height As Long, I As Long, x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For I = 1 To Map.EventCount
        If Map.Events(I).pageCount <= 0 Then
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(x), ConvertMapY(y), sRECT.Left, sRECT.Right, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        x = Map.Events(I).x * 32
        y = Map.Events(I).y * 32
        x = ConvertMapX(x)
        y = ConvertMapY(y)
        
        If I > Map.EventCount Then Exit Sub
        If 1 > Map.Events(I).pageCount Then Exit Sub
        Select Case Map.Events(I).Pages(1).GraphicType
            Case 0
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Case 1
                If Map.Events(I).Pages(1).Graphic > 0 And Map.Events(I).Pages(1).Graphic <= NumCharacters Then
                    
                    sRECT.Top = (Map.Events(I).Pages(1).GraphicY * (Tex_Character(Map.Events(I).Pages(1).Graphic).Height / 4))
                    
                    If VXFRAME = False Then
                        sRECT.Left = (Map.Events(I).Pages(1).GraphicX * (Tex_Character(Map.Events(I).Pages(1).Graphic).Width / 4))
                    Else
                        sRECT.Left = (Map.Events(I).Pages(1).GraphicX * (Tex_Character(Map.Events(I).Pages(1).Graphic).Width / 3))
                    End If
                    
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Character(Map.Events(I).Pages(1).Graphic), x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            Case 2
                If Map.Events(I).Pages(1).Graphic > 0 And Map.Events(I).Pages(1).Graphic < NumTileSets Then
                    sRECT.Top = Map.Events(I).Pages(1).GraphicY * 32
                    sRECT.Left = Map.Events(I).Pages(1).GraphicX * 32
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Tileset(Map.Events(I).Pages(1).Graphic), x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        End Select
nextevent:
    Next
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawEvents", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorEvent_DrawGraphic()
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.value > 0 And frmEditor_Events.scrlGraphic.value <= NumCharacters Then
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.value
                        sRECT.Right = sRECT.Left + (Tex_Character(frmEditor_Events.scrlGraphic.value).Width - sRECT.Left)
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Character(frmEditor_Events.scrlGraphic.value).Width
                    End If
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.hScrlGraphicSel.value
                        sRECT.Bottom = sRECT.Top + (Tex_Character(frmEditor_Events.scrlGraphic.value).Height - sRECT.Top)
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.value).Height
                    End If
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    With destRect
                        .x1 = dRect.Left
                        .x2 = dRect.Right
                        .y1 = dRect.Top
                        .y2 = dRect.Bottom
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.value), sRECT, dRect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            If VXFRAME = False Then
                                .x1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.value).Width / 4)) - sRECT.Left
                                .x2 = (Tex_Character(frmEditor_Events.scrlGraphic.value).Width / 4) + .x1
                            Else
                                .x1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.value).Width / 3)) - sRECT.Left
                                .x2 = (Tex_Character(frmEditor_Events.scrlGraphic.value).Width / 3) + .x1
                            End If
                            .y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.value).Height / 4)) - sRECT.Top
                            .y2 = (Tex_Character(frmEditor_Events.scrlGraphic.value).Height / 4) + .y1
                        End With

                    Else
                        With destRect
                            .x1 = (GraphicSelX * 32) - sRECT.Left
                            .x2 = ((GraphicSelX2 - GraphicSelX) * 32) + .x1
                            .y1 = (GraphicSelY * 32) - sRECT.Top
                            .y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .y1
                        End With
                    End If
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .x1 = dRect.Left
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y1 = dRect.Top
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .x1 = 0
                        .y1 = 0
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hwnd, ByVal (0)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.value > 0 And frmEditor_Events.scrlGraphic.value <= NumTileSets Then
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.value
                        sRECT.Right = sRECT.Left + 800
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.value).Width
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.value = 0
                    End If
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.vScrlGraphicSel.value
                        sRECT.Bottom = sRECT.Top + 512
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.value).Height
                        frmEditor_Events.vScrlGraphicSel.value = 0
                    End If
                    
                    If sRECT.Left = -1 Then sRECT.Left = 0
                    If sRECT.Top = -1 Then sRECT.Top = 0
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.value), sRECT, dRect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .x1 = (GraphicSelX * 32) - sRECT.Left
                            .x2 = PIC_X + .x1
                            .y1 = (GraphicSelY * 32) - sRECT.Top
                            .y2 = PIC_Y + .y1
                        End With

                    Else
                        With destRect
                            .x1 = (GraphicSelX * 32) - sRECT.Left
                            .x2 = ((GraphicSelX2 - GraphicSelX) * 32) + .x1
                            .y1 = (GraphicSelY * 32) - sRECT.Top
                            .y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .y1
                        End With
                    End If
                    
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .x1 = dRect.Left
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y1 = dRect.Top
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .x1 = 0
                        .y1 = 0
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hwnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    If VXFRAME = False Then
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                        sRECT.Right = sRECT.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    Else
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 3)
                        sRECT.Right = sRECT.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 3)
                    End If
                    sRECT.Bottom = sRECT.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    With dRect
                        dRect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                        dRect.Bottom = dRect.Top + (sRECT.Bottom - sRECT.Top)
                        dRect.Left = (121 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                        dRect.Right = dRect.Left + (sRECT.Right - sRECT.Left)
                    End With
                    With destRect
                        .x1 = dRect.Left
                        .x2 = dRect.Right
                        .y1 = dRect.Top
                        .y2 = dRect.Bottom
                    End With
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.value), sRECT, dRect
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + 32
                        sRECT.Right = sRECT.Left + 32
                        With dRect
                            dRect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRECT.Bottom - sRECT.Top)
                            dRect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRect.Right = dRect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRect
                            .x1 = dRect.Left
                            .x2 = dRect.Right
                            .y1 = dRect.Top
                            .y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.value), sRECT, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                    Else
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRECT.Right = sRECT.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With dRect
                            dRect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRECT.Bottom - sRECT.Top)
                            dRect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRect.Right = dRect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRect
                            .x1 = dRect.Left
                            .x2 = dRect.Right
                            .y1 = dRect.Top
                            .y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.value), sRECT, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                    End If
                End If
        End Select
    End If
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEvent(id As Long)
    Dim x As Long, y As Long, Width As Long, Height As Long, sRECT As RECT, dRect As RECT, anim As Long, spritetop As Long
    If Map.MapEvents(id).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(id).GraphicType
        Case 0
            Exit Sub
            
        Case 1
            If Map.MapEvents(id).GraphicNum <= 0 Or Map.MapEvents(id).GraphicNum > NumCharacters Then Exit Sub
            If VXFRAME = False Then
                Width = Tex_Character(Map.MapEvents(id).GraphicNum).Width / 4
            Else
                Width = Tex_Character(Map.MapEvents(id).GraphicNum).Width / 3
            End If
            Height = Tex_Character(Map.MapEvents(id).GraphicNum).Height / 4
            ' Reset frame
            If VXFRAME = False Then
                If Map.MapEvents(id).Step = 3 Then
                    anim = 0
                ElseIf Map.MapEvents(id).Step = 1 Then
                    anim = 2
                End If
            Else
                
            End If
            
            Select Case Map.MapEvents(id).Dir
                Case DIR_UP
                    If (Map.MapEvents(id).yOffset > 8) Then anim = Map.MapEvents(id).Step
                Case DIR_DOWN
                    If (Map.MapEvents(id).yOffset < -8) Then anim = Map.MapEvents(id).Step
                Case DIR_LEFT
                    If (Map.MapEvents(id).xOffset > 8) Then anim = Map.MapEvents(id).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(id).xOffset < -8) Then anim = Map.MapEvents(id).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(id).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
            
            If Map.MapEvents(id).WalkAnim = 1 Then anim = 0
            
            If Map.MapEvents(id).Moving = 0 Then anim = Map.MapEvents(id).GraphicX
            
            With sRECT
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            x = Map.MapEvents(id).x * PIC_X + Map.MapEvents(id).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                y = Map.MapEvents(id).y * PIC_Y + Map.MapEvents(id).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                y = Map.MapEvents(id).y * PIC_Y + Map.MapEvents(id).yOffset
            End If
        
            ' render the actual sprite
            Call DrawSprite(Map.MapEvents(id).GraphicNum, x, y, sRECT)
            
            
        Case 2
            If Map.MapEvents(id).GraphicNum < 1 Or Map.MapEvents(id).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(id).GraphicY2 > 0 Or Map.MapEvents(id).GraphicX2 > 0 Then
                With sRECT
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) * 32)
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(id).GraphicX2 - Map.MapEvents(id).GraphicX) * 32)
                End With
            Else
                With sRECT
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            x = Map.MapEvents(id).x * 32
            y = Map.MapEvents(id).y * 32
            
            x = x - ((sRECT.Right - sRECT.Left) / 2)
            y = y - (sRECT.Bottom - sRECT.Top) + 32
            
            
            If Map.MapEvents(id).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).x * 32), ConvertMapY((Map.MapEvents(id).y - ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) - 1)) * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).x * 32), ConvertMapY(Map.MapEvents(id).y * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
    End Select
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(x As Single, y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.x = x
    Create_TLVertex.y = y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    'Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
' round it
Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
' if it rounded down, force it up
If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
End Sub

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.Visible Then
        If frmMenu.picCharacter.Visible Then NewCharacterDrawSprite
    End If
    
    If frmEditor_Animation.Visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
    End If
    
    If frmEditor_Map.Visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
        If frmEditor_Map.fraMapKey.Visible Then EditorMap_DrawKey
    End If
    
    If frmEditor_NPC.Visible Then
        EditorNpc_DrawSprite
    End If
    
    If frmEditor_Resource.Visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
    End If
    
    If frmEditor_Events.Visible Then
        EditorEvent_DrawGraphic
    End If
End Sub
Public Sub DrawGUI()
Dim I As Long, x As Long, y As Long
Dim Width As Long, Height As Long

    ' render shadow
    RenderTexture Tex_GUI(23), 0, 0, 0, 0, 800, 64, 1, 64
    RenderTexture Tex_GUI(22), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    ' render chatbox
        If Not inChat Then
            If chatOn Then
                Width = 412
                Height = 145
                RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, Width, Height, Width, Height
                RenderText Font_Default, RenderChatText & chatShowLine, GUIWindow(GUI_CHAT).x + 38, GUIWindow(GUI_CHAT).y + 126, White
                ' draw buttons
                For I = 34 To 35
                    ' set co-ordinate
                    x = GUIWindow(GUI_CHAT).x + Buttons(I).x
                    y = GUIWindow(GUI_CHAT).y + Buttons(I).y
                    Width = Buttons(I).Width
                    Height = Buttons(I).Height
                    ' check for state
                    If Buttons(I).state = 2 Then
                        ' we're clicked boyo
                        'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
                    ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
                        ' we're hoverin'
                        'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
                        ' play sound if needed
                        If Not lastButtonSound = I Then
                            PlaySound Sound_ButtonHover, -1, -1
                            lastButtonSound = I
                        End If
                    Else
                        ' we're normal
                        'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
                        ' reset sound if needed
                        If lastButtonSound = I Then lastButtonSound = 0
                    End If
                Next
            Else
                RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y + 123, 0, 123, 412, 22, 412, 22
            End If
            RenderChatTextBuffer
        Else
            If GUIWindow(GUI_CURRENCY).Visible Then DrawCurrency
            If GUIWindow(GUI_EVENTCHAT).Visible Then DrawEventChat
            If GUIWindow(GUI_QUESTDIALOGUE).Visible Then DrawQuestDialogue
        End If
    
    'DrawGUIBars
    If GUIWindow(GUI_BARS).Visible Then
        If OldGuiBars Then DrawGUIBars Else DrawNewGUIBars
    End If
    
    ' render menu
    If GUIWindow(GUI_MENU).Visible And Options.Buttons = 1 Then DrawMenu
    
    ' render hotbar
    If GUIWindow(GUI_HOTBAR).Visible Then DrawHotbar
    
    ' render menus
    If GUIWindow(GUI_BOOK).Visible Then DrawBook
    If GUIWindow(GUI_INVENTORY).Visible Then DrawInventory
    If GUIWindow(GUI_SPELLS).Visible Then DrawSkills
    If GUIWindow(GUI_CHARACTER).Visible Then DrawCharacter
    If GUIWindow(GUI_OPTIONS).Visible Then DrawOptions
    If GUIWindow(GUI_PARTY).Visible Then DrawParty
    If GUIWindow(GUI_SHOP).Visible Then DrawShop
    If GUIWindow(GUI_BANK).Visible Then DrawBank
    If GUIWindow(GUI_TRADE).Visible Then DrawTrade
    If GUIWindow(GUI_DIALOGUE).Visible Then DrawDialogue
    If GUIWindow(GUI_GUILD).Visible Then DrawGuildMenu
    If GUIWindow(GUI_QUESTLOG).Visible Then DrawQuestLog
    If GUIWindow(GUI_COMBAT).Visible Then DrawCombat
    If GUIWindow(GUI_FRIENDS).Visible Then DrawBuddies
    If GUIWindow(GUI_FRIENDREQUEST).Visible Then DrawFriendRequest
    If GUIWindow(GUI_PLAYERINFO).Visible Then DrawPlayerInfo
    If Scroll_Draw Then DrawScrollEditor
    
    ' Drag and drop
    DrawDragItem
    DrawDragSpell
    
    ' Descriptions
    DrawInventoryItemDesc
    DrawCharacterItemDesc
    DrawPlayerSpellDesc
    DrawBankItemDesc
    DrawTradeItemDesc
End Sub

Public Sub DrawMainMenu()
Dim I As Long, x As Long, y As Long
Dim Width As Long, Height As Long

    If frmMain.Visible = False Then Exit Sub
    
    ' render mainmenu
    Width = 800
    Height = 600
    RenderTexture Tex_GUI(28), 0, 0, 0, 0, Width, Height, Width, Height
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(x, y).layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .x = autoInner(1).x
                .y = autoInner(1).y
            Case "b"
                .x = autoInner(2).x
                .y = autoInner(2).y
            Case "c"
                .x = autoInner(3).x
                .y = autoInner(3).y
            Case "d"
                .x = autoInner(4).x
                .y = autoInner(4).y
            Case "e"
                .x = autoNW(1).x
                .y = autoNW(1).y
            Case "f"
                .x = autoNW(2).x
                .y = autoNW(2).y
            Case "g"
                .x = autoNW(3).x
                .y = autoNW(3).y
            Case "h"
                .x = autoNW(4).x
                .y = autoNW(4).y
            Case "i"
                .x = autoNE(1).x
                .y = autoNE(1).y
            Case "j"
                .x = autoNE(2).x
                .y = autoNE(2).y
            Case "k"
                .x = autoNE(3).x
                .y = autoNE(3).y
            Case "l"
                .x = autoNE(4).x
                .y = autoNE(4).y
            Case "m"
                .x = autoSW(1).x
                .y = autoSW(1).y
            Case "n"
                .x = autoSW(2).x
                .y = autoSW(2).y
            Case "o"
                .x = autoSW(3).x
                .y = autoSW(3).y
            Case "p"
                .x = autoSW(4).x
                .y = autoSW(4).y
            Case "q"
                .x = autoSE(1).x
                .y = autoSE(1).y
            Case "r"
                .x = autoSE(2).x
                .y = autoSE(2).y
            Case "s"
                .x = autoSE(3).x
                .y = autoSE(3).y
            Case "t"
                .x = autoSE(4).x
                .y = autoSE(4).y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim x As Long, y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).x = 32
    autoInner(1).y = 0
    
    ' NE - b
    autoInner(2).x = 48
    autoInner(2).y = 0
    
    ' SW - c
    autoInner(3).x = 32
    autoInner(3).y = 16
    
    ' SE - d
    autoInner(4).x = 48
    autoInner(4).y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).x = 0
    autoNW(1).y = 32
    
    ' NE - f
    autoNW(2).x = 16
    autoNW(2).y = 32
    
    ' SW - g
    autoNW(3).x = 0
    autoNW(3).y = 48
    
    ' SE - h
    autoNW(4).x = 16
    autoNW(4).y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).x = 32
    autoNE(1).y = 32
    
    ' NE - g
    autoNE(2).x = 48
    autoNE(2).y = 32
    
    ' SW - k
    autoNE(3).x = 32
    autoNE(3).y = 48
    
    ' SE - l
    autoNE(4).x = 48
    autoNE(4).y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).x = 0
    autoSW(1).y = 64
    
    ' NE - n
    autoSW(2).x = 16
    autoSW(2).y = 64
    
    ' SW - o
    autoSW(3).x = 0
    autoSW(3).y = 80
    
    ' SE - p
    autoSW(4).x = 16
    autoSW(4).y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).x = 32
    autoSE(1).y = 64
    
    ' NE - r
    autoSE(2).x = 48
    autoSE(2).y = 64
    
    ' SW - s
    autoSE(3).x = 32
    autoSE(3).y = 80
    
    ' SE - t
    autoSE(4).x = 48
    autoSE(4).y = 80
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile x, y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState x, y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If x < 0 Or x > Map.MaxX Or y < 0 Or y > Map.MaxY Then Exit Sub

    With Map.Tile(x, y)
        ' check if the tile can be rendered
        If .layer(layerNum).Tileset <= 0 Or .layer(layerNum).Tileset > NumTileSets Then
            Autotile(x, y).layer(layerNum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it's a key - hide mask if key is closed
        If layerNum = MapLayer.Mask Then
            If .Type = TILE_TYPE_KEY Then
                If TempTile(x, y).DoorOpen = NO Then
                    Autotile(x, y).layer(layerNum).RenderState = RENDER_STATE_NONE
                    Exit Sub
                Else
                    Autotile(x, y).layer(layerNum).RenderState = RENDER_STATE_NORMAL
                    Exit Sub
                End If
            End If
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(x, y).layer(layerNum).RenderState = RENDER_STATE_NORMAL
        Else
            Autotile(x, y).layer(layerNum).RenderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(x, y).layer(layerNum).srcX(quarterNum) = (Map.Tile(x, y).layer(layerNum).x * 32) + Autotile(x, y).layer(layerNum).QuarterTile(quarterNum).x
                Autotile(x, y).layer(layerNum).srcY(quarterNum) = (Map.Tile(x, y).layer(layerNum).y * 32) + Autotile(x, y).layer(layerNum).QuarterTile(quarterNum).y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(x, y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(x, y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, x, y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, x, y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, x, y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If x2 < 0 Or x2 > Map.MaxX Or y2 < 0 Or y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(x2, y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(x2, y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(x1, y1).layer(layerNum).Tileset <> Map.Tile(x2, y2).layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(x1, y1).layer(layerNum).x <> Map.Tile(x2, y2).layer(layerNum).x Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(x1, y1).layer(layerNum).y <> Map.Tile(x2, y2).layer(layerNum).y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal x As Long, ByVal y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(x, y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    RenderTexture Tex_Tileset(Map.Tile(x, y).layer(layerNum).Tileset), destX, destY, Autotile(x, y).layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, -1
End Sub

Public Sub DrawItem(ByVal itemNum As Long)
Dim PicNum As Integer, dontRender As Boolean, I As Long, tmpIndex As Long
    
    PicNum = Item(MapItem(itemNum).num).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

     ' if it's not us then don't render
    If MapItem(itemNum).PlayerName <> vbNullString Then
        If Trim$(MapItem(itemNum).PlayerName) <> Trim$(GetPlayerName(MyIndex)) Then
            dontRender = True
        End If
        ' make sure it's not a party drop
        If Party.Leader > 0 Then
            For I = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(I)
                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemNum).PlayerName) Then
                        dontRender = False
                    End If
                End If
            Next
        End If
    End If
    
    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemNum).x * PIC_X), ConvertMapY(MapItem(itemNum).y * PIC_Y), 0, 0, 32, 32, 32, 32
    End If
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemNum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not itemNum > 0 Then Exit Sub
    
    PicNum = Item(itemNum).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    'EngineRenderRectangle Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, spellnum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    spellnum = PlayerSpells(DragSpell)
    If Not spellnum > 0 Then Exit Sub
    
    PicNum = Spell(spellnum).Icon

    If PicNum < 1 Or PicNum > NumSpellIcons Then Exit Sub

    'EngineRenderRectangle Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_SpellIcon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawHotbar()
Dim I As Long, x As Long, y As Long, t As Long, sS As String
Dim Width As Long, Height As Long, Color As Long

    For I = 1 To MAX_HOTBAR
        ' draw the box
        x = GUIWindow(GUI_HOTBAR).x + ((I - 1) * (5 + 36))
        y = GUIWindow(GUI_HOTBAR).y
        Width = 36
        Height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        RenderTexture Tex_GUI(2), x, y, 0, 0, Width, Height, Width, Height
        ' draw the icon
        Select Case Hotbar(I).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(I).Slot).name) > 0 Then
                    If Item(Hotbar(I).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_Item(Item(Hotbar(I).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(I).Slot).name) > 0 Then
                    If Spell(Hotbar(I).Slot).Icon > 0 Then
                        ' render normal icon
                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_SpellIcon(Spell(Hotbar(I).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32
                        ' we got the spell?
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t) > 0 Then
                                If PlayerSpells(t) = Hotbar(I).Slot Then
                                    If SpellCD(t) > 0 Then
                                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                                        RenderTexture Tex_SpellIcon(Spell(Hotbar(I).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
        ' draw the numbers
        sS = str(I)
        If I = 10 Then sS = "0"
        If I = 11 Then sS = " -"
        If I = 12 Then sS = " ="
        RenderText Font_Default, sS, x + 4, y + 20, White
    Next
End Sub
Public Sub DrawInventory()
Dim I As Long, x As Long, y As Long, itemNum As Long, ItemPic As Long
Dim Amount As String
Dim colour As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, Width, Height, Width, Height
    
    For I = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, I)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    If TradeYourOffer(x).num = I Then
                        GoTo NextLoop
                    End If
                Next
            End If
            
            ' exit out if dragging
            If DragInvSlotNum = I Then GoTo NextLoop

            If ItemPic > 0 And ItemPic <= numitems Then
                Top = GUIWindow(GUI_INVENTORY).y + InvTop - 2 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                Left = GUIWindow(GUI_INVENTORY).x + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                If PlayerInv(I).Selected = 0 Then
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                Else
                    'LEFTOFF - Needs all items to have an _s copy to work
                    RenderTexture Tex_Item_S(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    'RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
                ' If item is a stack - draw the amount you have
                If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                    y = Top + 21
                    x = Left - 4
                    Amount = CStr(GetPlayerInvItemValue(MyIndex, I))
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawInventoryItemDesc()
Dim invSlot As Long, isSB As Boolean
        
    If Not GUIWindow(GUI_INVENTORY).Visible Then Exit Sub
    If DragInvSlotNum > 0 Then Exit Sub
    
    invSlot = IsInvItem(GlobalX, GlobalY)
    If invSlot > 0 Then
        If GetPlayerInvItemNum(MyIndex, invSlot) > 0 Then
            'If Item(GetPlayerInvItemNum(MyIndex, invSlot)).BindType > 0 And PlayerInv(invSlot).bound > 0 Then isSB = True
            DrawItemDesc GetPlayerInvItemNum(MyIndex, invSlot), GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).y, isSB
            ' value
            If InShop > 0 Then
                DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_DESCRIPTION).Height + 10
            End If
        End If
    End If
End Sub

Public Sub DrawShopItemDesc()
Dim shopSlot As Long
        
    If Not GUIWindow(GUI_SHOP).Visible Then Exit Sub
    
    shopSlot = IsShopItem(GlobalX, GlobalY)
    If shopSlot > 0 Then
        If Shop(InShop).TradeItem(shopSlot).Item > 0 Then
            DrawItemDesc Shop(InShop).TradeItem(shopSlot).Item, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).y
            DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).y + GUIWindow(GUI_DESCRIPTION).Height + 10
        End If
    End If
End Sub

Public Sub DrawCharacterItemDesc()
Dim eqSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_CHARACTER).Visible Then Exit Sub
    
    eqSlot = IsEqItem(GlobalX, GlobalY)
    If eqSlot > 0 Then
        If GetPlayerEquipment(MyIndex, eqSlot) > 0 Then
            If Item(GetPlayerEquipment(MyIndex, eqSlot)).BindType > 0 Then isSB = True
            DrawItemDesc GetPlayerEquipment(MyIndex, eqSlot), GUIWindow(GUI_CHARACTER).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_CHARACTER).y, isSB
        End If
    End If
End Sub

Public Sub DrawItemCost(ByVal isShop As Boolean, ByVal slotNum As Long, ByVal x As Long, ByVal y As Long)
Dim CostItem As Long, CostValue As Long, itemNum As Long, sString As String, Width As Long, Height As Long

    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    ' draw the window
    Width = 190
    Height = 36

    RenderTexture Tex_GUI(24), x, y, 0, 0, Width, Height, Width, Height
    
    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        itemNum = GetPlayerInvItemNum(MyIndex, slotNum)
        If itemNum = 0 Then Exit Sub
        CostItem = 1
        CostValue = (Item(itemNum).Price / 100) * Shop(InShop).BuyRate
        sString = "The shop will buy for"
    Else
        itemNum = Shop(InShop).TradeItem(slotNum).Item
        If itemNum = 0 Then Exit Sub
        CostItem = Shop(InShop).TradeItem(slotNum).CostItem
        CostValue = Shop(InShop).TradeItem(slotNum).CostValue
        sString = "The shop will sell for"
    End If
    
    'EngineRenderRectangle Tex_Item(Item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(Item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32
    
    RenderText Font_Default, sString, x + 4, y + 3, DarkGrey
    
    RenderText Font_Default, CostValue & " " & Trim$(Item(CostItem).name), x + 4, y + 18, White
End Sub

Public Sub DrawItemDesc(ByVal itemNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal soulBound As Boolean = False)
Dim colour As Long, descString As String, theName As String, className As String, levelTxt As String, sInfo() As String, I As Long, Width As Long, Height As Long

    ' get out
    If itemNum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Item(itemNum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), x, y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Item(itemNum).Pic > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 64, 64
        RenderTexture Tex_Item(Item(itemNum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not Trim$(Item(itemNum).Desc) = vbNullString Then
        RenderText Font_Default, WordWrap(Trim$(Item(itemNum).Desc), Width - 10), x + 10, y + 128, White
    End If
    ' work out name colour
    Select Case Item(itemNum).Rarity
        Case 0 ' white
            colour = White
        Case 1 ' green
            colour = Green
        Case 2 ' blue
            colour = Blue
        Case 3 ' maroon
            colour = Red
        Case 4 ' purple
            colour = Pink
        Case 5 ' orange
            colour = Brown
    End Select
    
    If Not soulBound Then
        theName = Trim$(Item(itemNum).name)
    Else
        theName = "(SB) " & Trim$(Item(itemNum).name)
    End If
    
    ' render name
    RenderText Font_Default, theName, x + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), y + 6, colour
    
    ' class req
    If Item(itemNum).ClassReq > 0 Then
        className = Trim$(Class(Item(itemNum).ClassReq).name)
        ' do we match it?
        If GetPlayerClass(MyIndex) = Item(itemNum).ClassReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        className = "No class req."
        colour = Green
    End If
    RenderText Font_Default, className, x + 48 - (EngineGetTextWidth(Font_Default, className) \ 2), y + 92, colour
    
    ' level
    If Item(itemNum).LevelReq > 0 Then
        levelTxt = "Level " & Item(itemNum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(itemNum).LevelReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        levelTxt = "No level req."
        colour = Green
    End If
    RenderText Font_Default, levelTxt, x + 48 - (EngineGetTextWidth(Font_Default, levelTxt) \ 2), y + 107, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE
            sInfo(I) = "No type"
        Case ITEM_TYPE_WEAPON
            sInfo(I) = "Weapon"
        Case ITEM_TYPE_ARMOR
            sInfo(I) = "Armour"
        Case ITEM_TYPE_HELMET
            sInfo(I) = "Helmet"
        Case ITEM_TYPE_SHIELD
            sInfo(I) = "Shield"
        Case ITEM_TYPE_CONSUME
            sInfo(I) = "Consume"
        Case ITEM_TYPE_KEY
            sInfo(I) = "Key"
        Case ITEM_TYPE_CURRENCY
            sInfo(I) = "Currency"
        Case ITEM_TYPE_SPELL
            sInfo(I) = "Spell"
    End Select
    
    ' more info
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
            ' binding
            If Item(itemNum).BindType = 1 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Pickup"
            ElseIf Item(itemNum).BindType = 2 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Equip"
            End If
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price & "g"
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' binding
            If Item(itemNum).BindType = 1 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Pickup"
            ElseIf Item(itemNum).BindType = 2 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Equip"
            End If
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price & "g"
            ' damage/defence
            If Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Damage: " & Item(itemNum).Data2
                ' speed
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Speed: " & (Item(itemNum).speed / 1000) & "s"
            End If
            ' stat bonuses
            If Item(itemNum).Add_Stat(Stats.Strength) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Strength) & " Str"
            End If
            If Item(itemNum).Add_Stat(Stats.Endurance) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(itemNum).Add_Stat(Stats.Intelligence) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(itemNum).Add_Stat(Stats.Agility) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(itemNum).Add_Stat(Stats.Willpower) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price & "g"
            If Item(itemNum).CastSpell > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Casts Spell"
            End If
            If Item(itemNum).AddHP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddHP & " HP"
            End If
            If Item(itemNum).AddMP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddMP & " SP"
            End If
            If Item(itemNum).AddEXP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddEXP & " EXP"
            End If
        Case ITEM_TYPE_SPELL
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price & "g"
    End Select
    
    ' go through and render all this shit
    y = y + 12
    For I = 1 To UBound(sInfo)
        y = y + 12
        RenderText Font_Default, sInfo(I), x + 141 - (EngineGetTextWidth(Font_Default, sInfo(I)) \ 2), y, White
    Next
End Sub
Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
        
    If Not GUIWindow(GUI_SPELLS).Visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If PlayerSpells(spellSlot) > 0 Then
            DrawSpellDesc PlayerSpells(spellSlot), GUIWindow(GUI_SPELLS).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_SPELLS).y, spellSlot
        End If
    End If
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal spellSlot As Long = 0)
Dim colour As Long, theName As String, sUse As String, sInfo() As String, I As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, Height As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If spellnum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Spell(spellnum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(29), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), x, y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Spell(spellnum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        RenderTexture Tex_SpellIcon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not Trim$(Spell(spellnum).Desc) = vbNullString Then
        RenderText Font_Default, WordWrap(Trim$(Spell(spellnum).Desc), Width - 10), x + 10, y + 128, White
    End If
    
    ' render name
    colour = White
    theName = Trim$(Spell(spellnum).name)
    RenderText Font_Default, theName, x + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), y + 6, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP
            sInfo(I) = "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            sInfo(I) = "Damage SP"
        Case SPELL_TYPE_HEALHP
            sInfo(I) = "Heal HP"
        Case SPELL_TYPE_HEALMP
            sInfo(I) = "Heal SP"
        Case SPELL_TYPE_WARP
            sInfo(I) = "Warp"
    End Select
    
    ' more info
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Vital: " & Spell(spellnum).Vital
            
            ' mp cost
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).AoE > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "AoE: " & Spell(spellnum).AoE
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
            End If
    End Select
    
    ' go through and render all this shit
    y = y + 12
    For I = 1 To UBound(sInfo)
        y = y + 12
        RenderText Font_Default, sInfo(I), x + 141 - (EngineGetTextWidth(Font_Default, sInfo(I)) \ 2), y, White
    Next
End Sub

Public Sub DrawSkills()
Dim I As Long, x As Long, y As Long, spellnum As Long, spellpic As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, Width, Height, Width, Height
    
    ' render skills
    For I = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(I)

        ' make sure not dragging it
        If DragSpell = I Then GoTo NextLoop
        
        ' actually render
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellpic = Spell(spellnum).Icon

            If spellpic > 0 And spellpic <= NumSpellIcons Then
                Top = GUIWindow(GUI_SPELLS).y + SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).x + SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
                If SpellCD(I) > 0 Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawEquipment()
Dim x As Long, y As Long, I As Long
Dim itemNum As Long, ItemPic As DX8TextureRec

    For I = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, I)

        ' get the item sprite
        If itemNum > 0 Then
            ItemPic = Tex_Item(Item(itemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(8 + I)
        End If
        
        y = GUIWindow(GUI_CHARACTER).y + EqTop
        x = GUIWindow(GUI_CHARACTER).x + EqLeft + ((EqOffsetX + 32) * (((I - 1) Mod EqColumns)))

        'EngineRenderRectangle itempic, x, y, 0, 0, 32, 32, 32, 32, 32, 32
        RenderTexture ItemPic, x, y, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawCharacter()
Dim x As Long, y As Long, I As Long, dX As Long, dY As Long, tmpString As String, buttonnum As Long
Dim Width As Long, Height As Long
    
    x = GUIWindow(GUI_CHARACTER).x
    y = GUIWindow(GUI_CHARACTER).y
    
    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(5), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(6), x, y, 0, 0, Width, Height, Width, Height
    
    ' render name
    tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    RenderText Font_Default, tmpString, x + 7 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), y + 9, White
    
    ' render stats
    dX = x + 20
    dY = y + 145
    RenderText Font_Default, "Str: " & GetPlayerStat(MyIndex, Strength), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "End: " & GetPlayerStat(MyIndex, Endurance), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Int: " & GetPlayerStat(MyIndex, Intelligence), dX, dY, White
    dY = y + 145
    dX = dX + 80
    RenderText Font_Default, "Agi: " & GetPlayerStat(MyIndex, Agility), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Will: " & GetPlayerStat(MyIndex, Willpower), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Pnts: " & GetPlayerPOINTS(MyIndex), dX, dY, White
    
    ' draw the face
    If GetPlayerSprite(MyIndex) > 0 And GetPlayerSprite(MyIndex) <= NumFaces Then
        'EngineRenderRectangle Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
        RenderTexture Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96
    End If
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        ' draw the buttons
        For buttonnum = 16 To 20
            x = GUIWindow(GUI_CHARACTER).x + Buttons(buttonnum).x
            y = GUIWindow(GUI_CHARACTER).y + Buttons(buttonnum).y
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                Width = Buttons(buttonnum).Width
                Height = Buttons(buttonnum).Height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
    
    ' draw the equipment
    DrawEquipment
End Sub

Public Sub DrawOptions()
Dim I As Long, x As Long, y As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(24), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(21), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, Width, Height, Width, Height
    RenderText Font_Default, "MiniMap", GUIWindow(GUI_OPTIONS).x + 15, GUIWindow(GUI_OPTIONS).y + 118, White
    RenderText Font_Default, "Buttons", GUIWindow(GUI_OPTIONS).x + 19, GUIWindow(GUI_OPTIONS).y + 143, White
    
    ' draw buttons
    For I = 26 To 33
        ' set co-ordinate
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    
    ' draw buttons
    For I = 59 To 62
        ' set co-ordinate
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawParty()
Dim I As Long, x As Long, y As Long, Width As Long, playerNum As Long, theName As String
Dim Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(7), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, Width, Height, Width, Height
    
    ' draw the bars
    If Party.Leader > 0 Then ' make sure we're in a party
        ' draw leader
        playerNum = Party.Leader
        ' name
        theName = Trim$(GetPlayerName(playerNum))
        ' draw name
        y = GUIWindow(GUI_PARTY).y + 12
        x = GUIWindow(GUI_PARTY).x + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
        RenderText Font_Default, theName, x, y, White
        ' draw hp
        y = GUIWindow(GUI_PARTY).y + 29
        x = GUIWindow(GUI_PARTY).x + 6
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
        End If
        'EngineRenderRectangle Tex_GUI(13), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(13), x, y, 0, 0, Width, 9, Width, 9
        ' draw mp
        y = GUIWindow(GUI_PARTY).y + 38
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
        End If
        'EngineRenderRectangle Tex_GUI(14), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(14), x, y, 0, 0, Width, 9, Width, 9
        
        ' draw members
        For I = 1 To MAX_PARTY_MEMBERS
            If Party.Member(I) > 0 Then
                If Party.Member(I) <> Party.Leader Then
                    ' cache the index
                    playerNum = Party.Member(I)
                    ' name
                    theName = Trim$(GetPlayerName(playerNum))
                    ' draw name
                    y = GUIWindow(GUI_PARTY).y + 12 + ((I - 1) * 49)
                    x = GUIWindow(GUI_PARTY).x + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
                    RenderText Font_Default, theName, x, y, White
                    ' draw hp
                    y = GUIWindow(GUI_PARTY).y + 29 + ((I - 1) * 49)
                    x = GUIWindow(GUI_PARTY).x + 6
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(13), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(13), x, y, 0, 0, Width, 9, Width, 9
                    ' draw mp
                    y = GUIWindow(GUI_PARTY).y + 38 + ((I - 1) * 49)
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(14), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(14), x, y, 0, 0, Width, 9, Width, 9
                End If
            End If
        Next
    End If
    
    ' draw buttons
    For I = 24 To 25
        ' set co-ordinate
        x = GUIWindow(GUI_PARTY).x + Buttons(I).x
        y = GUIWindow(GUI_PARTY).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub
Public Sub DrawCurrency()
Dim x As Long, y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    x = GUIWindow(GUI_CURRENCY).x
    y = GUIWindow(GUI_CURRENCY).y
    ' render chatbox
    Width = GUIWindow(GUI_CURRENCY).Width
    Height = GUIWindow(GUI_CURRENCY).Height
    RenderTexture Tex_GUI(27), x, y, 0, 0, Width, Height, Width, Height
    Width = EngineGetTextWidth(Font_Default, CurrencyText)
    RenderText Font_Default, CurrencyText, x + 87 + (123 - (Width / 2)), y + 40, White
    RenderText Font_Default, sDialogue & chatShowLine, x + 90, y + 65, White
    
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    x = GUIWindow(GUI_CURRENCY).x + 155
    y = GUIWindow(GUI_CURRENCY).y + 96
    If CurrencyAcceptState = 2 Then
        ' clicked
        RenderText Font_Default, "[Accept]", x, y, Grey
    Else
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            ' hover
            RenderText Font_Default, "[Accept]", x, y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 1 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 1
            End If
        Else
            ' normal
            RenderText Font_Default, "[Accept]", x, y, Green
            ' reset sound if needed
            If lastNpcChatsound = 1 Then lastNpcChatsound = 0
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    x = GUIWindow(GUI_CURRENCY).x + 218
    y = GUIWindow(GUI_CURRENCY).y + 96
    If CurrencyCloseState = 2 Then
        ' clicked
        RenderText Font_Default, "[Close]", x, y, Grey
    Else
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            ' hover
            RenderText Font_Default, "[Close]", x, y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 2 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 2
            End If
        Else
            ' normal
            RenderText Font_Default, "[Close]", x, y, Yellow
            ' reset sound if needed
            If lastNpcChatsound = 2 Then lastNpcChatsound = 0
        End If
    End If
End Sub
Public Sub DrawDialogue()
Dim I As Long, x As Long, y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    x = GUIWindow(GUI_DIALOGUE).x
    y = GUIWindow(GUI_DIALOGUE).y
    
    ' render chatbox
    Width = GUIWindow(GUI_DIALOGUE).Width
    Height = GUIWindow(GUI_DIALOGUE).Height
    RenderTexture Tex_GUI(19), x, y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(Dialogue_TitleCaption, 392), x + 10, y + 10, White
    RenderText Font_Default, WordWrap(Dialogue_TextCaption, 392), x + 10, y + 25, White
    
    If Dialogue_ButtonVisible(1) Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 90
            If Dialogue_ButtonState(1) = 2 Then
                ' clicked
                RenderText Font_Default, "[Accept]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Accept]", x, y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Accept]", x, y, Green
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(2) Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 105
            If Dialogue_ButtonState(2) = 2 Then
                ' clicked
                RenderText Font_Default, "[Okay]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Okay]", x, y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Okay]", x, y, BrightRed
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(3) Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 120
        If Dialogue_ButtonState(3) = 2 Then
            ' clicked
            RenderText Font_Default, "[Close]", x, y, Grey
        Else
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                ' hover
                RenderText Font_Default, "[Close]", x, y, Cyan
                ' play sound if needed
                If Not lastNpcChatsound = 3 Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = 3
                End If
            Else
                ' normal
                RenderText Font_Default, "[Close]", x, y, Yellow
                ' reset sound if needed
                If lastNpcChatsound = 3 Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub


Public Sub DrawScrollEditor()
Dim x As Long, y As Long

        x = 400
        y = 300
        Select Case Scroll_Editor
            Case 1
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "M", x, y, White
                'RenderTexture Tex_GUI(20), x, y, 0, 0, Width, Height, Width, Height
                Exit Sub
            Case 2
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "N", x, y, White
                Exit Sub
            Case 3
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "I", x, y, White
                Exit Sub
            Case 4
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "R", x, y, White
                Exit Sub
            Case 5
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "Q", x, y, White
                Exit Sub
            Case 6
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "Sp", x, y, White
                Exit Sub
            Case 7
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "Ch", x, y, White
                Exit Sub
            Case 8
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "A", x, y, White
                Exit Sub
            Case 9
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "Sh", x, y, White
                Exit Sub
            Case 10
                RenderTexture Tex_GUI(2), x - 14, y - 13, 0, 0, 50, 50, 50, 50
                RenderText Font_Georgia, "Cb", x, y, White
                Exit Sub
            Case Else
                Exit Sub
        End Select
        
End Sub

Public Sub DrawShop()
Dim I As Long, x As Long, y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = GUIWindow(GUI_SHOP).Width
    Height = GUIWindow(GUI_SHOP).Height
    'EngineRenderRectangle Tex_GUI(23), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(20), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, Width, Height, Width, Height
    
    ' render the shop items
    For I = 1 To MAX_TRADES
        itemNum = Shop(InShop).TradeItem(I).Item
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= numitems Then
                
                Top = GUIWindow(GUI_SHOP).y + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                Left = GUIWindow(GUI_SHOP).x + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(I).ItemValue > 1 Then
                    y = GUIWindow(GUI_SHOP).y + Top + 22
                    x = GUIWindow(GUI_SHOP).x + Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(I).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                End If
            End If
        End If
    Next
    
    ' draw buttons
    For I = 23 To 23
        ' set co-ordinate
        x = GUIWindow(GUI_SHOP).x + Buttons(I).x
        y = GUIWindow(GUI_SHOP).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    
    ' draw item descriptions
    DrawShopItemDesc
End Sub

Public Sub DrawMenu()
Dim I As Long, x As Long, y As Long
Dim Width As Long, Height As Long

    ' draw background
    x = GUIWindow(GUI_MENU).x
    y = GUIWindow(GUI_MENU).y
    Width = GUIWindow(GUI_MENU).Width
    Height = GUIWindow(GUI_MENU).Height
    'EngineRenderRectangle Tex_GUI(3), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(3), x, y, 0, 0, Width, Height, Width, Height
    
    ' draw buttons
    For I = 1 To 6
        If Buttons(I).Visible Then
            ' set co-ordinate
            x = GUIWindow(GUI_MENU).x + Buttons(I).x
            y = GUIWindow(GUI_MENU).y + Buttons(I).y
            Width = Buttons(I).Width
            Height = Buttons(I).Height
            ' check for state
            If Buttons(I).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = I
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = I Then lastButtonSound = 0
            End If
        End If
    Next
    
    ' draw quest and guild buttons
    I = 42
    If Buttons(I).Visible Then
        ' set co-ordinate
        x = GUIWindow(GUI_MENU).x + Buttons(I).x
        y = GUIWindow(GUI_MENU).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
        If lastButtonSound = I Then lastButtonSound = 0
        End If
    End If
End Sub


Public Sub DrawBank()
Dim I As Long, x As Long, y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_BANK).Width
    Height = GUIWindow(GUI_BANK).Height
    
    RenderTexture Tex_GUI(26), GUIWindow(GUI_BANK).x, GUIWindow(GUI_BANK).y, 0, 0, Width, Height, Width, Height
    
    ' render the bank items' are you serous? that is it??? maybe... one sec :D :Polol
        For I = 1 To MAX_BANK
            itemNum = GetBankItemNum(I)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                        
                     Top = GUIWindow(GUI_BANK).y + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                     Left = GUIWindow(GUI_BANK).x + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))

                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                       
                    ' If the bank item is in a stack, draw the amount...
                    If GetBankItemValue(I) > 1 Then
                        y = Top + 22
                        x = Left - 4
                        Amount = CStr(GetBankItemValue(I))
                            
                        ' Draw the currency
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                    
                        RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                    End If
                End If
            End If
    Next
                
             DrawBankItemDesc
                            
                        
End Sub
Public Sub DrawBankItemDesc()
Dim bankNum As Long

    If Not GUIWindow(GUI_BANK).Visible Then Exit Sub
        
        bankNum = IsBankItem(GlobalX, GlobalY)
     
        
    If bankNum > 0 Then
        If bankNum > 0 Then
            If GetBankItemNum(bankNum) > 0 Then
                DrawItemDesc GetBankItemNum(bankNum), GUIWindow(GUI_BANK).x + 480, GUIWindow(GUI_BANK).y
           End If
        End If
    End If
            
End Sub

Public Sub DrawTrade()
Dim I As Long, x As Long, y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_TRADE).Width
    Height = GUIWindow(GUI_TRADE).Width
    RenderTexture Tex_GUI(18), GUIWindow(GUI_TRADE).x, GUIWindow(GUI_TRADE).y, 0, 0, Width, Height, Width, Height
        For I = 1 To MAX_INV
            ' render your offer
            itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                    Top = GUIWindow(GUI_TRADE).y + 31 + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).x + 29 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(I).value > 1 Then
                        y = Top + 21
                        x = Left - 4
                            
                        Amount = CStr(TradeYourOffer(I).value)
                            
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                    End If
                End If
            End If
            
            ' draw their offer
            itemNum = TradeTheirOffer(I).num
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                
                    Top = GUIWindow(GUI_TRADE).y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).x + 257 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(I).value > 1 Then
                        y = Top + 21
                        x = Left - 4
                                
                        Amount = CStr(TradeTheirOffer(I).value)
                                
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                    End If
                End If
            End If
        Next
        ' draw buttons
    For I = 40 To 41
        ' set co-ordinate
        x = Buttons(I).x
        y = Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    RenderText Font_Default, "Your worth: " & YourWorth, GUIWindow(GUI_TRADE).x + 21, GUIWindow(GUI_TRADE).y + 299, White
    RenderText Font_Default, "Their worth: " & TheirWorth, GUIWindow(GUI_TRADE).x + 250, GUIWindow(GUI_TRADE).y + 299, White
    RenderText Font_Default, TradeStatus, (GUIWindow(GUI_TRADE).Width / 2) - (EngineGetTextWidth(Font_Default, TradeStatus) / 2), GUIWindow(GUI_TRADE).y + 317, Yellow
    DrawTradeItemDesc
End Sub

Public Sub DrawTradeItemDesc()
Dim tradeNum As Long

    If Not GUIWindow(GUI_TRADE).Visible Then Exit Sub
        
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum > 0 Then
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).num) > 0 Then
            DrawItemDesc GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).num), GUIWindow(GUI_TRADE).x + 480 + 10, GUIWindow(GUI_TRADE).y
        End If
    End If
End Sub

Public Sub DrawGUIBars()
Dim tmpWidth As Long, barWidth As Long, x As Long, y As Long, dX As Long, dY As Long, sString As String
Dim Width As Long, Height As Long

    ' backwindow + empty bars
    x = GUIWindow(GUI_BARS).x
    y = GUIWindow(GUI_BARS).y
    Width = 254
    Height = 75
    'EngineRenderRectangle Tex_GUI(4), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(4), x, y, 0, 0, Width, Height, Width, Height
    
    ' hardcoded for POT textures
    barWidth = 241
    
    ' health bar
    BarWidth_GuiHP = ((GetPlayerVital(MyIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(13), x + 7, y + 9, 0, 0, BarWidth_GuiHP, Tex_GUI(13).Height, BarWidth_GuiHP, Tex_GUI(13).Height
    ' render health
    sString = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    dX = x + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = y + 9
    RenderText Font_Default, sString, dX, dY, White
    
    ' spirit bar
    BarWidth_GuiSP = ((GetPlayerVital(MyIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(14), x + 7, y + 31, 0, 0, BarWidth_GuiSP, Tex_GUI(14).Height, BarWidth_GuiSP, Tex_GUI(14).Height
    ' render spirit
    sString = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    dX = x + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = y + 31
    RenderText Font_Default, sString, dX, dY, White
    
    ' exp bar
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        BarWidth_GuiEXP = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
    Else
        BarWidth_GuiEXP = barWidth
    End If
    RenderTexture Tex_GUI(15), x + 7, y + 53, 0, 0, BarWidth_GuiEXP, Tex_GUI(15).Height, BarWidth_GuiEXP, Tex_GUI(15).Height
    ' render exp
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        sString = GetPlayerExp(MyIndex) & "/" & TNL
    Else
        sString = "Max Level"
    End If
    dX = x + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = y + 53
    RenderText Font_Default, sString, dX, dY, White
End Sub

Public Sub DrawNewGUIBars()
Dim tmpWidth As Long, barWidthR As Long, barWidth As Long, x As Long, y As Long, dX As Long, dY As Long, sString As String
Dim Width As Long, Height As Long, HPBHeight As Long, HPbarWidth As Long, SPBHeight As Long, SPbarWidth As Long, I As Long
' backwindow + empty bars
x = 15 'Xr
y = 15 'Yr
Width = 142
Height = 116
RenderTexture Tex_GUI(42), x, y, 0, 0, Width, Height, Width, Height

' hardcoded for POT textures
HPbarWidth = -72
HPBHeight = -128
' health bar
BarWidth_GuiHP = ((GetPlayerVital(MyIndex, Vitals.HP) * HPbarWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPbarWidth)) / HPbarWidth
BarHeight_GuiHP = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBHeight) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBHeight)) * HPBHeight
If BarHeight_GuiHP <= -64 Then
RenderTexture Tex_GUI(43), x + 25, y + 138, 0, 0, Tex_GUI(43).Width, BarHeight_GuiHP, Tex_GUI(43).Width, BarHeight_GuiHP
End If
If BarHeight_GuiHP > -64 Then
     RenderTexture Tex_GUI(43), x + 154, y + 138, 0, 0, BarWidth_GuiHP - 66, -64, BarWidth_GuiHP - 66, -64
End If
' render health
'sString = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
'dX = X + 130 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
'dY = Y + 39
'RenderText Font_Default, sString, dX, dY, White

' spirit bar
     SPbarWidth = 69
     SPBHeight = -128
     BarWidth_GuiSP = ((GetPlayerVital(MyIndex, Vitals.MP) * SPbarWidth) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPbarWidth)) / SPbarWidth
     BarHeight_GuiSP = ((GetPlayerVital(MyIndex, Vitals.MP) / SPBHeight) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPBHeight)) * SPBHeight
If BarHeight_GuiSP <= -64 Then
RenderTexture Tex_GUI(44), x + 66, y + 132, 0, 0, Tex_GUI(44).Width, BarHeight_GuiSP, Tex_GUI(44).Width, BarHeight_GuiSP
End If

If BarHeight_GuiSP > -64 Then
     RenderTexture Tex_GUI(44), x + 185, y + 132, 0, 0, BarWidth_GuiSP + 60, -50, BarWidth_GuiSP + 60, -50
End If
' render spirit
' sString = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
' dX = X + 130 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
' dY = Y + 61
' RenderText Font_Default, sString, dX, dY, White


' exp bar
barWidth = 72
If GetPlayerLevel(MyIndex) < MAX_LEVELS = True Then
     BarHeight_GuiEXP = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
Else
     BarHeight_GuiEXP = barWidth
End If
RenderTexture Tex_GUI(45), x + 12, y + 39, 0, 0, Tex_GUI(45).Width, BarHeight_GuiEXP, Tex_GUI(45).Width, BarHeight_GuiEXP

' render exp
If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
     sString = Int(GetPlayerExp(MyIndex) / TNL * 100) & "%"
Else
     sString = "Max Level"
End If
dX = x + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2) - 18
dY = y + 18
RenderText Font_Default, sString, dX, dY, White

' render gold
sString = Format(GetPlayerCoins(MyIndex), "###,###,###,###")
dX = (Tex_GUI(42).Width + 2) - (Len(sString) * 6)
dY = Tex_GUI(42).Height - 6
RenderText Font_Default, sString, dX, dY, White

' draw buttons
    For I = 63 To 64
        ' set co-ordinate
        x = GUIWindow(GUI_BARS).x + Buttons(I).x
        y = GUIWindow(GUI_BARS).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next

'Call DrawMainGUI
End Sub
Public Sub DrawEventChat()
Dim I As Long, x As Long, y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    x = GUIWindow(GUI_EVENTCHAT).x
    y = GUIWindow(GUI_EVENTCHAT).y
    
    ' render chatbox
    Width = GUIWindow(GUI_EVENTCHAT).Width
    Height = GUIWindow(GUI_EVENTCHAT).Height
    RenderTexture Tex_GUI(19), x, y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(chatText, GUIWindow(GUI_EVENTCHAT).Width - 20), x + 10, y + 22, White
    
    If chatOnlyContinue = False Then
        ' Draw replies
        For I = 1 To 4
            If Len(Trim$(chatOpt(I))) > 0 Then
                Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(I)) & "]")
                x = GUIWindow(GUI_CHAT).x + 95 + (155 - (Width / 2))
                y = GUIWindow(GUI_CHAT).y + 115 - ((I - 1) * 15)
                If chatOptState(I) = 2 Then
                    ' clicked
                    RenderText Font_Default, "[" & Trim$(chatOpt(I)) & "]", x, y, Grey
                Else
                    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                        ' hover
                        RenderText Font_Default, "[" & Trim$(chatOpt(I)) & "]", x, y, Yellow
                        ' play sound if needed
                        If Not lastNpcChatsound = I Then
                            PlaySound Sound_ButtonHover, -1, -1
                            lastNpcChatsound = I
                        End If
                    Else
                        ' normal
                        RenderText Font_Default, "[" & Trim$(chatOpt(I)) & "]", x, y, BrightBlue
                        ' reset sound if needed
                        If lastNpcChatsound = I Then lastNpcChatsound = 0
                    End If
                End If
            End If
        Next
    Else
        Width = EngineGetTextWidth(Font_Default, "[Continue]")
        x = GUIWindow(GUI_EVENTCHAT).x + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
        y = GUIWindow(GUI_EVENTCHAT).y + 100
        If chatContinueState = 2 Then
            ' clicked
            RenderText Font_Default, "[Continue]", x, y, Grey
        Else
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                ' hover
                RenderText Font_Default, "[Continue]", x, y, Yellow
                ' play sound if needed
                If Not lastNpcChatsound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = I
                End If
            Else
                ' normal
                RenderText Font_Default, "[Continue]", x, y, BrightBlue
                ' reset sound if needed
                If lastNpcChatsound = I Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub

Public Sub DrawMiniMap()
Dim I As Long, Z As Long
Dim x As Integer, y As Integer
Dim Direction As Byte
Dim CameraX As Long, CameraY As Long, playerNum As Long
Dim MapX As Long, MapY As Long
Dim CameraXSize As Long, CameraYSize As Long

    CameraXSize = MAX_MAPX * 32 + 1

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo ErrorHandler

        MapX = Map.MaxX
        MapY = Map.MaxY
        
        ' Draw Outline
        For x = 0 To MapX
                For y = 0 To MapY
                        CameraX = CameraXSize - (MapX * 4) + (x * 4)
                        CameraY = 55 + (y * 4)
                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(255, 255, 255, 150)
                Next y
        Next x

        ' Draw Tile Attribute
        For x = 0 To MapX
                For y = 0 To MapY
                    CameraX = CameraXSize - (MapX * 4) + (x * 4)
                    CameraY = 55 + (y * 4)
                        Select Case Map.Tile(x, y).Type
                                Case TILE_TYPE_BLOCKED
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(255, 0, 0, 200)
                                Case TILE_TYPE_WARP
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(75, 0, 155, 200)
                                Case TILE_TYPE_ITEM
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(0, 155, 0, 200)
                                Case TILE_TYPE_SHOP
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(255, 125, 0, 200)
                                Case TILE_TYPE_PLAYERSPAWN
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(0, 155, 255, 200)
                        End Select
                Next y
        Next x

         ' Draw Player dot
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If GetPlayerMap(I) = GetPlayerMap(MyIndex) And (Not GetPlayerVisible(I) = 1 Or I = MyIndex) Then
                        Select Case Player(I).PK
                                Case 0
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapPlayer(I).x)
                                        CameraY = 55 + (MiniMapPlayer(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(0, 255, 0, 200)
                                Case 1
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapPlayer(I).x)
                                        CameraY = 55 + (MiniMapPlayer(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(100, 0, 0, 200)
                        End Select
                End If
            End If
        Next I
        
        ' Draw NPC dot
        For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).num > 0 Then
                        Select Case NPC(I).Behaviour
                                Case NPC_BEHAVIOUR_ATTACKONSIGHT
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapNPC(I).x)
                                        CameraY = 55 + (MiniMapNPC(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(0, 0, 0, 200)
                                Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapNPC(I).x)
                                        CameraY = 55 + (MiniMapNPC(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(100, 0, 100, 200)
                                Case NPC_BEHAVIOUR_SHOPKEEPER
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapNPC(I).x)
                                        CameraY = 55 + (MiniMapNPC(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(255, 125, 0, 200)
                                Case NPC_BEHAVIOUR_FRIENDLY
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapNPC(I).x)
                                        CameraY = 55 + (MiniMapNPC(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(255, 0, 255, 200)
                                Case NPC_BEHAVIOUR_GUARD
                                        CameraX = CameraXSize - (MapX * 4) + (MiniMapNPC(I).x)
                                        CameraY = 55 + (MiniMapNPC(I).y)
                                        RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(0, 255, 255, 200)
                    End Select
                End If
        Next I

        ' Draw Tile Attribute
        For x = 0 To MapX
            For y = 0 To MapY
                For I = 1 To Map.CurrentEvents
                    If Map.MapEvents(I).Visible = 1 Then
                        If Map.MapEvents(I).x = x Then
                            If Map.MapEvents(I).y = y Then
                                CameraX = CameraXSize - (MapX * 4) + (x * 4)
                                CameraY = 55 + (y * 4)
                                    Select Case Map.MapEvents(I).ShowName
                                            Case 0 ' Tile
                                                    RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(0, 0, 155, 200)
                                            Case 1 ' Sprite
                                                    RenderTexture Tex_White, CameraX, CameraY, 0, 0, 4, 4, 4, 4, D3DColorRGBA(255, 255, 0, 200)
                                            End Select
                            End If
                        End If
                    End If
                Next I
            Next y
        Next x

        ' Error handler
        Exit Sub
ErrorHandler:
        HandleError "DrawMiniMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

' Guild Menu code by: Generalissimo Pony
Public Sub DrawGuildMenu()
Dim c As Long, GuildRank As String, RankNum As Long, M As Long, o As Long
Dim Width As Long, Height As Long, textW As Long, textL As Long, text As String, sInfo() As String, I As Long, y As Long, x As Long, gInfo() As String
Dim y1 As Long
    y = GUIWindow(GUI_GUILD).y + 5
    text = "Guild Name: " & Player(MyIndex).GuildName
        Width = GUIWindow(GUI_GUILD).Width
        Height = GUIWindow(GUI_GUILD).Height
    RenderTexture Tex_GUI(26), GUIWindow(GUI_GUILD).x, GUIWindow(GUI_GUILD).y, 0, 0, Width, Height, Width, Height
        
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    If Not Player(MyIndex).GuildName = vbNullString Then
        sInfo(I) = "Guild Name: " + Trim$(GuildData.Guild_Name)
    Else
        sInfo(I) = "Guild Name: "
    End If
    
    I = I + 1
    ReDim Preserve sInfo(1 To I) As String
    If Not Player(MyIndex).GuildName = vbNullString Then
        sInfo(I) = "Guild Tag: " + Trim$(GuildData.Guild_Tag)
    Else
        sInfo(I) = "Guild Tag: "
    End If
    
    I = I + 1
    ReDim Preserve sInfo(1 To I) As String
    sInfo(I) = "Message Of The Day:"
    
    I = I + 1
    ReDim Preserve sInfo(1 To I) As String
    If Not Player(MyIndex).GuildName = vbNullString Then
        sInfo(I) = Trim$(GuildData.Guild_MOTD)
    End If
    
    I = I + 1
    ReDim Preserve sInfo(1 To I) As String
    sInfo(I) = "- Guild Members -"
        
    If Not Player(MyIndex).GuildName = vbNullString Then
        For M = 1 To MAX_GUILD_MEMBERS
            If M > GuildScroll - (M - GuildScroll) - 2 And M < GuildScroll + 5 Then
                If Not GuildData.Guild_Members(M).User_Name = vbNullString Then
                    If GuildData.Guild_Members(M).Online = True Then
                        RenderText Font_Default, "-  " & GuildData.Guild_Members(M).User_Name, GUIWindow(GUI_GUILD).x + 40, GUIWindow(GUI_GUILD).y + 160 + ((M - GuildScroll) * 19), Green
                    Else
                        RenderText Font_Default, "-  " & GuildData.Guild_Members(M).User_Name, GUIWindow(GUI_GUILD).x + 40, GUIWindow(GUI_GUILD).y + 160 + ((M - GuildScroll) * 19), Red
                    End If
                End If
            End If
        Next M
    End If
    
      y = y + 12
    For I = 1 To UBound(sInfo)
        If Not I = 2 And Not I = 4 And Not Player(MyIndex).GuildName = vbNullString Then
            x = GUIWindow(GUI_GUILD).x + (GUIWindow(GUI_GUILD).Width / 2) - (Len(sInfo(I)) * 2)
            y = y + 25
            RenderText Font_Default, WordWrap(sInfo(I), Width - 65), x, y, Brown
        ElseIf I = 4 Then
            x = GUIWindow(GUI_GUILD).x + (GUIWindow(GUI_GUILD).Width / 2) - (Len(sInfo(I)) * 2)
            y = y + 20
            RenderText Font_Georgia, WordWrap(sInfo(I), Width - 65), x, y, Brown
        Else
            x = GUIWindow(GUI_GUILD).x + (GUIWindow(GUI_GUILD).Width / 2) - (Len(sInfo(I)) * 2)
            y = y + 25
            RenderText Font_Default, WordWrap(sInfo(I), Width - 65), x, y, GuildData.Guild_Color
        End If
    Next
    
    ' draw buttons
    For I = 43 To 44
        ' set co-ordinate
        x = GUIWindow(GUI_GUILD).x + Buttons(I).x
        y = GUIWindow(GUI_GUILD).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawQuestLog()
Dim buttonnum As Long, x As Long, y As Long
Dim Width As Long, Height As Long
    ' render the window
    Width = GUIWindow(GUI_QUESTLOG).Width
    Height = GUIWindow(GUI_QUESTLOG).Height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_QUESTLOG).x, GUIWindow(GUI_QUESTLOG).y, 0, 0, Width, Height, Width, Height
    
    ' draw the buttons
    For buttonnum = 45 To 50
        x = GUIWindow(GUI_QUESTLOG).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_QUESTLOG).y + Buttons(buttonnum).y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawPlayerInfo()
Dim x As Long, y As Long, I As Long
Dim Width As Long, Height As Long
    ' render the window
    Width = GUIWindow(GUI_FRIENDS).Width
    Height = GUIWindow(GUI_FRIENDS).Height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_FRIENDS).x, GUIWindow(GUI_FRIENDS).y, 0, 0, Width, Height, Width, Height

        Width = EngineGetTextWidth(Font_Georgia, "[X]")
        x = (GUIWindow(GUI_FRIENDS).x + GUIWindow(GUI_FRIENDS).Width - 25)
        y = GUIWindow(GUI_FRIENDS).y + 10
            If PlayerInfoX = 2 Then
                ' clicked
                RenderText Font_Georgia, "[X]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, "[X]", x, y, Cyan
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "[X]", x, y, Green
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
        'Display info
        For I = 1 To UBound(PlayerInfoText)
            If I = 1 Then Width = EngineGetTextWidth(Font_Georgia, PlayerInfoText(I))
            If I > 1 Then Width = EngineGetTextWidth(Font_Georgia, (PlayerInfoText(I) & PlayerInfoValue(I - 1)))
            
            If I = 1 Then
                x = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width / 2)) - (Width / 2)
                RenderText Font_Georgia, PlayerInfoText(I), x, y, White
            ElseIf I = 2 Then
                x = (GUIWindow(GUI_FRIENDS).x + 20)
                RenderText Font_Georgia, PlayerInfoText(I) & PlayerInfoValue(I - 1), x, y + 30, White
            Else
                x = (GUIWindow(GUI_FRIENDS).x + 20)
                RenderText Font_Georgia, PlayerInfoText(I) & PlayerInfoValue(I - 1), x, y + (I * 15) + 15, White
            End If
        Next I
End Sub

Public Sub DrawBuddies()
Dim x As Long, y As Long, buttonnum As Long
Dim Width As Long, Height As Long
    ' render the window
    Width = GUIWindow(GUI_FRIENDS).Width
    Height = GUIWindow(GUI_FRIENDS).Height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_FRIENDS).x, GUIWindow(GUI_FRIENDS).y, 0, 0, Width, Height, Width, Height

    ' set co-ordinate
    For buttonnum = 54 To 55
        x = GUIWindow(GUI_FRIENDS).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_FRIENDS).y + Buttons(buttonnum).y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' check for state
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    Next buttonnum
End Sub

Sub DrawBook()
Dim x As Long, y As Long, buttonnum As Long
Dim Width As Long, Height As Long
Static ms As Long
Dim Parse() As String

    ' render the window
    Width = GUIWindow(GUI_BOOK).Width
    Height = GUIWindow(GUI_BOOK).Height
    
    'open the book
    If Not OpeningBook Then
        RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
    Else
        ms = ms + 1
        If ms <= 2 Then
            RenderTexture Tex_GUI(29), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 4 Then
            RenderTexture Tex_GUI(30), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 6 Then
            RenderTexture Tex_GUI(31), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        Else
            ms = 0
            OpeningBook = False
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
        End If
    End If
    
    'turn the page to the left
    If Book_PageLeft Then
        ms = ms + 1
        If ms <= 2 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(33), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 4 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(34), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 6 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(35), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 8 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(36), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 10 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(39), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 12 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(38), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 14 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(37), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        Else
            ms = 0
            Book_PageLeft = False
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
        End If
        Exit Sub
    End If
    
    'turn the page to the left
    If Book_PageRight Then
        ms = ms + 1
        If ms <= 2 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(37), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 4 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(38), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 6 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(39), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 8 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(40), GUIWindow(GUI_BOOK).x + 275, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 10 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(35), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 12 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(34), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        ElseIf ms <= 14 Then
            Width = GUIWindow(GUI_BOOK).Width
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Width = 230
            RenderTexture Tex_GUI(33), GUIWindow(GUI_BOOK).x + 75, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
            Exit Sub
        Else
            ms = 0
            Book_PageRight = False
            RenderTexture Tex_GUI(32), GUIWindow(GUI_BOOK).x, GUIWindow(GUI_BOOK).y, 0, 0, Width, Height, Width, Height
        End If
        Exit Sub
    End If
    
    ' set co-ordinate
    For buttonnum = 56 To 58
        x = GUIWindow(GUI_BOOK).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_BOOK).y + Buttons(buttonnum).y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' check for state
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    Next buttonnum
End Sub

Public Sub DrawQuestDialogue()
Dim I As Long, x As Long, y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    x = GUIWindow(GUI_QUESTDIALOGUE).x
    y = GUIWindow(GUI_QUESTDIALOGUE).y
    
    ' render chatbox
    Width = GUIWindow(GUI_QUESTDIALOGUE).Width
    Height = GUIWindow(GUI_QUESTDIALOGUE).Height + 100 'QUICKCHANGE
    RenderTexture Tex_GUI(19), x, y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Georgia, WordWrap(QuestName, Width - 10), x + 10, y + 10, White
    RenderText Font_Georgia, WordWrap(QuestSubtitle, Width - 10), x + 10, y + 25, White
    RenderText Font_Georgia, WordWrap(QuestSay, Width - 10), x + 10, y + 40, White
    
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_QUESTDIALOGUE).y + 100
            If QuestAcceptState = 2 Then
                ' clicked
                RenderText Font_Georgia, "[Accept]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, "[Accept]", x, y, Cyan
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "[Accept]", x, y, Green
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_QUESTDIALOGUE).y + 100
            If QuestExtraState = 2 Then
                ' clicked
                RenderText Font_Georgia, "[" & QuestExtra & "]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, "[" & QuestExtra & "]", x, y, Cyan
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "[" & QuestExtra & "]", x, y, BrightRed
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_QUESTDIALOGUE).y + 120
    If QuestCloseState = 2 Then
        ' clicked
        RenderText Font_Georgia, "[Close]", x, y, Grey
    Else
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            ' hover
            RenderText Font_Georgia, "[Close]", x, y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 3 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 3
            End If
        Else
            ' normal
            RenderText Font_Georgia, "[Close]", x, y, Yellow
            ' reset sound if needed
            If lastNpcChatsound = 3 Then lastNpcChatsound = 0
        End If
    End If
End Sub

Public Sub DrawFriendRequest()
Dim I As Long, x As Long, y As Long, Sprite As Long, Width As Long
Dim Height As Long, tempText As String

    ' draw background
    x = GUIWindow(GUI_FRIENDREQUEST).x
    y = GUIWindow(GUI_FRIENDREQUEST).y
    
    ' render chatbox
    Width = GUIWindow(GUI_FRIENDREQUEST).Width
    Height = GUIWindow(GUI_FRIENDREQUEST).Height
    RenderTexture Tex_GUI(19), x, y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    tempText = FriendRequestSender & " is requesting your friendship. How do you respond?"
    RenderText Font_Georgia, WordWrap(tempText, Width - 10), x + 10, y + 10, White
    
    If FriendRequestVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_FRIENDREQUEST).y + 80
            If FriendRequestAcceptState = 2 Then
                ' clicked
                RenderText Font_Georgia, "[Accept]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, "[Accept]", x, y, Cyan
                    ' play sound if needed
                    If Not LastButtonSound_Main = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        LastButtonSound_Main = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "[Accept]", x, y, Green
                    ' reset sound if needed
                    If LastButtonSound_Main = 1 Then LastButtonSound_Main = 0
                End If
            End If
            
        Width = EngineGetTextWidth(Font_Georgia, "[Decline]")
        x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_FRIENDREQUEST).y + 100
            If FriendRequestDeclineState = 2 Then
                ' clicked
                RenderText Font_Georgia, "[Decline]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, "[Decline]", x, y, Cyan
                    ' play sound if needed
                    If Not LastButtonSound_Main = 2 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        LastButtonSound_Main = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "[Decline]", x, y, Green
                    ' reset sound if needed
                    If LastButtonSound_Main = 2 Then LastButtonSound_Main = 0
                End If
            End If
    End If
End Sub

' Combat Menu code by: Ertzel
Public Sub DrawCombat()
Dim c As Long
Dim Width As Long, Height As Long, textW As Long, textL As Long, text As String, sInfo() As String, I As Long, y As Long, x As Long, gInfo() As String
Dim y1 As Long
    y = GUIWindow(GUI_COMBAT).y
    Width = GUIWindow(GUI_COMBAT).Width
    Height = GUIWindow(GUI_COMBAT).Height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_COMBAT).x, GUIWindow(GUI_COMBAT).y, 0, 0, Width, Height, Width, Height
        
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    sInfo(I) = "/--------------\" & vbNewLine & "   Skill List   " & vbNewLine & "\--------------/"
    x = GUIWindow(GUI_COMBAT).x + (GUIWindow(GUI_COMBAT).Width / 2) - Len(sInfo(I)) + 10
    y = y + 5
    RenderText Font_Default, sInfo(I), x, y, Brown
        
    For c = 1 To MAX_COMBAT + MAX_SKILLS
        If c > CombatScroll - (c - CombatScroll) And c < CombatScroll + 5 Then
            Select Case c
                Case 1
                    text = "Large Blades "
                Case 2
                    text = "Small Blades "
                Case 3
                    text = "Blunt Weapons "
                Case 4
                    text = "Axes "
                Case 5
                    text = "Polearms "
                Case 6
                    text = "Mage Weapons "
                Case 7
                    text = "Body Magic "
                Case 8
                    text = "Soul Magic "
                Case Else
                    text = Trim$(Skill(c - 8).name) & " "
            End Select
            
            If c < 9 Then
                RenderText Font_Default, "-  " & text & "(" & Player(MyIndex).Combat(c).Level & ")", GUIWindow(GUI_COMBAT).x + 20, GUIWindow(GUI_COMBAT).y + ((c - CombatScroll) * 50), Brown
                RenderText Font_Default, "Exp: " & Player(MyIndex).Combat(c).EXP & " / " & CombatTNL(c), GUIWindow(GUI_COMBAT).x + 20, GUIWindow(GUI_COMBAT).y + ((c - CombatScroll) * 50) + 20, Brown
            Else
                RenderText Font_Default, "-  " & text & "(" & Player(MyIndex).Skills(c - 8).Level & ")", GUIWindow(GUI_COMBAT).x + 20, GUIWindow(GUI_COMBAT).y + ((c - CombatScroll) * 50), Brown
                RenderText Font_Default, "Exp: " & Player(MyIndex).Skills(c - 8).EXP & " / " & Player(MyIndex).Skills(c - 8).EXP_Needed, GUIWindow(GUI_COMBAT).x + 20, GUIWindow(GUI_COMBAT).y + ((c - CombatScroll) * 50) + 20, Brown
            End If
        End If
    Next c
    
    ' draw buttons
    For I = 52 To 53
        ' set co-ordinate
        x = GUIWindow(GUI_COMBAT).x + Buttons(I).x
        y = GUIWindow(GUI_COMBAT).y + Buttons(I).y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub LoadDX8Vars()
    MAX_MAPX = (frmMain.Width / 503.75) '24
    MAX_MAPY = (frmMain.Height / 523.33) '18
    HalfX = ((MAX_MAPX + 1) / 2) * PIC_X
    HalfY = ((MAX_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MAX_MAPX + 1) * PIC_X
    ScreenY = (MAX_MAPY + 1) * PIC_Y
    StartXValue = ((MAX_MAPX + 1) / 2)
    StartYValue = ((MAX_MAPY + 1) / 2)
    EndXValue = (MAX_MAPX + 1) + 1
    EndYValue = (MAX_MAPY + 1) + 1
    Half_PIC_X = PIC_X / 2
    Half_PIC_Y = PIC_Y / 2
End Sub

' player Projectiles
Public Sub DrawProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, PicNum As Long, I As Long
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' check for subscript error
    If Index < 1 Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' check to see if it's time to move the Projectile
    If GetTickCount > Player(Index).ProjecTile(PlayerProjectile).TravelTime Then
        With Player(Index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case 0
                    .y = .y + 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' up
                Case 1
                    .y = .y - 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' right
                Case 2
                    .x = .x + 1
                    ' check if they reached max range
                    If .x = (GetPlayerX(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' left
                Case 3
                    .x = .x - 1
                    ' check if they reached maxrange
                    If .x = (GetPlayerX(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetTickCount + .speed
        End With
    End If
    
    ' set the x, y & pic values for future reference
    x = Player(Index).ProjecTile(PlayerProjectile).x
    y = Player(Index).ProjecTile(PlayerProjectile).y
    PicNum = Player(Index).ProjecTile(PlayerProjectile).Pic
    
    ' check if left map
    If x > Map.MaxX Or y > Map.MaxY Or x < 0 Or y < 0 Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit a block or resource
    If Map.Tile(x, y).Type = TILE_TYPE_BLOCKED Or Map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check for player hit
    For I = 1 To Player_HighIndex
        If x = GetPlayerX(I) And y = GetPlayerY(I) Then
            ' they're hit, remove it
            If Not x = Player(MyIndex).x Or Not y = GetPlayerY(MyIndex) Then
                ClearProjectile Index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' check for npc hit
    For I = 1 To MAX_MAP_NPCS
        If x = MapNpc(I).x And y = MapNpc(I).y Then
            ' they're hit, remove it
            ClearProjectile Index, PlayerProjectile
            Exit Sub
        End If
    Next

    
    ' get positioning in the texture
    With rec
        .Top = 0
        .Bottom = SIZE_Y
        .Left = Player(Index).ProjecTile(PlayerProjectile).Direction * SIZE_X
        .Right = .Left + SIZE_X
    End With

    ' blt the projectile
    RenderTexture Tex_Projectile(PicNum), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    'Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Projectile(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "BltProjectile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
