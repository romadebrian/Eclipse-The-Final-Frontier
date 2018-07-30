VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   762
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1264
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrScrollEditor 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   4680
   End
   Begin VB.ListBox lstFriends 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      ItemData        =   "frmMain.frx":3332
      Left            =   600
      List            =   "frmMain.frx":3334
      MousePointer    =   1  'Arrow
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00B5B5B5&
      ForeColor       =   &H80000008&
      Height          =   5730
      Left            =   6480
      ScaleHeight     =   380
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   347
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   5235
      Begin VB.CommandButton btnWalkthrough 
         Caption         =   "Toggle WalkThrough"
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton cmdAKill 
         Caption         =   "Auto Kill"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAHeal 
         Caption         =   "Full Heal"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAName 
         Caption         =   "Set Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSSMap 
         Caption         =   "Screenshot Map"
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level Up"
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   4440
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   2760
         Min             =   1
         TabIndex        =   23
         Top             =   4080
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   2760
         Min             =   1
         TabIndex        =   22
         Top             =   3480
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Spell"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Resource"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Del Bans"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Loc"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAVisible 
         Caption         =   "Visibility"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdACharEdit 
         Caption         =   "Character"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAQuest 
         Caption         =   "Quest"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Give me more controls using VB6"
         Height          =   495
         Left            =   2880
         TabIndex        =   46
         Top             =   720
         Width           =   2055
      End
      Begin VB.Line Line6 
         X1              =   176
         X2              =   176
         Y1              =   8
         Y2              =   368
      End
      Begin VB.Line Line5 
         X1              =   184
         X2              =   336
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   184
         X2              =   336
         Y1              =   208
         Y2              =   208
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   336
         Y1              =   376
         Y2              =   376
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Editors:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   168
         Y1              =   248
         Y2              =   248
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   192
         Y2              =   192
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Panel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.ListBox lstQuestLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      ItemData        =   "frmMain.frx":3336
      Left            =   3360
      List            =   "frmMain.frx":3338
      MousePointer    =   1  'Arrow
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   17280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   9960
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long
Private mouseClicked As Boolean

Private Sub btnWalkthrough_Click()
    SendWalkthrough
End Sub

Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub

    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdACharEdit_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    
    frmEditor_Character.Visible = True
End Sub

Private Sub cmdAHeal_Click()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub

    If IsNumeric(Trim$(txtAName.text)) Then Exit Sub

    If Len(Trim$(txtAName.text)) > 2 Then
        SendHealPlayer Trim$(txtAName.text)
    Else
        If Len(txtAName.text) = 0 Then SendHealPlayer GetPlayerName(MyIndex)
    End If
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAHeal_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAKill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub

    If Len(Trim$(txtAName.text)) < 2 Then Exit Sub

    If IsNumeric(Trim$(txtAName.text)) Then Exit Sub

    SendKillPlayer Trim$(txtAName.text)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAKill_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAName_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    
    Exit Sub
    End If
    
    If Len(Trim$(txtAName.text)) < 2 Then
    Exit Sub
    End If
    
    If IsNumeric(Trim$(txtAName.text)) Or IsNumeric(Trim$(txtAAccess.text)) Then
    Exit Sub
    End If
    
    SendSetName Trim$(txtAName.text), (Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAName_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAQuest_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditQuest
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAVisible_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendVisibility
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        SendRequestLevelUp GetPlayerName(MyIndex)
    Else
        SendRequestLevelUp Trim$(txtAName.text)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' render the map temp
    ScreenshotMap
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Click()
    HandleSingleClick
End Sub

Private Sub Form_DblClick()
    HandleDoubleClick
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' move GUI
    picAdmin.Left = 444
    picAdmin.Top = 8
    mouseClicked = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseDown Button
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseUp Button
End Sub

Private Sub Form_Resize()
    If Not frmMain.Visible Then Exit Sub
    ' Fullscreen work
    'LoadDX8Vars
    'InitDX8
    'GUIWindow(GUI_CHAT).y = frmMain.ScaleHeight - 155
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleMouseMove CLng(x), CLng(y), Button
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Function IsShopItem(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsShopItem = 0

    For I = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.Bottom Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub lstFriends_DblClick()
Dim Parse() As String
    If lstFriends.ListIndex < 0 Then Exit Sub ' Nothing selected
    If Not Len(lstFriends.List(lstFriends.ListIndex)) > 0 Then Exit Sub 'No name in selection
    'If InStr(lstFriends.List(lstFriends.ListIndex), "Offline") > 0 Then Exit Sub ' Player is offline
    
    'This will load a gui for the player's data.
    Parse() = Split(lstFriends.List(lstFriends.ListIndex), " ")
    Call RequestFriendData(Parse(0))
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.value).name)
    If Item(scrlAItem.value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.value).Stackable > 0 Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleKeyUp KeyCode

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
    
    Exit Sub
    End If
    
    If Len(Trim$(txtASprite.text)) < 1 Then
    Exit Sub
    End If
    
    If Not IsNumeric(Trim$(txtASprite.text)) Then
    Exit Sub
    End If
    
    If Len(Trim$(txtAName.text)) > 1 Then
    SendSetSprite CLng(Trim$(txtASprite.text)), txtAName.text
    Else
    SendSetSprite CLng(Trim$(txtASprite.text)), GetPlayerName(MyIndex)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    
    SendSpawnItem scrlAItem.value, scrlAAmount.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picAdmin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseClicked = True
End Sub

Private Sub picAdmin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseClicked = False
End Sub

Private Sub picAdmin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mouseClicked Then
        picAdmin.Left = picAdmin.Left + x
        picAdmin.Top = picAdmin.Top + y
    End If
End Sub

Private Sub tmrScrollEditor_Timer()
    Scroll_Timer = Scroll_Timer + 1
    
    If Scroll_Timer >= 10 Then '1 second
        Select Case Scroll_Editor
            Case 1 'map editor
                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                SendRequestEditMap
            Case 2 'npc editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditNpc
            Case 3 'item editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditItem
            Case 4 'resource editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditResource
            Case 5 'quest editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditQuest
            Case 6 'spell editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditSpell
            Case 7 'character editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                frmEditor_Character.Visible = True
                frmEditor_Character.txtEName.text = GetPlayerName(MyIndex)
            Case 8 'animation editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditAnimation
            Case 9 'shop editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditShop
            Case 10 'combo editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditCombo
        End Select
Continue:
        
        Scroll_Timer = 0
        Scroll_Editor = 0
        Scroll_Draw = False
        tmrScrollEditor.Enabled = False
    End If
End Sub
