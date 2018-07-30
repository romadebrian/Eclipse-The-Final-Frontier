VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13800
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   8640
      TabIndex        =   61
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   7560
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   12240
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spell Properties"
      Height          =   7335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NEW Data"
         Height          =   6975
         Left            =   6840
         TabIndex        =   62
         Top             =   240
         Width           =   3375
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   65
            Top             =   2040
            Width           =   3135
         End
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   64
            Top             =   600
            Width           =   3135
         End
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   63
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Neutral Dmg: 0"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Light Dmg: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   975
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dark Dmg: 0"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   66
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   960
         End
      End
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   6240
         Width           =   2775
      End
      Begin VB.HScrollBar scrlCombatLvl 
         Height          =   255
         LargeChange     =   10
         Left            =   5760
         Max             =   100
         TabIndex        =   56
         Top             =   6840
         Width           =   975
      End
      Begin VB.ComboBox cmdCombatType 
         Height          =   300
         ItemData        =   "frmEditor_Spell.frx":0000
         Left            =   4800
         List            =   "frmEditor_Spell.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   6360
         Width           =   1575
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data"
         Height          =   5895
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtVital 
            Height          =   270
            Left            =   120
            TabIndex        =   60
            Text            =   "0"
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   5520
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4920
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4320
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Area of Effect spell?"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   3240
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   37
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   35
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            Max             =   3
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblStun 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   4080
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblRange 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Interval: 0s"
            Height          =   255
            Left            =   1680
            TabIndex        =   36
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Duration: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVital 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vital: "
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblDir 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dir: Up"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            BackColor       =   &H00E0E0E0&
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Basic Information"
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   49
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   300
            TabIndex        =   32
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   300
            TabIndex        =   30
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0040
            Left            =   120
            List            =   "frmEditor_Spell.frx":0053
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblCool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Class Required:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label lblCombatLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Combat Level: 0"
         Height          =   180
         Left            =   4320
         TabIndex        =   59
         Top             =   6840
         Width           =   1245
      End
      Begin VB.Label lblCombatType 
         BackStyle       =   0  'Transparent
         Caption         =   "Combat Type:"
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spell List"
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If chkAOE.value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).Type = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCombatType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    Spell(EditorIndex).CombatTypeReq = cmdCombatType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCombatType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorOk False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAccess.value > 0 Then
        lblAccess.Caption = "Access Required: " & scrlAccess.value
    Else
        lblAccess.Caption = "Access Required: None"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAnim.value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.value).name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAnimCast.value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.value).name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAOE.value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    Spell(EditorIndex).AoE = scrlAOE.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblCast.Caption = "Casting Time: " & scrlCast.value & "s"
    Spell(EditorIndex).CastTime = scrlCast.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCombatLvl_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    lblCombatLvl.Caption = "Combat Level: " & scrlCombatLvl
    Spell(EditorIndex).CombatLvlReq = scrlCombatLvl.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCombatLvl_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblCool.Caption = "Cooldown Time: " & scrlCool.value & "s"
    Spell(EditorIndex).CDTime = scrlCool.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Select Case scrlDir.value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = "Dir: " & sDir
    Spell(EditorIndex).Dir = scrlDir.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblDuration.Caption = "Duration: " & scrlDuration.value & "s"
    Spell(EditorIndex).Duration = scrlDuration.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlElement_Change(Index As Integer)
Dim txt As String
    Select Case Index
        Case 1
            txt = "Light Damage: "
            Spell(EditorIndex).Dmg_Light = scrlElement(Index).value
        Case 2
            txt = "Dark Damage: "
            Spell(EditorIndex).Dmg_Dark = scrlElement(Index).value
        Case 3
            txt = "Neutral Damage: "
            Spell(EditorIndex).Dmg_Neut = scrlElement(Index).value
    End Select
    
    lblElement(Index).Caption = txt & scrlElement(Index).value
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlIcon.value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlInterval_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblInterval.Caption = "Interval: " & scrlInterval.value & "s"
    Spell(EditorIndex).Interval = scrlInterval.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlLevel.value > 0 Then
        lblLevel.Caption = "Level Required: " & scrlLevel.value
    Else
        lblLevel.Caption = "Level Required: None"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblMap.Caption = "Map: " & scrlMap.value
    Spell(EditorIndex).Map = scrlMap.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlMP.value > 0 Then
        lblMP.Caption = "MP Cost: " & scrlMP.value
    Else
        lblMP.Caption = "MP Cost: None"
    End If
    Spell(EditorIndex).MPCost = scrlMP.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlRange.value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).Range = scrlRange.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlStun.value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblX.Caption = "X: " & scrlX.value
    Spell(EditorIndex).x = scrlX.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblY.Caption = "Y: " & scrlY.value
    Spell(EditorIndex).y = scrlY.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not IsNumeric(txtVital.text) And Len(txtVital.text) > 0 Then
        txtVital.text = 1
        txtVital.SelStart = Len(txtVital.text)
    ElseIf Len(txtVital.text) < 1 Then
        Spell(EditorIndex).Vital = 1
        Exit Sub
    End If
    
    Spell(EditorIndex).Vital = CLng(txtVital.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
