VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   833
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7800
      TabIndex        =   58
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NEW Options"
      Height          =   5895
      Left            =   8520
      TabIndex        =   21
      Top             =   120
      Width           =   3855
      Begin VB.HScrollBar scrlSkillReqLvl 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   53
         Top             =   3720
         Width           =   3615
      End
      Begin VB.ComboBox cmbSkillReq 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   120
         List            =   "frmEditor_Resource.frx":3334
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox txtExp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   9
         TabIndex        =   49
         Text            =   "1"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cmbSkill 
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CheckBox chkSkillExp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Give Skill Exp"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   3615
      End
      Begin VB.CheckBox chkDistItems 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Distribute items during attack? [NEW]"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   3495
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   35
         Top             =   4440
         Width           =   3615
      End
      Begin VB.CheckBox chkRandHP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Random? [NEW]"
         Height          =   180
         Left            =   720
         TabIndex        =   34
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtHPMax 
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtHPMin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   30
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   23
         Top             =   5040
         Width           =   3615
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3720
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label lblSkillLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Required Skill Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Required Skill:"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3720
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3600
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         Height          =   180
         Left            =   120
         TabIndex        =   46
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   180
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   4200
         Width           =   2100
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max:"
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min:"
         Height          =   180
         Left            =   2160
         TabIndex        =   31
         Top             =   480
         Width           =   330
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   4800
         Width           =   1260
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   5430
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9360
      TabIndex        =   20
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10920
      TabIndex        =   19
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   7080
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resource Properties"
      Height          =   6855
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbColor_Empty 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3336
         Left            =   1800
         List            =   "frmEditor_Resource.frx":3370
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1800
         Width           =   3135
      End
      Begin VB.ComboBox cmbColor_Success 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3410
         Left            =   1800
         List            =   "frmEditor_Resource.frx":344A
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   960
         Width           =   3135
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   43
         Top             =   6480
         Width           =   4815
      End
      Begin VB.CheckBox chkRewardRand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Random? [NEW]"
         Height          =   180
         Left            =   1800
         TabIndex        =   42
         Top             =   5640
         Width           =   2895
      End
      Begin VB.TextBox txtAmountMin 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3480
         TabIndex        =   41
         Text            =   "0"
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txtAmountMax 
         Height          =   270
         Left            =   600
         TabIndex        =   38
         Text            =   "0"
         Top             =   5880
         Width           =   1455
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   2880
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   15
         Top             =   3240
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":34EA
         Left            =   960
         List            =   "frmEditor_Resource.frx":34FA
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2280
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2295
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   6
         Top             =   3240
         Width           =   2280
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "^-> Color"
         Height          =   180
         Left            =   960
         TabIndex        =   56
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "^-> Color"
         Height          =   180
         Left            =   960
         TabIndex        =   55
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   6240
         Width           =   1530
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Min:"
         Height          =   180
         Left            =   3000
         TabIndex        =   40
         Top             =   5880
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max:"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   5880
         Width           =   375
      End
      Begin VB.Label lblRewardAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Reward Amount:"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   5640
         Width           =   1635
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   5040
         Width           =   4800
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   17
         Top             =   2640
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resource List"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6360
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDistItems_Click()
    Resource(EditorIndex).DistItems = chkDistItems.value
End Sub

Private Sub chkRandHP_Click()
    txtHPMin.Enabled = chkRandHP.value
    Resource(EditorIndex).HPRand = chkRandHP.value
End Sub

Private Sub chkRewardRand_Click()
    Resource(EditorIndex).ItemRewardRand = chkRewardRand.value
    txtAmountMin.Enabled = chkRewardRand.value
End Sub

Private Sub chkSkillExp_Click()
    cmbSkill.Enabled = chkSkillExp.value
    txtEXP.Enabled = chkSkillExp.value
    Resource(EditorIndex).Exp_Give = chkSkillExp.value
End Sub

Private Sub cmbColor_Empty_Click()
    If Not EditorIndex > 0 Then Exit Sub
    If Len(cmbColor_Empty.text) > 0 Then
        Resource(EditorIndex).Color_Empty = cmbColor_Empty.ListIndex
    End If
End Sub

Private Sub cmbColor_Success_Click()
    If Not EditorIndex > 0 Then Exit Sub
    If Len(cmbColor_Success.text) > 0 Then
        Resource(EditorIndex).Color_Success = cmbColor_Success.ListIndex
    End If
End Sub

Private Sub cmbSkill_Click()
    If Len(cmbSkill.text) > 0 Then
        Resource(EditorIndex).Exp_Skill = cmbSkill.ListIndex + 1
    End If
End Sub

Private Sub cmbSkillReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not cmbSkillReq.ListIndex < 0 Then
        Resource(EditorIndex).Skill_Req = cmbSkillReq.ListIndex
        scrlSkillReqLvl.max = Skill(cmbSkillReq.ListIndex + 1).MaxLvl
    End If
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbSkillReq_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ResourceEditorOk(False)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlReward.max = MAX_ITEMS
    cmbColor_Success.ListIndex = 10
    cmbColor_Empty.ListIndex = 12
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).name)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.value
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.value
    Resource(EditorIndex).ResourceImage = scrlNormalPic.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRespawn_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.value
    Resource(EditorIndex).RespawnTime = scrlRespawn.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRespawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlReward.value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.value).name)
    Else
        lblReward.Caption = "Item Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.value

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSkillReqLvl_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSkillLvl.Caption = "Required Skill Level: " & scrlSkillReqLvl.value
    Resource(EditorIndex).Skill_LvlReq = scrlSkillReqLvl.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSkillReqLvl_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTool_Change()
    Dim name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Select Case scrlTool.value
        Case 0
            name = "None"
        Case 1
            name = "Hatchet"
        Case 2
            name = "Rod"
        Case 3
            name = "Pickaxe"
    End Select

    lblTool.Caption = "Tool Required: " & name
    
    Resource(EditorIndex).ToolRequired = scrlTool.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAmountMin_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtAmountMin.text) And Len(txtAmountMin.text) > 0 Then
        txtAmountMin.text = "1"
        txtAmountMin.SelStart = Len(txtAmountMin.text)
    Else
        If Len(txtAmountMin.text) = 0 Then
            Resource(EditorIndex).ItemRewardAmount = 1
            Exit Sub
        End If
    End If
    
    If IsNumeric(txtAmountMax.text) And IsNumeric(txtAmountMin.text) Then
        If CLng(txtAmountMin.text) > CLng(txtAmountMax.text) Then
            txtAmountMin.text = txtAmountMax.text
        End If
    End If
    
    Resource(EditorIndex).ItemRewardAmountMin = CLng(txtAmountMin.text)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtAmountMin_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    If Not IsNumeric(txtEXP.text) Then
        If Len(txtEXP.text) > 0 Then
            txtEXP.text = "1"
        Else
            Resource(EditorIndex).Exp_Amnt = 1
            Exit Sub
        End If
    End If
    
    Resource(EditorIndex).Exp_Amnt = CLng(txtEXP.text)
End Sub

Private Sub txtHPMax_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtHPMax.text) Then
        If Len(txtHPMax.text) > 0 Then
            txtHPMax.text = "1"
        Else
            Resource(EditorIndex).health = 1
            Exit Sub
        End If
    End If
    
    If IsNumeric(txtHPMax.text) And IsNumeric(txtHPMin.text) Then
        If CLng(txtHPMin.text) > CLng(txtHPMax.text) Then
            txtHPMin.text = txtHPMax.text
        End If
    End If
    
    Resource(EditorIndex).health = CLng(txtHPMax.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtHPMax_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHPMin_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtHPMin.text) Then
        If Len(txtHPMin.text) > 0 Then
            txtHPMin.text = "1"
        Else
            Resource(EditorIndex).healthmin = 1
            Exit Sub
        End If
    End If
    
    Resource(EditorIndex).healthmin = CLng(txtHPMin.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtHPMin_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAmountMax_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtAmountMax.text) And Len(txtAmountMax.text) > 0 Then
        txtAmountMax.text = "1"
        txtAmountMax.SelStart = Len(txtAmountMax.text)
    Else
        If Len(txtAmountMax.text) = 0 Then
            Resource(EditorIndex).ItemRewardAmount = 1
            Exit Sub
        End If
    End If
    
    If IsNumeric(txtAmountMax.text) And IsNumeric(txtAmountMin.text) Then
        If CLng(txtAmountMin.text) > CLng(txtAmountMax.text) Then
            txtAmountMin.text = txtAmountMax.text
        End If
    End If
    
    Resource(EditorIndex).ItemRewardAmount = CLng(txtAmountMax.text)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtAmountMax_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
