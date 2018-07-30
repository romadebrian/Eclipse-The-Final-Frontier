VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12615
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   841
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7920
      TabIndex        =   69
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame fraQuest 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quest"
      Height          =   615
      Left            =   8520
      TabIndex        =   60
      Top             =   120
      Width           =   4095
      Begin VB.HScrollBar scrlQuest 
         Height          =   255
         Left            =   1440
         TabIndex        =   62
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkQuest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quest giver?"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblQuest 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   63
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NEW Settings"
      Height          =   7215
      Left            =   8520
      TabIndex        =   47
      Top             =   840
      Width           =   3975
      Begin VB.CheckBox chkRndSpawn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Randomize Spawn Rate?"
         Height          =   180
         Left            =   120
         TabIndex        =   87
         Top             =   5160
         Width           =   3615
      End
      Begin VB.TextBox txtSpawnSecsMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   85
         Text            =   "0"
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   83
         Text            =   "0"
         Top             =   5520
         Width           =   1095
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   75
         Top             =   2880
         Width           =   2175
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   74
         Top             =   2520
         Width           =   2175
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   73
         Top             =   3600
         Width           =   2175
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   72
         Top             =   3960
         Width           =   2175
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   71
         Top             =   3240
         Width           =   2175
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   70
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox txtHPMin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   59
         Text            =   "0"
         ToolTipText     =   "Minimum amount of health the NPC can have."
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkRandHP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Randomize Health?"
         Height          =   180
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton opPercent_20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Within 20%"
         Height          =   255
         Left            =   2640
         TabIndex        =   55
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton opPercent_10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Within 10%"
         Height          =   255
         Left            =   1320
         TabIndex        =   54
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton opPercent_5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Within 5%"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2040
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   600
         TabIndex        =   51
         Text            =   "0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   600
         TabIndex        =   49
         Text            =   "0"
         ToolTipText     =   "Max health the NPC can have."
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkRndExp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Randomize Exp?"
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min:"
         Height          =   180
         Left            =   2160
         TabIndex        =   86
         Top             =   5520
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max:"
         Height          =   180
         Left            =   120
         TabIndex        =   84
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Rate (in seconds)"
         Height          =   180
         Left            =   120
         TabIndex        =   82
         Top             =   4800
         UseMnemonic     =   0   'False
         Width           =   1845
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3840
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3840
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Dmg: 0"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   81
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   960
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Dmg: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   80
         Top             =   2520
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Resist: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   79
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Resist: 0"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   78
         Top             =   3960
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neutral Dmg: 0"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   77
         Top             =   3240
         UseMnemonic     =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neutral Resist: 0"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   76
         Top             =   4320
         UseMnemonic     =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min:"
         Height          =   180
         Left            =   2160
         TabIndex        =   58
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exp Varies: 0 - 0"
         Height          =   255
         Left            =   1800
         TabIndex        =   56
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max:"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   8160
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11040
      TabIndex        =   35
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9480
      TabIndex        =   34
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NPC Properties"
      Height          =   7935
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.HScrollBar scrlMoveSpeed 
         Height          =   255
         Left            =   2640
         Max             =   10
         Min             =   1
         TabIndex        =   46
         Top             =   3240
         Value           =   1
         Width           =   2175
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3240
         TabIndex        =   40
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   39
         Top             =   2400
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   2760
         Width           =   2175
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   28
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   27
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1320
         List            =   "frmEditor_NPC.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   24
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame fraDrop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Drop"
         Height          =   2775
         Left            =   120
         TabIndex        =   15
         Top             =   5040
         Width           =   4815
         Begin VB.HScrollBar scrlDropIndex 
            Height          =   255
            Left            =   120
            Max             =   255
            Min             =   1
            TabIndex        =   88
            Top             =   480
            Value           =   1
            Width           =   4575
         End
         Begin VB.OptionButton opPercent 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Within 5%"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   67
            Top             =   2400
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton opPercent 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Within 10%"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   66
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton opPercent 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Within 20%"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   65
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkRandCurrency 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Randomize Currency?"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2160
            Width           =   1935
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   18
            Top             =   1800
            Width           =   3495
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   17
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Width           =   1815
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   4680
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblDI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index: 1"
            Height          =   180
            Left            =   120
            TabIndex        =   89
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblCurOutput 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Currency Varies: 0 - 0"
            Height          =   255
            Left            =   2160
            TabIndex        =   68
            Top             =   2160
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chance 1 out of"
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            Max             =   255
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            Max             =   255
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   6
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            Max             =   255
            TabIndex        =   5
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   13
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   12
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   10
            Top             =   1080
            Width           =   480
         End
      End
      Begin VB.Label lblMoveSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Speed: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   180
         Left            =   2640
         TabIndex        =   41
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: None"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NPC List"
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7620
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
      Top             =   8160
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkRandCurrency_Click()
    NPC(EditorIndex).Drops(scrlDropIndex.value).RandCurrency = chkRandCurrency.value
    lblCurOutput.Visible = chkRandCurrency.value
    opPercent(0).Visible = chkRandCurrency.value
    opPercent(1).Visible = chkRandCurrency.value
    opPercent(2).Visible = chkRandCurrency.value
    'recheck varies text
    If Not chkRandCurrency.value = vbChecked Then Exit Sub
    If opPercent(0).value Then Call opPercent_Click(0)
    If opPercent(1).value Then Call opPercent_Click(1)
    If opPercent(2).value Then Call opPercent_Click(2)
End Sub

Private Sub chkRandHP_Click()
    txtHPMin.Enabled = chkRandHP.value
    NPC(EditorIndex).RandHP = chkRandHP.value
End Sub

Private Sub chkRndExp_Click()
Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long
    NPC(EditorIndex).RandExp = chkRndExp.value
    opPercent_5.Visible = chkRndExp.value
    opPercent_10.Visible = chkRndExp.value
    opPercent_20.Visible = chkRndExp.value
    lblOutput.Visible = chkRndExp.value
    
    'recheck varies text
    If Not chkRndExp.value = vbChecked Then Exit Sub
    If opPercent_5.value Then Call opPercent_5_Click
    If opPercent_10.value Then Call opPercent_10_Click
    If opPercent_20.value Then Call opPercent_20_Click
End Sub

Private Sub chkRndSpawn_Click()
    txtSpawnSecsMin.Enabled = chkRndSpawn.value
    NPC(EditorIndex).RndSpawn = chkRndSpawn.value
End Sub

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call NpcEditorOk(False)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    scrlNum.max = MAX_ITEMS
    scrlDropIndex.max = MAX_NPC_DROP_ITEMS
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub opPercent_10_Click()
Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).Percent_5 = Abs(CInt(opPercent_5.value))
    NPC(EditorIndex).Percent_10 = Abs(CInt(opPercent_10.value))
    NPC(EditorIndex).Percent_20 = Abs(CInt(opPercent_20.value))
    
    If Not IsNumeric(txtEXP.text) Then Exit Sub
    If lblOutput.Visible Then
        ThisExp = CLng(txtEXP.text)
        RangeLow = ThisExp - (ThisExp * 0.1)
        RangeHigh = ThisExp + (ThisExp * 0.1)
        lblOutput.Caption = "Exp Varies: " & RangeLow & " - " & RangeHigh
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "opPercent_10_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub opPercent_20_Click()
Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).Percent_5 = Abs(CInt(opPercent_5.value))
    NPC(EditorIndex).Percent_10 = Abs(CInt(opPercent_10.value))
    NPC(EditorIndex).Percent_20 = Abs(CInt(opPercent_20.value))
    
    If Not IsNumeric(txtEXP.text) Then Exit Sub
    If lblOutput.Visible Then
        ThisExp = CLng(txtEXP.text)
        RangeLow = ThisExp - (ThisExp * 0.2)
        RangeHigh = ThisExp + (ThisExp * 0.2)
        lblOutput.Caption = "Exp Varies: " & RangeLow & " - " & RangeHigh
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "opPercent_20_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub opPercent_5_Click()
Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).Percent_5 = Abs(CInt(opPercent_5.value))
    NPC(EditorIndex).Percent_10 = Abs(CInt(opPercent_10.value))
    NPC(EditorIndex).Percent_20 = Abs(CInt(opPercent_20.value))
    
    If Not IsNumeric(txtEXP.text) Then Exit Sub
    If lblOutput.Visible Then
        ThisExp = CLng(txtEXP.text)
        RangeLow = ThisExp - (ThisExp * 0.05)
        RangeHigh = ThisExp + (ThisExp * 0.05)
        lblOutput.Caption = "Exp Varies: " & RangeLow & " - " & RangeHigh
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "opPercent_5_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub opPercent_Click(Index As Integer)
Dim RangeLow As Long, RangeHigh As Long, ThisCur As Long, CurMulti As Double
    ' set the array value
    If Index = 0 Then NPC(EditorIndex).Drops(scrlDropIndex.value).P_5 = Abs(CInt(opPercent(Index).value))
    If Index = 1 Then NPC(EditorIndex).Drops(scrlDropIndex.value).P_10 = Abs(CInt(opPercent(Index).value))
    If Index = 2 Then NPC(EditorIndex).Drops(scrlDropIndex.value).P_20 = Abs(CInt(opPercent(Index).value))
    
    'make sure we're good to go
    If Not scrlNum.value = 1 Then Exit Sub
    If Not scrlValue.value > 0 Then Exit Sub
    
    'get curmulti
    If Index = 0 Then CurMulti = 0.05
    If Index = 1 Then CurMulti = 0.1
    If Index = 2 Then CurMulti = 0.2
    
    If lblCurOutput.Visible Then
        ThisCur = scrlValue.value
        RangeLow = ThisCur - (ThisCur * CurMulti)
        RangeHigh = ThisCur + (ThisCur * CurMulti)
        lblCurOutput.Caption = "Currency Varies: " & RangeLow & " - " & RangeHigh
    End If
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).Name)
    lblAnimation.Caption = "Anim: " & sString
    NPC(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDropIndex_Change()
    lblDI.Caption = "Index: " & scrlDropIndex.value
    NpcEditorInit
End Sub

Private Sub scrlElement_Change(Index As Integer)
Dim text As String
Dim value As Long
    Select Case Index
        Case 1
            text = "Light Dmg: "
            NPC(EditorIndex).Element_Light_Dmg = scrlElement(Index).value
        Case 2
            text = "Dark Dmg: "
            NPC(EditorIndex).Element_Dark_Dmg = scrlElement(Index).value
        Case 3
            text = "Neutral Dmg: "
            NPC(EditorIndex).Element_Neut_Dmg = scrlElement(Index).value
        Case 4
            text = "Light Resist: "
            NPC(EditorIndex).Element_Light_Res = scrlElement(Index).value
        Case 5
            text = "Dark Resist: "
            NPC(EditorIndex).Element_Dark_Res = scrlElement(Index).value
        Case 6
            text = "Neutral Resist: "
            NPC(EditorIndex).Element_Neut_Res = scrlElement(Index).value
    End Select
    
    lblElement(Index).Caption = text & scrlElement(Index).value
    
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    NPC(EditorIndex).Sprite = scrlSprite.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRange.Caption = "Range: " & scrlRange.value
    NPC(EditorIndex).Range = scrlRange.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblNum.Caption = "Num: " & scrlNum.value

    If scrlNum.value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.value).Name)
        If scrlNum.value > 0 Then
            If Item(scrlNum.value).Type = ITEM_TYPE_CURRENCY Then
                chkRandCurrency.Enabled = True
            Else
                chkRandCurrency.value = vbUnchecked
                chkRandCurrency.Enabled = False
            End If
        End If
    End If
    
    NPC(EditorIndex).Drops(scrlDropIndex.value).DropItem = scrlNum.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).value
    NPC(EditorIndex).Stat(Index) = scrlStat(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblValue.Caption = "Value: " & scrlValue.value
    NPC(EditorIndex).Drops(scrlDropIndex.value).DropItemValue = scrlValue.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).AttackSay = txtAttackSay.text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtChance.text) > 0 Then Exit Sub
    If IsNumeric(txtChance.text) Then NPC(EditorIndex).Drops(scrlDropIndex.value).DropChance = Val(txtChance.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtChance_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtDamage.text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.text) Then NPC(EditorIndex).Damage = Val(txtDamage.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtEXP.text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.text) Then NPC(EditorIndex).EXP = Val(txtEXP.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtHP.text) > 0 Then Exit Sub
    If IsNumeric(txtHP.text) Then NPC(EditorIndex).HP = Val(txtHP.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHPMin_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtHPMin.text) > 0 Then Exit Sub
    If IsNumeric(txtHPMin.text) Then NPC(EditorIndex).HPMin = Val(txtHPMin.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtHPMin_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtLevel.text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.text) Then NPC(EditorIndex).Level = Val(txtLevel.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtSpawnSecs.text) > 0 Then Exit Sub
    NPC(EditorIndex).SpawnSecs = Val(txtSpawnSecs.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        NPC(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        NPC(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMoveSpeed_Change()
    lblMoveSpeed.Caption = "Movement Speed: " & scrlMoveSpeed.value
    NPC(EditorIndex).Speed = scrlMoveSpeed.value
End Sub

'ALATAR
Private Sub chkQuest_Click()
    NPC(EditorIndex).Quest = chkQuest.value
End Sub

Private Sub scrlQuest_Change()
    lblQuest = scrlQuest.value
    NPC(EditorIndex).QuestNum = scrlQuest.value
End Sub
'/ALATAR
Private Sub txtSpawnSecsMin_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Len(txtSpawnSecsMin.text) > 0 Then Exit Sub
    NPC(EditorIndex).SpawnSecsMin = Val(txtSpawnSecsMin.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtSpawnSecsMin_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
