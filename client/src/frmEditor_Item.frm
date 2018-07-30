VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmEditor_Item 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   848
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Projectiles"
      Height          =   3375
      Left            =   9720
      TabIndex        =   101
      Top             =   120
      Width           =   2895
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   2760
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectileSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   2040
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectileDamage 
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   1320
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblProjectileRange 
         BackStyle       =   0  'Transparent
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label lblProjectileSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label lblProjectileDamage 
         BackStyle       =   0  'Transparent
         Caption         =   "Damage: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblProjectilePic 
         BackStyle       =   0  'Transparent
         Caption         =   "Pic: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraBooks 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Book"
      Height          =   3375
      Left            =   3360
      TabIndex        =   99
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
      Begin RichTextLib.RichTextBox rtbBookText 
         Height          =   3015
         Left            =   120
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         MaxLength       =   1000
         TextRTF         =   $"frmEditor_Item.frx":3332
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requirements"
      Height          =   1335
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   6255
      Begin VB.ComboBox cmbSkill 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   960
         Width           =   2055
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   4080
         Max             =   255
         TabIndex        =   83
         Top             =   960
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   180
         Left            =   120
         TabIndex        =   84
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   390
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Level: 0"
         Height          =   180
         Index           =   6
         Left            =   2880
         TabIndex        =   82
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   990
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6120
      TabIndex        =   86
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtDesc 
         Height          =   735
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlCombatLvl 
         Height          =   255
         LargeChange     =   10
         Left            =   1680
         Max             =   100
         TabIndex        =   77
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox cmdCombatType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33AC
         Left            =   1320
         List            =   "frmEditor_Item.frx":33CB
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox chkStackable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Stackable"
         Height          =   255
         Left            =   2880
         TabIndex        =   75
         Top             =   3000
         Width           =   1335
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   73
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   71
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3436
         Left            =   3720
         List            =   "frmEditor_Item.frx":3438
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1680
         Width           =   2415
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":343A
         Left            =   4200
         List            =   "frmEditor_Item.frx":3447
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3470
         Left            =   120
         List            =   "frmEditor_Item.frx":3492
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblCombatLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Combat Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   80
         Top             =   3000
         Width           =   1245
      End
      Begin VB.Label lblCombatType 
         BackStyle       =   0  'Transparent
         Caption         =   "Combat Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   74
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   72
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   70
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Item List"
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7800
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpell 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   52
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   53
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraEquipment 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Equipment Data"
      Height          =   3375
      Left            =   3360
      TabIndex        =   32
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   11
         LargeChange     =   10
         Left            =   4440
         Max             =   255
         TabIndex        =   97
         Top             =   2640
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   10
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   95
         Top             =   2640
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   8
         LargeChange     =   10
         Left            =   4440
         Max             =   255
         TabIndex        =   93
         Top             =   2280
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   7
         LargeChange     =   10
         Left            =   4440
         Max             =   255
         TabIndex        =   91
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   89
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   9
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   87
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkHanded 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Two-Handed"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   3000
         Width           =   1335
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5640
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   2040
         Width           =   480
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   4200
         TabIndex        =   57
         Top             =   3000
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   40
         Top             =   840
         Value           =   100
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   39
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":34E0
         Left            =   1320
         List            =   "frmEditor_Item.frx":34F0
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neutral Resist: 0"
         Height          =   180
         Index           =   11
         Left            =   2880
         TabIndex        =   98
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neutral Dmg: 0"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   96
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Resist: 0"
         Height          =   180
         Index           =   8
         Left            =   2880
         TabIndex        =   94
         Top             =   2280
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Resist: 0"
         Height          =   180
         Index           =   7
         Left            =   2880
         TabIndex        =   92
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Dmg: 0"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   90
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Dmg: 0"
         Height          =   180
         Index           =   9
         Left            =   120
         TabIndex        =   88
         Top             =   2280
         UseMnemonic     =   0   'False
         Width           =   960
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   56
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   48
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   47
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   45
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   44
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraVitals 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkInstant 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   64
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   62
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   60
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   50
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   61
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add new features using Visual Basic 6.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   9840
      TabIndex        =   102
      Top             =   3720
      Width           =   2655
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub chkHanded_Click()
    Item(EditorIndex).Handed = chkHanded.value
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSkill_Click()
    Item(EditorIndex).SkillReq = cmbSkill.ListIndex + 1
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCombatType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).CombatTypeReq = cmdCombatType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCombatType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ItemEditorOk(False)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim I As Integer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlPic.max = numitems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    scrlDamage.max = MAX_INTEGER
    
    'set main txt for books
    rtbBookText.text = "Input book text here." & vbNewLine & _
    "Use /t to automatically tab your text 4 spaces."
    rtbBookText.MaxLength = TEXT_LENGTH
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        chkStackable.Visible = False
        Item(EditorIndex).Stackable = 0
        'scrlDamage_Change
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            Frame4.Visible = True
            Me.Width = 12825
        End If
    Else
        fraEquipment.Visible = False
        Frame4.Visible = False
        Me.Width = 9810
        chkStackable.Visible = True
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_WEAPON Or cmbType.ListIndex = ITEM_TYPE_SPELL Then
        cmdCombatType.Enabled = True
        scrlCombatLvl.Enabled = True
    Else
        cmdCombatType.Enabled = False
        scrlCombatLvl.Enabled = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_BOOK Then
        fraBooks.Visible = True
        rtbBookText.text = "The Book feature has not been fully implemented. The text will not display but the book and animations will."
    Else
        fraBooks.Visible = False
        rtbBookText.text = "Add new features using Visual Basic 6.0"
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkStackable_Click()
    Item(EditorIndex).Stackable = chkStackable.value
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub rtbBookText_Change()
Dim tLoc As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not frmEditor_Item.Visible = True Then Exit Sub
    
    tLoc = InStr(rtbBookText.text, "/t")
    If tLoc > 0 Then
        rtbBookText.text = Replace$(rtbBookText.text, "/t", "    ")
        rtbBookText.SelStart = tLoc + 3
    End If
    
    If Len(rtbBookText.text) > 0 Then
        Item(EditorIndex).Book.name = Item(EditorIndex).name
        Item(EditorIndex).Book.text = rtbBookText.text
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "rtbBookText_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.value
    Item(EditorIndex).AccessReq = scrlAccessReq.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.value
    Item(EditorIndex).AddHP = scrlAddHp.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.value
    Item(EditorIndex).AddMP = scrlAddMP.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.value
    Item(EditorIndex).AddEXP = scrlAddExp.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.value).name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCombatLvl_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblCombatLvl.Caption = "Combat Level: " & scrlCombatLvl
    Item(EditorIndex).CombatLvlReq = scrlCombatLvl.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCombatLvl_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.value
    Item(EditorIndex).Data2 = scrlDamage.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.value
    Item(EditorIndex).Pic = scrlPic.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.value
    Item(EditorIndex).Price = scrlPrice.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.value
    Item(EditorIndex).Rarity = scrlRarity.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
        Case 6
            text = "Light Dmg: "
            Item(EditorIndex).Element_Light_Dmg = scrlStatBonus(Index).value
        Case 7
            text = "Light Resist: "
            Item(EditorIndex).Element_Light_Res = scrlStatBonus(Index).value
        Case 8
            text = "Dark Resist: "
            Item(EditorIndex).Element_Dark_Res = scrlStatBonus(Index).value
        Case 9
            text = "Dark Dmg: "
            Item(EditorIndex).Element_Dark_Dmg = scrlStatBonus(Index).value
        Case 10
            text = "Neutral Dmg: "
            Item(EditorIndex).Element_Neut_Dmg = scrlStatBonus(Index).value
        Case 11
            text = "Neutral Resist: "
            Item(EditorIndex).Element_Neut_Res = scrlStatBonus(Index).value
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).value
    If Index < 6 Then Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
        Case 6
            text = "Skill Level: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.value).name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.value).name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.value
    
    Item(EditorIndex).data1 = scrlSpell.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileDamage.Caption = "Damage: " & scrlProjectileDamage.value
    Item(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.value
    Item(EditorIndex).ProjecTile.Range = scrlProjectileRange.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileSpeed.Caption = "Speed: " & scrlProjectileSpeed.value
    Item(EditorIndex).ProjecTile.speed = scrlProjectileSpeed.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
