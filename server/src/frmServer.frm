VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   503
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCpsLock"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCPS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtChat"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtText"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraServer"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraDatabase"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Guilds"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "cmdGSave"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Kill Event"
      TabPicture(4)   =   "frmServer.frx":170FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTab2"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Control"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   91
         Top             =   3240
         Width           =   7455
         Begin VB.CheckBox chkGUIBars 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Original GUI Bars? (NO)"
            Height          =   255
            Left            =   3960
            TabIndex        =   98
            Top             =   360
            Width           =   3375
         End
         Begin VB.CheckBox chkProj 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Allow Projectiles? (NO)"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   1440
            Width           =   4455
         End
         Begin VB.CheckBox chkFS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Allow FullScreen? (NO)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   1080
            Width           =   4455
         End
         Begin VB.CheckBox chkDropInvItems 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Drop Inv Items On Death (Inactive)"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   720
            Width           =   4455
         End
         Begin VB.CheckBox chkFriendSystem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Friends System (InActive)"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   360
            Width           =   3135
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8281
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         BackColor       =   14737632
         TabCaption(0)   =   "Settings"
         TabPicture(0)   =   "frmServer.frx":17116
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(2)=   "Frame6"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "More Settings"
         TabPicture(1)   =   "frmServer.frx":17132
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtPlayerMsg"
         Tab(1).Control(1)=   "cboColor_PlayerMsg"
         Tab(1).Control(2)=   "chkPlayerMsg"
         Tab(1).Control(3)=   "cboColor_ActionMsg"
         Tab(1).Control(4)=   "opRise"
         Tab(1).Control(5)=   "opStat"
         Tab(1).Control(6)=   "txtActionMsg"
         Tab(1).Control(7)=   "chkActionMsg"
         Tab(1).Control(8)=   "Label14"
         Tab(1).Control(9)=   "Label13"
         Tab(1).Control(10)=   "Line1"
         Tab(1).Control(11)=   "Label12"
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "Status"
         TabPicture(2)   =   "frmServer.frx":1714E
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label9"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label8"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "lbl1"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "lbl2"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "lbl3"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "lblFirst"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "lblSecond"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "lblThird"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "lblFourth"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "lblFifth"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "Picture1"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "tmrGetTime"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).ControlCount=   12
         Begin VB.Timer tmrGetTime 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   3000
            Top             =   720
         End
         Begin VB.TextBox txtPlayerMsg 
            Height          =   1095
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Text            =   "frmServer.frx":1716A
            Top             =   2640
            Width           =   6975
         End
         Begin VB.ComboBox cboColor_PlayerMsg 
            Height          =   315
            ItemData        =   "frmServer.frx":17196
            Left            =   -74160
            List            =   "frmServer.frx":171D0
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   2160
            Width           =   2775
         End
         Begin VB.CheckBox chkPlayerMsg 
            Caption         =   "Send Player Message every kill."
            Height          =   255
            Left            =   -74760
            TabIndex        =   83
            Top             =   1800
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.ComboBox cboColor_ActionMsg 
            Height          =   315
            ItemData        =   "frmServer.frx":17270
            Left            =   -74280
            List            =   "frmServer.frx":172AA
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   1200
            Width           =   2775
         End
         Begin VB.OptionButton opRise 
            Caption         =   "Rising"
            Height          =   255
            Left            =   -69600
            TabIndex        =   81
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton opStat 
            Caption         =   "Stationary"
            Height          =   255
            Left            =   -71280
            TabIndex        =   80
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox txtActionMsg 
            Height          =   285
            Left            =   -74280
            MaxLength       =   30
            TabIndex        =   78
            Text            =   "#placement# Place"
            Top             =   840
            Width           =   3735
         End
         Begin VB.CheckBox chkActionMsg 
            Caption         =   "Send Action Message every kill."
            Height          =   255
            Left            =   -74760
            TabIndex        =   77
            Top             =   480
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00000000&
            Height          =   3255
            Left            =   3600
            ScaleHeight     =   3195
            ScaleWidth      =   3555
            TabIndex        =   61
            Top             =   960
            Width           =   3615
            Begin VB.CommandButton btnStart 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Activate TopKill Event"
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   2640
               Width           =   3255
            End
            Begin VB.Label lblTime 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Run Time: 00:00:00"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFC0&
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   2280
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Label lblStatus 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "??"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   2160
               TabIndex        =   65
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Game Won:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   375
               Left            =   240
               TabIndex        =   66
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lblWinner 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Waiting..."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   120
               TabIndex        =   64
               Top             =   480
               Width           =   3255
            End
            Begin VB.Label lblTotal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Total Kills: 0"
               ForeColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   1200
               Width           =   3255
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               X1              =   120
               X2              =   3360
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Shape shpColor 
               BackColor       =   &H000000FF&
               BackStyle       =   1  'Opaque
               FillColor       =   &H000000FF&
               Height          =   615
               Left            =   2160
               Shape           =   3  'Circle
               Top             =   1560
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Options"
            Height          =   4215
            Left            =   -71280
            TabIndex        =   54
            Top             =   360
            Width           =   3495
            Begin VB.ComboBox cboColor_End 
               Height          =   315
               ItemData        =   "frmServer.frx":1734A
               Left            =   720
               List            =   "frmServer.frx":17384
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   3840
               Width           =   2655
            End
            Begin VB.ComboBox cboColor_Start 
               Height          =   315
               ItemData        =   "frmServer.frx":17424
               Left            =   720
               List            =   "frmServer.frx":1745E
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   1800
               Width           =   2655
            End
            Begin VB.TextBox txtEndMsg 
               Height          =   1335
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   58
               Text            =   "frmServer.frx":174FE
               Top             =   2520
               Width           =   3255
            End
            Begin VB.TextBox txtStartMsg 
               Height          =   1335
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   56
               Text            =   "frmServer.frx":1764D
               Top             =   480
               Width           =   3255
            End
            Begin VB.Line Line3 
               X1              =   120
               X2              =   3360
               Y1              =   2160
               Y2              =   2160
            End
            Begin VB.Label Label16 
               Caption         =   "Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   3840
               Width           =   735
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "END Message:"
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   2280
               Width           =   3255
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "START Message:"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Experience"
            Height          =   3255
            Left            =   -74760
            TabIndex        =   40
            Top             =   1320
            Width           =   3255
            Begin VB.CheckBox chk5 
               Caption         =   "Use 5th?"
               Height          =   315
               Left            =   2040
               TabIndex        =   48
               Top             =   2570
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox chk4 
               Caption         =   "Use 4th?"
               Height          =   315
               Left            =   2040
               TabIndex        =   47
               Top             =   1970
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox chk3 
               Caption         =   "Use 3rd?"
               Height          =   315
               Left            =   2040
               TabIndex        =   46
               Top             =   1370
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.HScrollBar scrlFifth 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               TabIndex        =   45
               Top             =   2880
               Width           =   3015
            End
            Begin VB.HScrollBar scrlFourth 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               TabIndex        =   44
               Top             =   2280
               Width           =   3015
            End
            Begin VB.HScrollBar scrlThird 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               TabIndex        =   43
               Top             =   1680
               Width           =   3015
            End
            Begin VB.HScrollBar scrlSecond 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               TabIndex        =   42
               Top             =   1080
               Width           =   3015
            End
            Begin VB.HScrollBar scrlFirst 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   3015
            End
            Begin VB.Label lblFifthExp 
               Caption         =   "Fifth Place: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   2640
               Width           =   3015
            End
            Begin VB.Label lblFourthExp 
               Caption         =   "Fourth Place: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   2040
               Width           =   3015
            End
            Begin VB.Label lblThirdExp 
               Caption         =   "Third Place: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   1440
               Width           =   3015
            End
            Begin VB.Label lblSecondExp 
               Caption         =   "Second Place: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   840
               Width           =   3015
            End
            Begin VB.Label lblFirstExp 
               Caption         =   "First Place: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   3015
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Kills"
            Height          =   855
            Left            =   -74760
            TabIndex        =   37
            Top             =   360
            Width           =   3255
            Begin VB.HScrollBar scrlNeeded 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               TabIndex        =   38
               Top             =   480
               Width           =   3015
            End
            Begin VB.Label lblNeeded 
               Caption         =   "Kills Needed: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   3015
            End
         End
         Begin VB.Label Label14 
            Caption         =   "Color:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   87
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Color:"
            Height          =   255
            Left            =   -74760
            TabIndex        =   86
            Top             =   2160
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   -74880
            X2              =   -67800
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label12 
            Caption         =   "Msg:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   79
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblFifth 
            Alignment       =   2  'Center
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   240
            TabIndex        =   76
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label lblFourth 
            Alignment       =   2  'Center
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Left            =   240
            TabIndex        =   75
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label lblThird 
            Alignment       =   2  'Center
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   375
            Left            =   240
            TabIndex        =   74
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label lblSecond 
            Alignment       =   2  'Center
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   73
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label lblFirst 
            Alignment       =   2  'Center
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label lbl3 
            Caption         =   "Fifth Place:"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label lbl2 
            Caption         =   "Fourth Place:"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lbl1 
            Caption         =   "Third Place:"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Second Place:"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "First Place:"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Awesome Sh*t"
         Height          =   1215
         Left            =   -71880
         TabIndex        =   34
         Top             =   2040
         Width           =   1815
         Begin VB.CommandButton btnDubExp 
            Caption         =   "  Activate    Double Exp"
            Height          =   615
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtText 
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   915
         Width           =   7455
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   4635
         Width           =   7455
      End
      Begin VB.Frame fraDatabase 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reload"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   2895
         Begin VB.CommandButton cmdReloadQuests 
            Caption         =   "Quests"
            Height          =   375
            Left            =   1440
            TabIndex        =   95
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadCombos 
            Caption         =   "Combos"
            Height          =   375
            Left            =   1440
            TabIndex        =   94
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   375
            Left            =   1440
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   375
            Left            =   1440
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame fraServer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Server"
         Height          =   1575
         Left            =   -71880
         TabIndex        =   16
         Top             =   480
         Width           =   1815
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkServerLog 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdGSave 
         Caption         =   "Save Config"
         Height          =   255
         Left            =   -69960
         TabIndex        =   15
         Top             =   3195
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Join config"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   8
         Top             =   1920
         Width           =   6015
         Begin VB.TextBox txtGJoinItem 
            Height          =   285
            Left            =   960
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtGJoinLvl 
            Height          =   285
            Left            =   3960
            TabIndex        =   10
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtGJoinCost 
            Height          =   285
            Left            =   960
            TabIndex        =   9
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Item:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Level Req:"
            Height          =   255
            Left            =   2880
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Purchase config"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   6015
         Begin VB.TextBox txtGBuyItem 
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtGBuyLvl 
            Height          =   285
            Left            =   3960
            TabIndex        =   3
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtGBuyCost 
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Item:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Level Req:"
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value:"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   555
         End
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7858
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblCPS 
         Caption         =   "CPS: 0"
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuModPlayer 
         Caption         =   "Make Monitor"
      End
      Begin VB.Menu mnuMapPlayer 
         Caption         =   "Make Mapper"
      End
      Begin VB.Menu mnuDevPlayer 
         Caption         =   "Make Developer"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Player(1).Switches(1) = 1
End Sub

Private Sub Command2_Click()
    Player(1).Switches(1) = 0
End Sub

Private Sub btnDubExp_Click()
    DoubleExp = Not DoubleExp
    If DoubleExp Then
        Call GlobalMsg("Server: DOUBLE EXP has been activated. Enjoy.", Green)
        Call TextAdd("DOUBLE EXP activated.")
        btnDubExp.Caption = "Deactivate Double Exp"
    Else
        Call GlobalMsg("Server: DOUBLE EXP has been deactivated.", Green)
        Call TextAdd("DOUBLE EXP deactivated.")
        btnDubExp.Caption = "  Activate    Double Exp"
    End If
End Sub

Private Sub btnStart_Click()
    If scrlNeeded.Value < 1 Then
        SSTab2.Tab = 0
        lblNeeded.ForeColor = vbRed
        Exit Sub
    End If
    Call InitTopKillEvent
End Sub

Private Sub chk3_Click()
    scrlThird.Enabled = CBool(chk3.Value)
    If chk3.Value = vbUnchecked Then
        scrlThird.Value = 1
        chk4.Value = vbUnchecked
        chk5.Value = vbUnchecked
        chk4.Enabled = False
        chk5.Enabled = False
    Else
        chk4.Enabled = True
        chk5.Enabled = True
    End If
End Sub

Private Sub chk4_Click()
    scrlFourth.Enabled = CBool(chk4.Value)
    If chk4.Value = vbUnchecked Then
        scrlFourth.Value = 1
        chk5.Value = vbUnchecked
        chk5.Enabled = False
    Else
        chk5.Enabled = True
    End If
End Sub

Private Sub chk5_Click()
    scrlFifth.Enabled = CBool(chk5.Value)
    If chk5.Value = vbUnchecked Then scrlFifth.Value = 1
End Sub

Private Sub chkDropInvItems_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkDropInvItems.Value Then
        chkDropInvItems.Caption = "Drop Inv Items On Death (Active)"
    Else
        chkDropInvItems.Caption = "Drop Inv Items On Death (Inactive)"
    End If
    
    Call PutVar(Path, "OPTIONS", "DropOnDeath", CStr(chkDropInvItems.Value))
End Sub

Private Sub chkFriendSystem_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkFriendSystem.Value Then
        chkFriendSystem.Caption = "Friends System (Active)"
    Else
        chkFriendSystem.Caption = "Friends System (Inactive)"
    End If
    
    Call PutVar(Path, "OPTIONS", "FriendSystem", CStr(chkFriendSystem.Value))
End Sub

Private Sub chkFS_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkFS.Value Then
        chkFS.Caption = "Allow FullScreen? (YES)"
    Else
        chkFS.Caption = "Allow FullScreen? (NO)"
    End If
    
    Call PutVar(Path, "OPTIONS", "FullScreen", CStr(chkFS.Value))
    Options.FullScreen = chkFS.Value
    
    SendHighIndex
End Sub

Private Sub chkGUIBars_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkGUIBars.Value Then
        chkGUIBars.Caption = "Original GUI Bars? (YES)"
    Else
        chkGUIBars.Caption = "Original GUI Bars? (NO)"
    End If
    
    Call PutVar(Path, "OPTIONS", "OriginalGUIBars", CStr(chkGUIBars.Value))
    Options.OriginalGUIBars = chkGUIBars.Value
    
    SendGUIBarsToAll
End Sub

Private Sub chkProj_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkProj.Value Then
        chkProj.Caption = "Allow Projectiles? (YES)"
    Else
        chkProj.Caption = "Allow Projectiles? (NO)"
    End If
    
    Call PutVar(Path, "OPTIONS", "Projectiles", CStr(chkProj.Value))
    Options.Projectiles = chkProj.Value
End Sub

Private Sub cmdReloadCombos_Click()
Dim I As Long
    Call LoadCombos
    Call TextAdd("All combos reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendCombos I
        End If
    Next
End Sub

Private Sub cmdReloadQuests_Click()
Dim I As Long
    Call LoadQuests
    Call TextAdd("All quests reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendQuests I
        End If
    Next
End Sub

Private Sub scrlFifth_Change()
    lblFifthExp.Caption = "Fifth Place: " & scrlFifth.Value
End Sub

Private Sub scrlFirst_Change()
    lblFirstExp.Caption = "First Place: " & scrlFirst.Value
End Sub

Private Sub scrlFourth_Change()
    lblFourthExp.Caption = "Fourth Place: " & scrlFourth.Value
End Sub

Private Sub scrlNeeded_Change()
    lblNeeded.Caption = "Kills Needed: " & scrlNeeded.Value
    lblNeeded.ForeColor = vbBlack
End Sub

Private Sub scrlSecond_Change()
    lblSecondExp.Caption = "Second Place: " & scrlSecond.Value
End Sub

Private Sub scrlThird_Change()
    lblThirdExp.Caption = "Third Place: " & scrlThird.Value
End Sub

Private Sub cmdGSave_Click()
    Options.Buy_Cost = frmServer.txtGBuyCost.Text
    Options.Buy_Lvl = frmServer.txtGBuyLvl.Text
    Options.Buy_Item = frmServer.txtGBuyItem.Text
    Options.Join_Cost = frmServer.txtGJoinCost.Text
    Options.Join_Lvl = frmServer.txtGJoinLvl.Text
    Options.Join_Item = frmServer.txtGJoinItem.Text
    SaveOptions
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim I As Long
    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendClasses I
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
Dim I As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendItems I
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim I As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            PlayerWarp I, GetPlayerMap(I), GetPlayerX(I), GetPlayerY(I)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim I As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendNpcs I
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim I As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendShops I
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim I As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendSpells I
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim I As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendResources I
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim I As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendAnimations I
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call SetData
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub tmrGetTime_Timer()
Dim Display_S As String, Display_M As String, Display_H As String
    'get the time variables
    If Time_Seconds + 1 = 60 Then
        If Time_Minutes + 1 = 60 Then
            Time_Hours = Time_Hours + 1
            Time_Minutes = 0
        Else
            Time_Minutes = Time_Minutes + 1
            Time_Seconds = 0
        End If
    Else
        Time_Seconds = Time_Seconds + 1
    End If
    
    'prepare them
    Display_S = Time_Seconds
    Display_M = Time_Minutes
    Display_H = Time_Hours
    If Time_Seconds < 10 Then Display_S = "0" & Time_Seconds
    If Time_Minutes < 10 Then Display_M = "0" & Time_Minutes
    If Time_Hours < 10 Then Display_H = "0" & Time_Hours
    
    'show them
    lblTime.Caption = "Run Time:  " & Display_H & ":" & Display_M & ":" & Display_S
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (I)

        If I < 10 Then
            frmServer.lvwInfo.ListItems(I).Text = "00" & I
        ElseIf I < 100 Then
            frmServer.lvwInfo.ListItems(I).Text = "0" & I
        Else
            frmServer.lvwInfo.ListItems(I).Text = I
        End If

        frmServer.lvwInfo.ListItems(I).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(I).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(I).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Sub mnuModPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 1)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted monitor access.", BrightCyan)
    End If

End Sub

Sub mnuMapPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 2)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted mapper access.", BrightCyan)
    End If

End Sub

Sub mnuDevPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 3)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted developer access.", BrightCyan)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub
