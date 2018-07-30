VERSION 5.00
Begin VB.Form frmEditor_Animation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Editor"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
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
   ScaleHeight     =   6975
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5400
      TabIndex        =   31
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Animation Properties"
      Height          =   6255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   2415
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   28
         Top             =   3120
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   18
         Top             =   2520
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   16
         Top             =   1920
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Index           =   1
         Left            =   3360
         ScaleHeight     =   2475
         ScaleWidth      =   3075
         TabIndex        =   14
         Top             =   3600
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   13
         Top             =   1320
         Width           =   3135
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Index           =   0
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   3075
         TabIndex        =   7
         Top             =   3600
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLoopTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   27
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblLoopTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layer 1 (Above Player)"
         Height          =   180
         Left            =   3360
         TabIndex        =   19
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   17
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   12
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layer 0 (Below Player)"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Animation List"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5910
         ItemData        =   "frmEditor_Animation.frx":0000
         Left            =   120
         List            =   "frmEditor_Animation.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Animation(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Animation(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AnimationEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ClearAnimation EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AnimationEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AnimationEditorOk False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For i = 0 To 1
        scrlSprite(i).max = NumAnimations
        scrlLoopCount(i).max = 100
        scrlFrameCount(i).max = 100
        scrlLoopTime(i).max = 1000
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFrameCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFrameCount(Index).Caption = "Frame Count: " & scrlFrameCount(Index).value
    Animation(EditorIndex).Frames(Index) = scrlFrameCount(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlFrameCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFrameCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlFrameCount_Change Index
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlFrameCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLoopCount(Index).Caption = "Loop Count: " & scrlLoopCount(Index).value
    Animation(EditorIndex).LoopCount(Index) = scrlLoopCount(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLoopCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlLoopCount_Change Index
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLoopCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLoopTime(Index).Caption = "Loop Time: " & scrlLoopTime(Index).value
    Animation(EditorIndex).looptime(Index) = scrlLoopTime(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLoopTime_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlLoopTime_Change Index
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLoopTime_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSprite(Index).Caption = "Sprite: " & scrlSprite(Index).value
    Animation(EditorIndex).Sprite(Index) = scrlSprite(Index).value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSprite_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlSprite_Change Index
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
