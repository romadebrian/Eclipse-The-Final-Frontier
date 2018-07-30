VERSION 5.00
Begin VB.Form frmEditor_Shop 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
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
   Icon            =   "frmEditor_Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Shop Properties"
      Height          =   4455
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   2040
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   19
         Top             =   840
         Value           =   100
         Width           =   5055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   1860
         ItemData        =   "frmEditor_Shop.frx":3332
         Left            =   120
         List            =   "frmEditor_Shop.frx":334E
         TabIndex        =   11
         Top             =   2400
         Width           =   5055
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   13
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Shop List"
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   4020
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ShopEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk(False)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long
Dim tmpPos As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    tmpPos = lstTradeItem.ListIndex
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = cmbItem.ListIndex
        .ItemValue = Val(txtItemValue.text)
        .CostItem = cmbCostItem.ListIndex
        .CostValue = Val(txtCostValue.text)
    End With
    UpdateShopTrade tmpPos
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdUpdate_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDeleteTrade_Click()
Dim Index As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = 0
        .ItemValue = 0
        .CostItem = 0
        .CostValue = 0
    End With
    Call UpdateShopTrade
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDeleteTrade_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ShopEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBuy_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblBuy.Caption = "Buy Rate: " & scrlBuy.value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlBuy_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
