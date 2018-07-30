Attribute VB_Name = "modQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10 'Alatar v1.2
Public Const EDITOR_TASKS As Byte = 7

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8
Public Const QUEST_TYPE_GOGETFROMEVENT As Byte = 9

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

Public Quest_Changed(1 To MAX_QUESTS) As Boolean

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

Public Type PlayerQuestRec
    status As Long
    ActualTask As Long
    CurrentCount As Long 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Long
    value As Long
End Type

Private Type QuestGiveItemRec
    Item As Long
    value As Long
End Type

Private Type QuestTakeItemRec
    Item As Long
    value As Long
End Type

Private Type QuestRewardItemRec
    Item As Long
    value As Long
End Type
'/Alatar v1.2

Public Type TaskRec
    order As Long
    NPC As Long
    Item As Long
    Map As Long
    Resource As Long
    Amount As Long
    Speech As String * 300
    TaskLog As String * 300
    QuestEnd As Boolean
    Event As Long
End Type

Public Type QuestRec
    'Alatar v1.2
    name As String * 30
    Repeat As Long
    QuestLog As String * 300
    Speech(1 To 3) As String * 300
    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec
    
    RequiredLevel As Long
    RequiredQuest As Long
    RequiredClass(1 To 5) As Long
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec
    
    RewardExp As Long
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec
    
    Task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2
    
    '/escfoe2 :p
    Skill As Long
    SkillExp As Long
 
End Type

' ////////////
' // Editor //
' ////////////

Public Sub QuestEditorInit()
Dim I As Long
    
    If frmEditor_Quest.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1
    If Options.Debug Then On Error GoTo ErrHandler

    With frmEditor_Quest
        'Alatar v1.2
        .txtName = Trim$(Quest(EditorIndex).name)
        If Quest(EditorIndex).Repeat = 1 Then
            .chkRepeat.value = 1
        Else
            .chkRepeat.value = 0
        End If
        .txtQuestLog = Trim$(Quest(EditorIndex).QuestLog)
        For I = 1 To 3
            .txtSpeech(I).text = Trim$(Quest(EditorIndex).Speech(I))
        Next
        
        .scrlReqLevel.value = Quest(EditorIndex).RequiredLevel
        .scrlReqQuest.value = Quest(EditorIndex).RequiredQuest
        For I = 1 To 5
            .scrlReqClass.value = Quest(EditorIndex).RequiredClass(I)
        Next
        
        .scrlExp.value = Quest(EditorIndex).RewardExp
        If Quest(EditorIndex).Skill > 0 Then
            .cmbSkill.ListIndex = Quest(EditorIndex).Skill - 1
        Else
            .cmbSkill.ListIndex = 0
        End If
        .scrlSkillExp.value = Quest(EditorIndex).SkillExp
        
        'Update the lists
        UpdateQuestGiveItems
        UpdateQuestTakeItems
        UpdateQuestRewardItems
        UpdateQuestRequirementItems
        UpdateQuestClass
        
        '/Alatar v1.2
        
        'load task nº1
        .scrlTotalTasks.value = 1
        LoadTask EditorIndex, 1
        
    End With

    Quest_Changed(EditorIndex) = True
    Exit Sub
    
ErrHandler:
    Call HandleError("QuestEditorInit", "modQuests", Err.Number, Err.Description, Err.Source, Err.HelpContext)
End Sub

'Alatar v1.2
Public Sub UpdateQuestGiveItems()
    Dim I As Long
    
    frmEditor_Quest.lstGiveItem.Clear
    
    For I = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).GiveItem(I)
            If .Item = 0 Then
                frmEditor_Quest.lstGiveItem.AddItem "-"
            Else
                frmEditor_Quest.lstGiveItem.AddItem Trim$(Trim$(Item(.Item).name) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestTakeItems()
    Dim I As Long
    
    frmEditor_Quest.lstTakeItem.Clear
    
    For I = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).TakeItem(I)
            If .Item = 0 Then
                frmEditor_Quest.lstTakeItem.AddItem "-"
            Else
                frmEditor_Quest.lstTakeItem.AddItem Trim$(Trim$(Item(.Item).name) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRewardItems()
    Dim I As Long
    
    frmEditor_Quest.lstItemRew.Clear
    
    For I = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RewardItem(I)
            If .Item = 0 Then
                frmEditor_Quest.lstItemRew.AddItem "-"
            Else
                frmEditor_Quest.lstItemRew.AddItem Trim$(Trim$(Item(.Item).name) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRequirementItems()
    Dim I As Long
    
    frmEditor_Quest.lstReqItem.Clear
    
    For I = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RequiredItem(I)
            If .Item = 0 Then
                frmEditor_Quest.lstReqItem.AddItem "-"
            Else
                frmEditor_Quest.lstReqItem.AddItem Trim$(Trim$(Item(.Item).name) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestClass()
    Dim I As Long
    
    frmEditor_Quest.lstReqClass.Clear
    
    For I = 1 To 5
        If Quest(EditorIndex).RequiredClass(I) = 0 Then
            frmEditor_Quest.lstReqClass.AddItem "-"
        Else
            frmEditor_Quest.lstReqClass.AddItem Trim$(Trim$(Class(Quest(EditorIndex).RequiredClass(I)).name))
        End If
    Next
End Sub
'/Alatar v1.2

Public Sub QuestEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    For I = 1 To MAX_QUESTS
        If Quest_Changed(I) Then
            Call SendSaveQuest(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
End Sub

Public Sub QuestEditorCancel()
    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests
End Sub

Public Sub ClearChanged_Quest()
    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).name = vbNullString
End Sub

Sub ClearQuests()
Dim I As Long

    For I = 1 To MAX_QUESTS
        Call ClearQuest(I)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Public Sub SendRequestEditQuest()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditQuest
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

    Set buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    buffer.WriteLong CSaveQuest
    buffer.WriteLong QuestNum
    buffer.WriteBytes QuestData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Sub SendRequestQuests()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestQuests
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub UpdateQuestLog()
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CQuestLogUpdate
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub UpdateFriendsList()
    Dim buffer As clsBuffer
    
    ' Clear the list ahead of time just to be sure.
    frmMain.lstFriends.Clear
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUpdateFList
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub PlayerHandleQuest(ByVal QuestNum As Long, ByVal order As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong CPlayerHandleQuest
    buffer.WriteLong QuestNum
    buffer.WriteLong order '1=accept quest, 2=cancel quest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

' ///////////////
' // Functions //
' ///////////////

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_STARTED Then 'Status=1 means started
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED Or Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim I As Long
    GetQuestNum = 0
    
    For I = 1 To MAX_QUESTS
        If Trim$(Quest(I).name) = Trim$(QuestName) Then
            GetQuestNum = I
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

'Subroutine that load the desired task in the form
Public Sub LoadTask(ByVal QuestNum As Long, ByVal TaskNum As Long)
    Dim TaskToLoad As TaskRec
    TaskToLoad = Quest(QuestNum).Task(TaskNum)
    
    With frmEditor_Quest
        'Load the task type
        .optTask(TaskToLoad.order).value = True
        'Load textboxes
        .txtTaskSpeech.text = vbNullString
        .txtTaskLog.text = vbNullString & Trim$(TaskToLoad.TaskLog)
        'Set scrolls to 0 and disable them so they can be enabled when needed
        .scrlNPC.value = 0
        .scrlItem.value = 0
        .scrlMap.value = 0
        .scrlResource.value = 0
        .scrlAmount.value = 0
        .scrlEvent.value = 0
        .txtTaskSpeech.Enabled = False
        .scrlNPC.Enabled = False
        .scrlItem.Enabled = False
        .scrlMap.Enabled = False
        .scrlResource.Enabled = False
        .scrlAmount.Enabled = False
        .scrlEvent.Enabled = False
        
        If TaskToLoad.QuestEnd = True Then
            .chkEnd.value = 1
        Else
            .chkEnd.value = 0
        End If
        
        Select Case TaskToLoad.order
            Case 0 'Nothing
                
            Case QUEST_TYPE_GOSLAY '1
                .scrlNPC.Enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOGATHER '2
                .scrlItem.Enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOTALK '3
                .scrlNPC.Enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
                
            Case QUEST_TYPE_GOREACH '4
                .scrlMap.Enabled = True
                .scrlMap.value = TaskToLoad.Map
            
            Case QUEST_TYPE_GOGIVE '5
                .scrlItem.Enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
                .scrlNPC.Enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
            
            Case QUEST_TYPE_GOKILL '6
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOTRAIN '7
                .scrlResource.Enabled = True
                .scrlResource.value = TaskToLoad.Resource
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
            
            Case QUEST_TYPE_GOGET '8
                .scrlNPC.Enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .scrlItem.Enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
                
            Case QUEST_TYPE_GOGETFROMEVENT '9
                .scrlEvent.Enabled = True
                .scrlEvent.value = TaskToLoad.Event
                .scrlItem.Enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.Enabled = True
                .scrlAmount.value = TaskToLoad.Amount
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
                
        End Select
    End With
End Sub

Public Sub RefreshQuestLog()
    Dim I As Long
    
    frmMain.lstQuestLog.Clear
    For I = 1 To MAX_QUESTS
        If QuestInProgress(I) Then
            frmMain.lstQuestLog.AddItem Trim$(Quest(I).name)
        End If
    Next
    
End Sub

' ////////////////////////
' // Visual Interaction //
' ////////////////////////

Public Sub LoadQuestlogBox(ByVal ButtonPressed As Integer)
    Dim QuestNum As Long, I As Long
    Dim QuestSayMessage As String
    Dim ExpType As String
    
    With frmMain
        If Trim$(.lstQuestLog.text) = vbNullString Then Exit Sub
        
        QuestNum = GetQuestNum(Trim$(.lstQuestLog.text))
        
        Select Case ButtonPressed
            Case 1 'Actual Task
                QuestExtraVisible = False
                QuestSubtitle = "Task: [" + Trim$(Player(MyIndex).PlayerQuest(QuestNum).ActualTask) + "]"
                If QuestCompleted(QuestNum) = False Then
                    ' It's not trimming the text?? Maybe these aren't spaces idk
                    QuestSay = Trim$(Quest(QuestNum).Task(Player(MyIndex).PlayerQuest(QuestNum).ActualTask).TaskLog)
                Else
                    QuestSay = "."
                End If
                
            Case 2 'Last Speech
                QuestExtraVisible = False
                QuestSubtitle = "Last Speech:"
                If Player(MyIndex).PlayerQuest(QuestNum).ActualTask > 1 Then
                    QuestSay = Trim$(Quest(QuestNum).Task(Player(MyIndex).PlayerQuest(QuestNum).ActualTask - 1).Speech)
                    If QuestSay = "" Then
                        QuestSay = Trim$(Quest(QuestNum).Speech(1))
                    End If
                Else
                    QuestSay = Trim$(Quest(QuestNum).Speech(1))
                End If
            
            Case 3 'Quest Status
                QuestSubtitle = "Quest Status:"
                If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_STARTED Then
                    QuestSay = "Quest in Progress." & vbNewLine & "Step: " & Player(MyIndex).PlayerQuest(QuestNum).ActualTask & "."
                    QuestExtra = "Cancel Quest"
                    QuestExtraVisible = True
                ElseIf QuestCompleted(QuestNum) Then
                    QuestSay = "Completed"
                    QuestExtraVisible = False
                End If
                
            Case 4 'Quest Log (Main Task)
                QuestExtraVisible = False
                QuestSubtitle = "Main Task:"
                QuestSay = Trim$(Quest(QuestNum).QuestLog)
                
            Case 5 'Requirements
                QuestExtraVisible = False
                QuestSubtitle = "Requirements"
                QuestSayMessage = "Level: "
                If Quest(QuestNum).RequiredLevel > 0 Then
                    QuestSayMessage = QuestSayMessage & "" & Quest(QuestNum).RequiredLevel & vbNewLine & "Quest: "
                Else
                    QuestSayMessage = QuestSayMessage & " None." & vbNewLine & "Quest: "
                End If
                If Quest(QuestNum).RequiredQuest > 0 Then
                    QuestSayMessage = QuestSayMessage & "" & Trim$(Quest(Quest(QuestNum).RequiredQuest).name) & vbNewLine & "Race: "
                Else
                    QuestSayMessage = QuestSayMessage & " None." & vbNewLine & "Race: "
                End If
                For I = 1 To 5
                    If Quest(QuestNum).RequiredClass(I) > 0 Then
                        QuestSayMessage = QuestSayMessage & Trim$(Class(Quest(QuestNum).RequiredClass(I)).name) & ". "
                    End If
                Next
                QuestSayMessage = QuestSayMessage & vbNewLine & "Items:"
                For I = 1 To MAX_QUESTS_ITEMS
                    If Quest(QuestNum).RequiredItem(I).Item > 0 Then
                        QuestSayMessage = QuestSayMessage & " " & Trim$(Item(Quest(QuestNum).RequiredItem(I).Item).name) & "(" & Trim$(Quest(QuestNum).RequiredItem(I).value) & ")"
                    End If
                Next
                QuestSay = QuestSayMessage
            
            Case 6 'Rewards
                QuestExtraVisible = False
                QuestSubtitle = "Rewards"
                QuestSayMessage = "EXP: " & Quest(QuestNum).RewardExp & vbNewLine & "Items:"
                For I = 1 To MAX_QUESTS_ITEMS
                    If Quest(QuestNum).RewardItem(I).Item > 0 Then
                        QuestSayMessage = QuestSayMessage & " " & Trim$(Item(Quest(QuestNum).RewardItem(I).Item).name) & "(" & Trim$(Quest(QuestNum).RewardItem(I).value) & ")"
                    End If
                Next
                QuestSay = QuestSayMessage
            
            Case Else
                Exit Sub
        End Select
        
        QuestName = "Name: " & Trim$(Quest(QuestNum).name)
        inChat = True
        GUIWindow(GUI_QUESTDIALOGUE).Visible = True
        
    End With
End Sub

Public Sub RunQuestDialogueExtraLabel()
    If QuestExtra = "Cancel Quest" Then
        PlayerHandleQuest GetQuestNum(Trim$(QuestName)), 2
        QuestExtra = "Extra"
        QuestExtraVisible = False
        GUIWindow(GUI_QUESTLOG).Visible = False
        inChat = False
        GUIWindow(GUI_QUESTDIALOGUE).Visible = False
        frmMain.lstQuestLog.Visible = False
        RefreshQuestLog
    End If
End Sub

