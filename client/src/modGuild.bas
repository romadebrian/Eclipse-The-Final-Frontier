Attribute VB_Name = "modGuild"
Public Const MAX_GUILD_MEMBERS As Long = 50
Public Const MAX_GUILD_RANKS As Long = 6
Public Const MAX_GUILD_RANKS_PERMISSION As Long = 6


Public GuildData As GuildRec

Public Type GuildRanksRec
    'General variables
    Used As Boolean
    name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
    RankPermissionName(1 To MAX_GUILD_RANKS_PERMISSION) As String
End Type

Public Type GuildMemberRec
    'User login/name
    Used As Boolean
    
    User_Login As String
    User_Name As String
    Founder As Boolean
    
    Online As Boolean
    
    'Guild Variables
    Rank As Integer
    Comment As String * 300
     
End Type

Public Type GuildRec
    In_Use As Boolean
    
    Guild_Name As String
    Guild_Tag As String
    
    'Guild file number for saving
    Guild_Fileid As Long
    
    Guild_Members(1 To MAX_GUILD_MEMBERS) As GuildMemberRec
    Guild_Ranks(1 To MAX_GUILD_RANKS) As GuildRanksRec
    
    'Message of the day
    Guild_MOTD As String * 100
    
    'The rank recruits start at
    Guild_RecruitRank As Integer
    Guild_Color As Integer

End Type

Public Sub HandleAdminGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Integer
Dim B As Integer

    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    If buffer.ReadByte = 1 Then
        frmGuildAdmin.Visible = True
    Else
        frmGuildAdmin.Visible = False
    End If

    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAdminGuild", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub HandleSendGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Integer
Dim B As Integer

    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    GuildData.Guild_Name = buffer.ReadString
    GuildData.Guild_Tag = buffer.ReadString
    GuildData.Guild_Color = buffer.ReadInteger
    GuildData.Guild_MOTD = buffer.ReadString
    GuildData.Guild_RecruitRank = buffer.ReadInteger
    
    'Get Members
    For I = 1 To MAX_GUILD_MEMBERS
        GuildData.Guild_Members(I).User_Name = buffer.ReadString
        GuildData.Guild_Members(I).Rank = buffer.ReadInteger
        GuildData.Guild_Members(I).Comment = buffer.ReadString
        GuildData.Guild_Members(I).Online = buffer.ReadByte
    Next I
    
    'Get Ranks
    For I = 1 To MAX_GUILD_RANKS
        GuildData.Guild_Ranks(I).name = buffer.ReadString
        For B = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData.Guild_Ranks(I).RankPermission(B) = buffer.ReadByte
            GuildData.Guild_Ranks(I).RankPermissionName(B) = buffer.ReadString
        Next B
    Next I
    
    'Update Guildadmin data
    Call frmGuildAdmin.Load_Guild_Admin
    
    ' Reset Players Guild Tag/Name
    Player(MyIndex).GuildName = GuildData.Guild_Name
    Player(MyIndex).GuildTag = GuildData.Guild_Tag
    Player(MyIndex).GuildColor = GuildData.Guild_Color
    
    Set buffer = Nothing
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSendGuild", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Public Sub GuildMsg(ByVal text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSayGuild
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GuildMake(ByVal SendText As String, ByVal SendText2 As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildCommand
    buffer.WriteInteger 1
    buffer.WriteString SendText
    buffer.WriteString SendText2
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GuildCommand(ByVal Command As Integer, ByVal SendText As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildCommand
    buffer.WriteInteger Command
    buffer.WriteString SendText
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GuildSave(ByVal SaveType As Integer, ByVal Index As Integer)
Dim buffer As clsBuffer
Dim I As Integer
Dim B As Integer
'SaveType
'1=options
'2=users
'3=ranks
 If Index = 0 Then Exit Sub


    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveGuild
    
    buffer.WriteInteger SaveType
    buffer.WriteInteger Index
    
    Select Case SaveType
    Case 1
        'options
        buffer.WriteInteger GuildData.Guild_Color
        buffer.WriteInteger GuildData.Guild_RecruitRank
        buffer.WriteString GuildData.Guild_Name
        buffer.WriteString GuildData.Guild_Tag
        buffer.WriteString GuildData.Guild_MOTD
    Case 2
        'users
        buffer.WriteInteger GuildData.Guild_Members(Index).Rank
        buffer.WriteString GuildData.Guild_Members(Index).Comment
    Case 3
        'ranks
        buffer.WriteString GuildData.Guild_Ranks(Index).name
        For I = 1 To MAX_GUILD_RANKS_PERMISSION
            buffer.WriteByte GuildData.Guild_Ranks(Index).RankPermission(I)
        Next I
    End Select

    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
