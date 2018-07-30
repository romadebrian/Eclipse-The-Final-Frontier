Attribute VB_Name = "modSkills"
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Skill System By: escfoe2~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'

' The words you'll see below that are in capital letters represent sections.
' These sections are in big green borders like the one you see here that says "DIRECTIONS FOR USE".
' Except the other sections have other names like "MAX_SKILLS" and the rest you see below.
' Simply locate them and the rest should be self-explanatory

'(1) Raise MAX_SKILLS value by 1
'(2) Add another INDEX.                 [Example is shown.]
'(3) Add another item in NAMES.         [Example is shown.]
'(4) Add another item in MAX LEVELS.    [Example is shown.]
'(5) Add another item in DIVISORS.      [Example and explanation are shown.]

'Turn on your server, connect your client, and enjoy the new skill.
'SKILLS ARE SHOWN IN YOUR SKILL LIST BROUGHT UP BY PRESSING 'K' ON YOUR KEYBOARD

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'


'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~MAX_SKILLS~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
Public Const MAX_SKILLS As Long = 4 '<- Plus 1
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'


'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~INDEX~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
Public Const SKILL_CRAFTING = 1
Public Const SKILL_MINING = 2
Public Const SKILL_WOODCUTTING = 3
Public Const SKILL_FISHING = 4
'Public Const SKILL_YOURSKILL = NEXT-AVAILABLE-NUMBER (5 in this case)
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'


Sub SetMainSkillData()
    
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~NAMES~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
    Skill(SKILL_MINING).Name = "Mining"
    Skill(SKILL_CRAFTING).Name = "Crafting"
    Skill(SKILL_FISHING).Name = "Fishing"
    Skill(SKILL_WOODCUTTING).Name = "WoodCutting"
    'Skill(SKILL_YOURSKILL).Name = "NAMEofSKILL"
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
    
    
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~MAX LEVELS~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
    Skill(SKILL_MINING).MaxLvl = 100
    Skill(SKILL_CRAFTING).MaxLvl = 100
    Skill(SKILL_FISHING).MaxLvl = 100
    Skill(SKILL_WOODCUTTING).MaxLvl = 100
    'Skill(SKILL_YOURSKILL).MaxLvl = "MAXLVLofSKILL"
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
    
    
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~DIVISORS~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
' EXP Rise dividers
' 50 is the regular combat rate
' The higher the number, the more exp per level you'll need
' Crafting has more because with just these four skills, you'll use it the most.
    Skill(SKILL_MINING).Div = 35
    Skill(SKILL_CRAFTING).Div = 65
    Skill(SKILL_FISHING).Div = 35
    Skill(SKILL_WOODCUTTING).Div = 35
    'Skill(SKILL_YOURNEWSKILL_DIV).Div = YOURDIV
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~Get/Set Data~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
Function GetPlayerNextSkillLevel(ByVal index As Long, ByVal tSkill As Byte) As Long
Dim Level As Long, I As Long
    Level = GetPlayerSkillLevel(index, tSkill) + 1
    For I = 1 To MAX_SKILLS
        If I = tSkill Then
            GetPlayerNextSkillLevel = (Skill(I).Div / 3) * ((Level) ^ 3 - (6 * (Level) ^ 2) + 17 * (Level) - 12)
            Exit For
        End If
    Next I
End Function

Function GetPlayerSkillLevel(ByVal index As Long, ByVal Skill As Long) As Long
    If Skill < 1 Or Skill > MAX_SKILLS Then Exit Function
    GetPlayerSkillLevel = Player(index).Skills(Skill).Level
End Function

Function GetPlayerSkillExp(ByVal index As Long, ByVal Skill As Long) As Long
    GetPlayerSkillExp = Player(index).Skills(Skill).EXP
End Function

Sub SetPlayerSkillExp(ByVal index As Long, ByVal tSkill As Long, ByVal EXP As Long, Optional ByVal PlusVal As Boolean = True)
Dim OverHang As Long

    If PlusVal Then
        Player(index).Skills(tSkill).EXP = Player(index).Skills(tSkill).EXP + EXP
    Else
        Player(index).Skills(tSkill).EXP = EXP
    End If
    
    If GetPlayerSkillLevel(index, tSkill) < Skill(tSkill).MaxLvl And GetPlayerSkillExp(index, tSkill) > GetPlayerNextSkillLevel(index, tSkill) Then
        OverHang = GetPlayerSkillExp(index, tSkill) - GetPlayerSkillEXPNeeded(index, tSkill)
        Player(index).Skills(tSkill).EXP = OverHang
        Call SetPlayerSkillLevel(index, tSkill, 1)
        Player(index).Skills(tSkill).EXP_Needed = GetPlayerNextSkillLevel(index, tSkill)
    End If
End Sub

Function GetPlayerSkillEXPNeeded(ByVal index As Long, ByVal tSkill As Long)
    GetPlayerSkillEXPNeeded = Player(index).Skills(tSkill).EXP_Needed
End Function

Function SetPlayerSkillLevel(ByVal index As Long, ByVal tSkill As Long, ByVal Value As Long, Optional ByVal PlusVal As Boolean = True) As Boolean
    If PlusVal Then
        SetPlayerSkillLevel = False
        If Player(index).Skills(tSkill).Level + Value > Skill(tSkill).MaxLvl Then Exit Function
        Player(index).Skills(tSkill).Level = Player(index).Skills(tSkill).Level + Value
        SetPlayerSkillLevel = True
    Else
        SetPlayerSkillLevel = False
        If Value > MAX_LEVELS Then Exit Function
        Player(index).Skills(tSkill).Level = Value
        SetPlayerSkillLevel = True
    End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~SUBS~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
Sub CheckSkills(ByVal index As Long, Optional ByVal SaveIfNeeded As Boolean = True)
    Dim I As Integer
    Dim Save As Boolean
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Save = False
    For I = 1 To MAX_SKILLS
        If Player(index).Skills(I).Level = 0 Then
            Player(index).Skills(I).Level = 1
            Save = True
        End If
        
        If Player(index).Skills(I).EXP_Needed = 0 Then
            Player(index).Skills(I).EXP_Needed = GetPlayerNextSkillLevel(index, I)
            Save = True
        End If
    Next I
    If Save And SaveIfNeeded Then Call SavePlayer(index)
End Sub
