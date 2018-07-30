Attribute VB_Name = "modGlobals"
Option Explicit

' Used for closing key doors again
Public KeyTimer As Long

' TopKill Competition Event
Public TopKill_Activated As Boolean
Public Reg_Kills As Long
Public hkPlace(1 To 5) As String
Public hkKills(1 To 5) As Long
Public Placements(1 To 5) As String
Public Time_Seconds As Byte
Public Time_Minutes As Integer
Public Time_Hours As Integer

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Used for Double Exp
Public DoubleExp As Boolean

' Used for logging
Public ServerLog As Boolean

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean
