VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GS Browser [Alpha 0.1]"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   7680
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   6360
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrCheck2 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5280
      Top             =   6720
   End
   Begin VB.Timer tmrCheck 
      Interval        =   1
      Left            =   4800
      Top             =   6720
   End
   Begin VB.Frame fraLog 
      Caption         =   "Log:"
      Height          =   1330
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   7455
      Begin VB.TextBox txtLog 
         Height          =   950
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   6
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame fraServers 
      Caption         =   "Servers:"
      Height          =   6735
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdSaveServers 
         Caption         =   "Save Servers"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ListBox lstServers 
         Height          =   5910
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraGames 
      Caption         =   "Games:"
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin MSWinsockLib.Winsock wskTCP 
         Left            =   4320
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CheckBox chkSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   6300
         Value           =   1  'Aktiviert
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdateServers 
         Caption         =   "Update Serverlist"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdateGames 
         Caption         =   "Update Gamelist"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ListBox lstGames 
         Height          =   5910
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=======================================
'
' This piece of Code really sucks,
' but I hope it's useful... :-P
'
' - Thomas Reiser <fire_1@gmx.de>
'=======================================

Private Type Server
    IP As String
    Port As Long
End Type

Private Enum Socket_State
    RequestingGames
    ReceivingGamelist
    ReceivingGameinfos
    RequestingValidateKey
    ReceivingServers
    Idle
End Enum

'---------------------------------------
' GameSpy Servers..
'---------------------------------------
'MotD-Master
Const GSMOTDMasterHost As String = "motd.gamespy.com"
Const GSMOTDMasterPort As Integer = 80

'Server-Master
Const GSServerMasterHost As String = "master.gamespy.com"
Const GSServerMasterPort As Integer = 28900

'----

Dim MOTDRequest As String
Dim Gamelist(1) As String
Dim Game() As GameData
Dim TempFilters() As GameData
Dim Serverlist As String
Dim Servers() As Server
Dim SocketState As Socket_State

'----------------------------------------
' Save Serverlist
'----------------------------------------
Private Sub cmdSaveServers_Click()
Dim F As Integer
Dim IPs As String
Dim i As Integer

With dlgSave
    .DialogTitle = "Save Serverlist to.."
    .Filter = "*.* (All files)|*.*" 'Accept all files
    .CancelError = False
    .ShowSave
    
    If .FileName = "" Then 'Abort-Button pressed
        Exit Sub
    Else
        If GetUBound2(Servers()) < 0 Then
            AddLog vbCrLf & "Serverlist is empty!"
            Exit Sub
        Else
            For i = 0 To GetUBound2(Servers())
                IPs = IPs & Servers(i).IP & ":" & Servers(i).Port & vbCrLf
            Next
        End If
        IPs = Left$(IPs, Len(IPs) - 2)
        
        'Save IPs to file
        F = FreeFile 'Get new File-Number
        Open .FileName For Output As #F 'Open the file
        Print #F, Left$(IPs, Len(IPs) - 1) 'Write Serverlist to file
        Close #F 'Close
        
        AddLog vbCrLf & "Serverlist saved to " & Chr$(34) & .FileName & Chr$(34) & "!"
    End If
End With
End Sub

'----------------------------------------
' Update Gamelist
'----------------------------------------
Private Sub cmdUpdateGames_Click()
Dim i As Integer

For i = 0 To GetUBound(Game())
    ReDim Preserve TempFilters(i)
    TempFilters(i).Filter = Game(i).Filter
    TempFilters(i).Gamename = Game(i).Gamename
Next

wskTCP.Close

'Connect to motd.gamespy.com:80..
wskTCP.Connect GSMOTDMasterHost, GSMOTDMasterPort
lstGames.Clear
SocketState = RequestingGames
MOTDRequest = "" '!!!

AddLog vbCrLf & "Connecting to " & GSMOTDMasterHost & ":" & GSMOTDMasterPort & "... (Gamelist)"
End Sub

'----------------------------------------
' Update Serverlist
'----------------------------------------
Private Sub cmdUpdateServers_Click()
Dim RetVal As String 'Inputbox return value

If lstGames.ListIndex = -1 Then
    'No Game selected..
    Exit Sub
Else
    RetVal = InputBox("Filters for '" & Game(lstGames.ItemData(lstGames.ListIndex)).FullName & "':", _
                      "Filter...", Game(lstGames.ItemData(lstGames.ListIndex)).Filter)
    
    If StrPtr(RetVal) = 0 Then
        Exit Sub
    Else
        'Add the new Filter to the Array:
        Game(lstGames.ItemData(lstGames.ListIndex)).Filter = RetVal
    End If
End If

wskTCP.Close

'Connect to master.gamespy.com:28900..
wskTCP.Connect GSServerMasterHost, GSServerMasterPort

AddLog vbCrLf & "Trying to get the Serverlist for '" & Game(lstGames.ItemData(lstGames.ListIndex)).FullName & "'..."
SocketState = RequestingValidateKey
End Sub

'----------------------------------------
' Form_Load-Event
'----------------------------------------
Private Sub Form_Load()
Dim RetVal As String
Dim i As Long

If Dir$(App.Path & "\serverlist.dat") <> "" Then
    RetVal = ReadServerlist(Game())
    
    If RetVal = "" Then
        'OK
        AddLog "Reading Serverlist..."
        For i = 0 To UBound(Game)
            lstGames.AddItem Game(i).FullName
            lstGames.ItemData(lstGames.NewIndex) = i
        Next
        AddLog " Done!" & vbCrLf & "GS Browser successfully started!"
    Else
        'Error
        AddLog "[ERROR] " & RetVal & vbCrLf & "Error while loading the Serverlist!"
    End If
Else
    AddLog "GS Browser successfully started!"
End If

SocketState = Idle
End Sub

'----------------------------------------
' Form_Unload-Event
'----------------------------------------
Private Sub Form_Unload(Cancel As Integer)
If chkSave.Value = vbChecked Then
    'Save the Serverlist (!!FILTERS!!)
    SaveServerlist Game()
End If
End Sub

'----------------------------------------
' Check if the Buttons are available
'----------------------------------------
Private Sub tmrCheck_Timer()
If SocketState = Idle Then
    cmdUpdateGames.Enabled = True
    
    If lstGames.ListCount > 0 Then
        cmdUpdateServers.Enabled = True
    Else
        cmdUpdateServers.Enabled = False
    End If
Else
    cmdUpdateGames.Enabled = False
    cmdUpdateServers.Enabled = False
End If

If lstGames.ListCount = 0 Then
    cmdUpdateServers.Enabled = False
Else
    cmdUpdateServers.Enabled = True
End If
End Sub

'----------------------------------------
' 15 seconds Timeout (Serverlist request)
'----------------------------------------
Private Sub tmrCheck2_Timer()
If Len(Serverlist) = 0 Then
    wskTCP.Close
    SocketState = Idle

    AddLog " Done! (0 Servers found!)"
End If

tmrCheck2.Enabled = False
End Sub

'----------------------------------------
' Request the Gamelist
'----------------------------------------
Private Sub wskTCP_Connect()
AddLog vbCrLf & "Connected! Requesting Data..."

If SocketState = RequestingGames Then
    'Request the Games:
    wskTCP.SendData "GET /software/services/index.aspx" & MOTDRequest & " HTTP/1.0" & vbCrLf & _
                    "Host: " & GSMOTDMasterHost & vbCrLf & _
                    "User-Agent: GS Browser/0.1" & vbCrLf & vbCrLf
End If
End Sub

'----------------------------------------
' Parse IPs in an 6-Byte-IP Packet
'----------------------------------------
Private Function DecompressIps(ByRef Servers() As Server, ByVal Data As String) As Boolean
Dim IP(3) As String
Dim Port(1) As Long
Dim i As Long
Dim c As Long

If Len(Data) = 0 Then
    DecompressIps = False
    Exit Function
Else
    Data = Replace$(Data, "\final\", "")
End If

i = 1
While i < Len(Data) - 6
    'IP:
    IP(0) = Asc(Mid$(Data, i, 1))       'XXX.000.000.000
    IP(1) = Asc(Mid$(Data, i + 1, 1))   '000.XXX.000.000
    IP(2) = Asc(Mid$(Data, i + 2, 1))   '000.000.XXX.000
    IP(3) = Asc(Mid$(Data, i + 3, 1))   '000.000.000.XXX
    
    'Port:
    Port(0) = Asc(Mid$(Data, i + 4, 1))
    Port(1) = Asc(Mid$(Data, i + 5, 1))
    
    c = GetUBound2(Servers) + 1
    ReDim Preserve Servers(c)
    
    With Servers(c)
        .IP = IP(0) & "." & IP(1) & "." & IP(2) & "." & IP(3)
        .Port = (Port(0) * 256) + Port(1)
    End With
    
    i = i + 6
Wend

DecompressIps = True
End Function

'----------------------------------------
' Receive TCP-Data
'----------------------------------------
Private Sub wskTCP_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Dim Splitted() As String
Dim i As Long

wskTCP.GetData Data, vbString, bytesTotal

Select Case SocketState
    Case Idle 'Do nothing
        Exit Sub
    Case RequestingValidateKey
        Dim ValidateKey As String
        Dim Index As Integer
    
        If Left$(Data, 15) = "\basic\\secure\" Then
            Index = lstGames.ItemData(lstGames.ListIndex)
            
            'Create the Validate-Key:
            ValidateKey = makeValidate(Right$(Data, 7), getHandoff(Game(Index).Handoff))
            
            'Request the Serverlist:
            wskTCP.SendData createPacket(Game(Index).Gamename, ValidateKey, _
                                         Game(Index).Filter, True)
            SocketState = ReceivingServers
            Serverlist = ""
            
            Erase Servers
            tmrCheck2.Enabled = True
            AddLog vbCrLf & "Receiving Serverlist..."
        End If
    Case ReceivingServers
        tmrCheck2.Enabled = False
        Serverlist = Serverlist & Data
        AddLog "."
        
        'Check if all Servers are received:
        If InStr(1, Serverlist, "\final\", vbBinaryCompare) > 0 Then
            Serverlist = Replace$(Serverlist, "\final\", "")
            wskTCP.Close
            
            AddLog " OK!" & vbCrLf & "Decompressing Servers..."
            lstServers.Clear '!!!
            
            'Decompress the IPs (4 Byte IP, 2 Byte Port)
            If DecompressIps(Servers(), Serverlist) = True Then
                If GetUBound2(Servers()) > -1 Then
                    For i = 0 To UBound(Servers)
                        'Add Server to listbox
                        lstServers.AddItem Servers(i).IP & ":" & Servers(i).Port
                    Next
                    AddLog " Done! (" & UBound(Servers) + 1 & " Servers found!)"
                Else
                    GoTo NoServers
                End If
            Else
NoServers:
                AddLog " Done! (0 Servers found!)"
            End If
            
            SocketState = Idle
        End If
    Case RequestingGames
        Dim HeaderLen As Integer
        
        AddLog "."
        
        'Split Header/Data
        Splitted() = Split(Data, vbCrLf & vbCrLf, 2, vbTextCompare)
        HeaderLen = Len(Splitted(0))
        Gamelist(1) = Splitted(1)
        
        'Split the Header
        Splitted() = Split(Data, vbCrLf, , vbTextCompare)
        For i = 0 To UBound(Splitted)
            
            'Get the Content-Length:
            If Left$(Splitted(i), 16) = "Content-Length: " Then
                Gamelist(0) = CLng(Mid$(Splitted(i), 17)) - (bytesTotal - HeaderLen)
                Exit For
            End If
        Next
        
        If Len(MOTDRequest) = 0 Then
            SocketState = ReceivingGamelist
        Else
            SocketState = ReceivingGameinfos
        End If
    Case ReceivingGamelist 'Parse the incoming Gamelist..
        Gamelist(0) = Gamelist(0) - bytesTotal
        Gamelist(1) = Gamelist(1) & Data
        AddLog "."
        
        'Check if all Data is received:
        If Gamelist(0) <= 0 Then
            AddLog " Done!"
            
            Dim Split2() As String
        
            wskTCP.Close
            MOTDRequest = "?mode=full&services="
            
            'Parse the Gamelist
            Splitted() = Split(Gamelist(1), vbLf, , vbTextCompare)
            For i = 0 To UBound(Splitted)
                If Len(Splitted(i)) > 0 Then
                    Split2() = Split(Splitted(i), " - ", 2, vbTextCompare)
                    
                    'Filter invalid games:
                    If Left$(Split2(0), 2) <> "g_" And _
                       Left$(Split2(0), 4) <> "test" And _
                       Left$(Split2(0), 1) <> "_" And _
                       Split2(0) <> "fileplanet" Then
                       
                        MOTDRequest = MOTDRequest & Split2(0) & "\" 'Add game to our Request-String
                    End If
                End If
            Next
            
            'Remove the "\" at the end:
            MOTDRequest = Left$(MOTDRequest, Len(MOTDRequest) - 1)

            'Connect to motd.gamespy.com:80..
            wskTCP.Connect GSMOTDMasterHost, GSMOTDMasterPort
            
            AddLog vbCrLf & "Connecting to " & GSMOTDMasterHost & ":" & GSMOTDMasterPort & "... (Gameinfos)"
            Erase Game
            SocketState = RequestingGames
        End If
    Case ReceivingGameinfos
        Gamelist(0) = Gamelist(0) - bytesTotal
        Gamelist(1) = Gamelist(1) & Data
        AddLog "."

        Dim FirstSection As Boolean
        Dim Key() As String
        Dim c As Integer
        Dim j As Integer
        Dim Temp As GameData
        
        FirstSection = True
        c = 0

        'Check if all Data is received:
        If Gamelist(0) <= 0 Then
            AddLog " OK." & vbCrLf & "Parsing Gameinfos..."
            Splitted() = Split(Gamelist(1), vbLf, , vbTextCompare)
            
            'Parse the INI-Strings
            For i = 0 To UBound(Splitted)
                Splitted(i) = Trim$(Splitted(i))
                If Left$(Splitted(i), 1) = "[" Then
                    '--------------------------
                    ' INI-Section
                    '--------------------------
                    If FirstSection = True Then
                        'Gamename
                        Temp.Gamename = Mid$(Splitted(i), 2, Len(Splitted(i)) - 2)
                        FirstSection = False
                    Else
                        'Check for invalid INI-Keys:
                        If Len(Temp.Gamename) > 0 And _
                           Len(Temp.Handoff) > 0 And _
                           Len(Temp.FullName) > 0 Then
                           
                            'Add the valid game to our Array:
                            c = GetUBound(Game) + 1
                            ReDim Preserve Game(c)

                            With Game(c)
                                .Gamename = Temp.Gamename
                                .FullName = Temp.FullName
                                .Handoff = Temp.Handoff
                                
                                For j = 0 To GetUBound(TempFilters())
                                    If TempFilters(j).Gamename = Temp.Gamename Then
                                        Game(c).Filter = TempFilters(j).Filter
                                        Exit For
                                    End If
                                Next
                                
                                lstGames.AddItem Temp.FullName
                                lstGames.ItemData(lstGames.NewIndex) = c
                                    
                                'Delete the Temp-Array:
                                Temp.Gamename = Mid$(Splitted(i), 2, Len(Splitted(i)) - 2)
                                Temp.FullName = ""
                                Temp.Handoff = ""
                            End With
                        End If
                    End If
                Else
                    '--------------------------
                    ' INI-Key
                    '--------------------------
                    Key() = Split(Splitted(i), "=", 2, vbTextCompare)
                    If UBound(Key) > 0 Then
                        If LCase$(Key(0)) = "handoff" Then
                            Temp.Handoff = Key(1)
                        ElseIf LCase$(Key(0)) = "fullname" Then
                            Temp.FullName = Key(1)
                        End If
                    End If
                End If
            Next
            
            Dim RetVal As String
            If chkSave.Value = vbChecked Then
                AddLog vbCrLf & "Saving Gamelist... "
                
                RetVal = SaveServerlist(Game())
                If RetVal <> "" Then
                    'Error
                    AddLog vbCrLf & "[ERROR] " & RetVal & vbCrLf
                End If
            End If
            
            AddLog "Done!"
            SocketState = Idle
        End If
End Select
End Sub

'----------------------------------------
' Log-Sub :D
'----------------------------------------
Private Sub AddLog(ByVal Text As String)
If Len(Text) > 0 Then
    txtLog.Text = txtLog.Text & Text
    txtLog.SelStart = Len(txtLog.Text)
End If
End Sub

'GRML:
Private Function GetUBound(ByRef Arr() As GameData) As Long
On Error GoTo IndexError

GetUBound = UBound(Arr)
Exit Function

IndexError:
    GetUBound = -1
End Function

Private Function GetUBound2(ByRef Arr() As Server) As Long
On Error GoTo IndexError

GetUBound2 = UBound(Arr)
Exit Function

IndexError:
    GetUBound2 = -1
End Function
