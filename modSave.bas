Attribute VB_Name = "modSave"
Option Explicit

Public Type GameData
    Gamename As String
    FullName As String
    Handoff As String
    Filter As String
End Type

'----------------------------------------
' Save the Serverlist
'----------------------------------------
Public Function SaveServerlist(ByRef Game() As GameData) As String
Dim Data As String
Dim F As Integer
Dim i As Long

On Error GoTo SaveError
Data = ""

For i = 0 To UBound(Game)
    Data = Data & Game(i).FullName & Chr$(0) & _
           Game(i).Gamename & Chr$(0) & Game(i).Handoff & _
           Chr$(0) & Game(i).Filter & Chr$(0)
Next

F = FreeFile 'Get new File-Number
Open App.Path & "\serverlist.dat" For Output As #F 'Open the file
Print #F, Left$(Data, Len(Data) - 1) 'Write Serverlist to file
Close #F 'Close

Exit Function
SaveError:
    SaveServerlist = "(" & Err.Number & ") " & Err.Description
End Function

'----------------------------------------
' Read the Serverlist
'----------------------------------------
Public Function ReadServerlist(ByRef Game() As GameData) As String
Dim Data As String
Dim Splitted() As String
Dim F As Integer
Dim i As Long
Dim j As Long

On Error GoTo ReadError

F = FreeFile 'Get new File-Number
Open App.Path & "\serverlist.dat" For Binary As #F 'Open the file
Data = Space$(LOF(F)) 'Fill the String with spaces..
Get #F, , Data 'Get the Serverlist
Close #F 'Close

Splitted() = Split(Data, Chr$(0), , vbTextCompare)
For i = 0 To UBound(Splitted) Step 4
    ReDim Preserve Game(j)
    
    'Fill the array:
    With Game(j)
        .FullName = Splitted(i)
        .Gamename = Splitted(i + 1)
        .Handoff = Splitted(i + 2)
        .Filter = Splitted(i + 3)
    End With

    j = j + 1
Next

Exit Function
ReadError:
    ReadServerlist = "(" & Err.Number & ") " & Err.Description
End Function
