Attribute VB_Name = "modGameSpy"
'+----------------------------------------------------------------------------
'| modGameSpy v0.4.1
'|
'| Written by FiRe^ (fire_1@gmx.de)
'| Last edit: 2004-11-04
'+----------------------------------------------------------------------------
'| Information:
'|  The algorithm for the function makeValidate() was converted from
'|  Luigi Auriemma's C-Code gsmsalg.h (http://aluigi.altervista.org/papers/gsmsalg.h)
'|
'|
'| Public functions:
'|  Create the Validate-Key:
'|    str makeValidate(str SecureKey, str Handoff)
'|
'|  Create a valid 6-char Handoff:
'|    str getHandoff(str Handoff)
'|
'|  Create a Master Packet:
'|    str createPacket(str Gamename, str ValidateKey [, str Filter [, bool CompressedServers = False]])
'|
'|
'| License (http://www.gnu.org/licenses/gpl.txt):
'|  This program is free software; you can redistribute it and/or modify
'|  it under the terms of the GNU General Public License as published by
'|  the Free Software Foundation; either version 2 of the License, or
'|  (at your option) any later version.
'|
'|  This program is distributed in the hope that it will be useful,
'|  but WITHOUT ANY WARRANTY; without even the implied warranty of
'|  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'|  GNU General Public License for more details.
'|
'|  You should have received a copy of the GNU General Public License
'|  along with this program; if not, write to the Free Software
'|  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'+----------------------------------------------------------------------------

Option Explicit

'+----------------------------------------------------------------------------
'| Function: getHandoff
'| Params:   Handoff: 14-char Handoff
'| Return:   6-char Handoff
'+----------------------------------------------------------------------------
Public Function getHandoff(ByVal Handoff As String) As String
Dim newHandoff  As String   '// The new 6-char Handoff
Dim i           As Byte     '// Loop-var

'// Handoff is too short:
If Len(Handoff) < 13 Then
    getHandoff = Handoff ':-(
    Exit Function
End If

For i = 3 To 13 Step 2
    '// Add next char to the new Handoff
    newHandoff = newHandoff & Mid$(Handoff, i, 1)
Next

getHandoff = newHandoff
End Function

'+----------------------------------------------------------------------------
'| Function: createPacket
'| Params:   Gamename: The internal Gamename (bfield1942, quake3, ...)
'|           ValidateKey: The ValidateKey (created with makeValidate!)
'|           [Filter]: Apply some Filters..
'|           [CompressedServers]: 'True' will return an 6-Byte-IP Packet
'| Return:   A valid GameSpy Master-Packet
'+----------------------------------------------------------------------------
Public Function createPacket(ByVal Gamename As String, ByVal ValidateKey As String, Optional ByVal Filter As String, Optional ByVal CompressedServers As Boolean) As String
'// This small function will create a GSMaster-Packet for you :)
createPacket = _
"\gamename\" & Gamename & "\enctype\0\validate\" & ValidateKey & "\final\" & _
"\queryid\1.1\list\" & IIf(CompressedServers = True, "cmp", "") & _
"\gamename\" & Gamename & IIf(Len(Filter) > 0, "\where\" & Filter, "") & "\final\"
End Function

'+----------------------------------------------------------------------------
'| Function: makeValidate
'| Params:   SecureKey: The Key received from the GS Master
'|           Handoff: Your game's handoff
'| Return:   Validate-Key
'+----------------------------------------------------------------------------
Public Function makeValidate(ByVal SecureKey As String, ByVal Handoff As String) As String
Dim Table(255)  As Byte     '// Buffer
Dim Key()       As Byte     '// (Secure)Key
Dim Length(1)   As Byte     '// Length(0): Handoff length
                            '// Length(1): SecureKey length
Dim Temp(3)     As Integer  '// Some temporary variables
Dim i           As Integer  '// Loop-var
Dim Validate    As String   '// ValidateKey

For i = 0 To 255
    Table(i) = i '// Fill the Buffer
Next

'// Add the length of the Keys to the array
Length(0) = Len(Handoff) '// Should be 6 chars
Length(1) = Len(SecureKey) '// Default is 6 chars

For i = 0 To 255
    '// Scramble the Table with the Handoff:
    Temp(0) = (Temp(0) + Table(i) + Asc(Mid$(Handoff, (i Mod Length(0)) + 1, 1))) And 255
    Temp(1) = Table(Temp(0))
    
    '// Update the buffer:
    Table(Temp(0)) = Table(i)
    Table(i) = Temp(1)
Next

Temp(0) = 0

ReDim Key(Length(1) - 1) '// Set the Array-Size to the SecureKey-Length
Length(1) = Length(1) \ 3

'// Scramble the SecureKey with the Table:
For i = 0 To UBound(Key)
    '// Add the next char to the Array
    Key(i) = Asc(Mid$(SecureKey, i + 1, 1))
    
    Temp(0) = (Temp(0) + Key(i) + 1) And 255
    Temp(1) = Table(Temp(0))
    
    Temp(2) = (Temp(2) + Temp(1)) And 255
    Temp(3) = Table(Temp(2))
    
    Table(Temp(2)) = Temp(1)
    Table(Temp(0)) = Temp(3)
    
    '// XOR the Key with the Buffer:
    Key(i) = Key(i) Xor Table((Temp(1) + Temp(3)) And 255)
    
    '// Encoding Type 2 (Completely useless)
    'Key(i) = Key(i) Xor Asc(Mid$(Handoff, (i Mod Length(0)) + 1, 1))
Next

i = 0
'// Create the valid ValidateKey:
While Length(1) >= 1 '// Default are 3 loops
    Length(1) = Length(1) - 1
    
    Temp(1) = Key(i)
    Temp(3) = Key(i + 1)
    
    '// VB has no >> << Operators :-(
    addChar Validate, RShift(Temp(1), 2)
    addChar Validate, LShift(Temp(1) And 3, 4) Or RShift(Temp(3), 4)
    
    Temp(1) = Key(i + 2)
    
    addChar Validate, LShift(Temp(3) And 15, 2) Or RShift(Temp(1), 6)
    addChar Validate, Temp(1) And 63
    
    i = i + 3
Wend

makeValidate = Validate '// Return the valid ValidateKey
End Function

Private Sub addChar(ByRef Validate As String, ByVal Number As Byte)
Dim newChar As String * 1

'// Check the Charcode, create a new Char ...
Select Case Number
    Case Is < 26
        newChar = Chr$(Number + 65)
    Case Is < 52
        newChar = Chr$(Number + 71)
    Case Is < 62
        newChar = Chr$(Number - 4)
    Case 62
        newChar = "+"
    Case 63
        newChar = "/"
End Select

'// ... and add it to the ValidateKey
Validate = Validate & newChar
End Sub

'// The << (LShift) and >> (RShift) functions:
Private Function LShift(ByVal Value As Byte, ByVal Shift As Byte) As Byte
LShift = Value * (2 ^ Shift)
End Function

Private Function RShift(ByVal Value As Byte, ByVal Shift As Byte) As Byte
RShift = Value \ (2 ^ Shift)
End Function
