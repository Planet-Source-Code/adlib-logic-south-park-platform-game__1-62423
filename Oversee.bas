Attribute VB_Name = "Oversee"
'======================This module is required by Characer.cls=======================
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function deleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As Long) As Long

Dim handelToCharacter() As Character
Public offset As Integer
Dim DC() As Long

'PAINTS
Public Sub paintAll()
Dim i As Integer
Dim t As Long

For i = 0 To UBound(handelToCharacter)
t = handelToCharacter(i).Left
handelToCharacter(i).Left = handelToCharacter(i).Left - offset * handelToCharacter(i).Distance
handelToCharacter(i).paint
handelToCharacter(i).Left = t
Next i

End Sub


'CHARACTERS CALL THIS TO REGISTER THEMSELVES FOR PAINTING
'it also returns a number signifying the zOrder
Public Function giveMeAZorder(ByRef C As Character, Optional resetVars As Boolean = False) As Integer
Static i As Integer

    If resetVars = True Then
    i = 0
    Exit Function
    End If

ReDim Preserve handelToCharacter(0 To i)
Set handelToCharacter(i) = C

giveMeAZorder = i
i = i + 1
End Function

Public Function DCfromBMPfile(rpath As String, Optional resetVars As Boolean = False) As Long
Static path() As String
Static i As Integer
Static InitYet As Boolean
Dim j As Integer
Dim r As Long

    If resetVars = True Then
    i = 0
    InitYet = False
    Exit Function
    End If
    
If InitYet = False Then GoTo newDCNeeded

    For j = 0 To i
        If path(j) = rpath Then
        DCfromBMPfile = DC(j)
        Exit Function
        End If
    Next j


newDCNeeded:
If InitYet = True Then
i = i + 1
Else
InitYet = True
End If

ReDim Preserve path(0 To i) As String
ReDim Preserve DC(0 To i) As Long

Dim pic As StdPicture
Set pic = LoadPicture(rpath)
DCfromBMPfile = CreateCompatibleDC(Form1.hdc)
 r = SelectObject(DCfromBMPfile, pic.Handle)
path(i) = rpath
DC(i) = DCfromBMPfile
'SelectObject DCfromBMPfile, r
End Function

Public Sub deleteAllDCs()
Dim i As Integer
For i = 0 To UBound(DC)
deleteDC DC(i)
Next i
End Sub
