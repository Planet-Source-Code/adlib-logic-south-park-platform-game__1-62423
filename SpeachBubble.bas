Attribute VB_Name = "SpeachBubble"
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Const maxLength = 300
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Private Type POINTAPI
        x  As Long
        y  As Long
End Type


Public Sub speak(text As String, Right As Integer, Bottom As Integer, DC As Long, tLine() As String)
Dim pic As StdPicture
Dim myDC As Long
Dim s As Long
Dim length As Long
Dim height As Integer
Dim myF As Long
Dim F As Long
Dim ULine As Integer
Dim p As Integer
Dim lp As Integer
Dim le As Integer
Dim j As Integer

Static wc(255) As Long
Static initedYet As Boolean

'set the font to the one that is to be used
myF = CreateFont(14, 8, 0, 0, 100, False, False, False, 1, 0, 0, 1, 0, "")
F = SelectObject(DC, myF)

'Get the caracter dimensions (only happens the first time)
If initedYet = False Then
    For i = 0 To 255
    GetCharWidth DC, i, i, wc(i)
    Next i
initedYet = True
End If

If hasBounds(tLine) = False Then
'Split the text into several lines
lp = 1
ULine = -1
    For i = 1 To Len(text)
    le = le + wc(Asc(Mid(text, i, 1)))
        If le > maxLength Then
        ULine = ULine + 1
        p = InStrRev(text, " ", i)
        ReDim Preserve tLine(0 To ULine) As String
        tLine(ULine) = Mid(text, lp, p + 1 - lp)
        le = 0
            For j = p + 1 To i
            le = le + wc(Asc(Mid(text, j, 1)))
            Next j
        lp = p + 1
        length = maxLength
        End If
    Next i

    If length = 0 Then length = le
    
ULine = ULine + 1
ReDim Preserve tLine(0 To ULine) As String
tLine(ULine) = Mid(text, lp)
height = (ULine + 1) * 20
ReDim Preserve tLine(0 To ULine + 2) As String
tLine(UBound(tLine)) = length
tLine(UBound(tLine) - 1) = height
Else
length = Val(tLine(UBound(tLine)))
height = Val(tLine(UBound(tLine) - 1))
End If

'At this point:
'length is the length of the textpart of the bubble
'and height is the height of the text part of the bubble

Dim b As RECT
b.Right = Right
b.Bottom = Bottom
b.Left = b.Right - length
b.Top = b.Bottom - height
Dim brush As LOGBRUSH
Dim o As Long
Dim nb As Long
brush.lbColor = vbWhite
nb = CreateBrushIndirect(brush)
o = SelectObject(DC, nb)
RoundRect DC, b.Left - 5, b.Top - 5, b.Right + 5, b.Bottom + 5, 20, 20

Dim tick(0 To 2) As POINTAPI
tick(0).x = 5 + b.Right
tick(0).y = b.Bottom
tick(1).x = 10 + b.Right
tick(1).y = 10 + b.Bottom
tick(2).x = b.Right
tick(2).y = 7 + b.Bottom


Polygon DC, tick(0), 3

SelectObject DC, o
DeleteObject nb

For i = 0 To ULine
TextOut DC, b.Left, b.Top + i * 20, tLine(i), Len(tLine(i))
Next i

SelectObject DC, F
DeleteObject myF

End Sub

Private Function hasBounds(a() As String) As Boolean
On Error GoTo no
Dim p As Integer
hasBounds = True
p = UBound(a)
Exit Function
no:
hasBounds = False
End Function
