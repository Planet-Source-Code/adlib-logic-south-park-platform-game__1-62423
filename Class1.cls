VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 '=== Character.cls === ERRORCODE100 ==== NOVEMBER 2004 ====

Option Explicit

'APIs
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal np_widthth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest


'the handel to the DC containing the bitmap
Private DC As Long

'the hanle to the DC of the main game
Private parentDC As Long


'in pixels (buit converted to form1's scale mode when returned/saved as properties
Private p_width As Long
Private p_height As Integer
Private p_Left As Integer
Private p_Top As Integer
Private p_SpeachText As String
Private space As Integer

Private p_Zorder As Integer
Private TextLines() As String
Private p_Distance As Single

'states wherether the picture contains a mask or not
Private isSprite As Boolean

'the index of the picture to display (one horizontal slice of the bitmap)
Private p_Index As Integer

'if this is false then painting does not do anything
Private p_Visible As Boolean


'get a DC of the requested bitmap and set up various properties
Public Sub init(ByVal pictureAddress As String, ByVal desieredParentDC As Long, ByVal heightOfPictureInPixels, PictureIsASprite As Boolean, d As Single)
p_Distance = d
parentDC = desieredParentDC
Dim pic As StdPicture
Set pic = LoadPicture(pictureAddress)
DC = DCfromBMPfile(pictureAddress)
isSprite = PictureIsASprite

If isSprite Then
p_width = Form1.ScaleX(pic.width / 2, vbHimetric, vbPixels)
Else
p_width = Form1.ScaleX(pic.width, vbHimetric, vbPixels)
End If

p_height = heightOfPictureInPixels
p_Zorder = giveMeAZorder(Me)
p_Visible = True
space = p_width / 3
End Sub

'draw the object as its properties state to the parent DC (the parent must be refreshed to see the changes)
Public Sub paint()
If Visible = False Then Exit Sub
If isSprite Then
BitBlt parentDC, p_Left, p_Top, p_width, p_height, DC, 0, p_height * p_Index, SRCAND
BitBlt parentDC, p_Left, p_Top, p_width, p_height, DC, p_width, p_Index * p_height, SRCPAINT
Else
BitBlt parentDC, p_Left, p_Top, p_width, p_height, DC, 0, p_height * p_Index, SRCCOPY
End If
If p_SpeachText <> "" Then speak p_SpeachText, p_Left, p_Top, parentDC, TextLines
End Sub

'read/write properties
Public Property Let Left(l As Long)
p_Left = l
End Property
Public Property Get Left() As Long
Left = p_Left
End Property
Public Property Get LeftF() As Integer
LeftF = p_Left + space
End Property
Public Property Get WidthF() As Integer
WidthF = p_width - space * 2
End Property

Public Property Let Top(t As Integer)
p_Top = t
End Property
Public Property Get Top() As Integer
Top = p_Top
End Property

Public Property Let index(i As Integer)
p_Index = i
End Property
Public Property Get index() As Integer
index = p_Index
End Property

Public Property Let Visible(b As Boolean)
p_Visible = b
End Property
Public Property Get Visible() As Boolean
Visible = p_Visible
End Property

'read-only properties
Public Property Get Zorder() As Integer
Zorder = p_Zorder
End Property

Public Property Get width() As Long
width = p_width
End Property

Public Property Get height() As Integer
height = p_height
End Property

Property Get hdc() As Long
hdc = DC
End Property


'Detect wether the character overlaps a rectangle (expressed by (x,y) of top left and with and height)
Public Function overlapsRectangle(x As Integer, y As Integer, w As Integer, h As Integer) As Boolean
overlapsRectangle = True
If x + w < p_Left Then overlapsRectangle = False
If x > p_Left + p_width Then overlapsRectangle = False
If y + h < p_Top Then overlapsRectangle = False
If y > p_Top + p_height Then overlapsRectangle = False
End Function

'Detect whether the character's shape overlaps with another shape, this is a bit by bit comparison and is slow
Public Function overlapsRectangleAccurate(x As Integer, y As Integer, w As Integer, h As Integer, maskDC As Long) As Boolean
Dim nx, ny, nw, nh As Integer
Dim X1, X2, Y1, Y2, aw, ah As Integer

nx = Form1.ScaleX(x, Form1.ScaleMode, vbPixels)
ny = Form1.ScaleY(y, Form1.ScaleMode, vbPixels)
nw = Form1.ScaleX(w, Form1.ScaleMode, vbPixels)
nh = Form1.ScaleX(h, Form1.ScaleMode, vbPixels)

If nx > p_Left Then
X1 = nx - p_Left
X2 = 0
aw = p_Left + p_width - nx
    If aw > nw Then aw = nw
Else
X1 = 0
X2 = p_Left - nx
aw = nx + nw - p_Left
    If aw > p_width Then aw = p_width
End If

If ny < p_Top Then
Y1 = 0
Y2 = p_Top - ny
ah = ny + nh - p_Top
    If p_height < ah Then ah = p_height
Else
Y1 = ny - p_Top
Y2 = 0
ah = p_Top + p_height - ny
If ah > nh Then ah = nh
End If


Dim i, j As Integer

    For j = 0 To ah - 1
        For i = 0 To aw - 1
            If GetPixel(DC, i + X1, j + Y1) = 0 And GetPixel(maskDC, i + X2, j + Y2) = 0 Then
            overlapsRectangleAccurate = True
            End If
        Next i
    Next j


End Function

Public Property Get Distance() As Single
Distance = p_Distance
End Property

Public Property Let speachText(t As String)
p_SpeachText = t
End Property
Public Property Get speachText() As String
speachText = p_SpeachText
End Property

