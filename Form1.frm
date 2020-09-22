VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim coin() As Character
Dim Player As Character
Dim Background As Character
Dim exitPlease As Boolean
Dim keyDown(255) As Boolean
Dim signPost As Character
Dim road() As Character
Dim maxOffset As Long
Dim block() As Character
Dim sea() As Character
Dim baddy() As Character
Dim playerName As String
Dim baddyPause() As Integer
Dim shark As Character
Dim bus As Character
Dim NoSea As Boolean
Dim FinishLine As Long
Dim heart As Character
Dim lives As Integer

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long


Dim PlayerMode As Integer
Dim PlayerInAir As Boolean
Dim cCount As Integer

Enum spPlayerConstants
StillAlive = 0
died = 1
wonLevel = 2
End Enum


Private Sub Form_Click()
exitPlease = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
keyDown(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
keyDown(KeyCode) = False
End Sub

Private Function playlevel(i As Integer) As Boolean 'true = win, false = loose
Dim runningCutScene As Integer

cCount = 0
PlayerMode = 0
Oversee.giveMeAZorder Player, True
Oversee.DCfromBMPfile "", True
Oversee.offset = 0

'Load the characters
loadLevel App.path & "\level " & i & ".txt"
DoCoins True
doBaddy True
doPlayer True

'put the counting coin in the top left corner
Set coin(0) = New Character
coin(0).init App.path & "\coin.bmp", Form1.hdc, 23, True, 0
coin(0).Top = 6
coin(0).Left = 3

'put the heart in the top left corner
Set heart = New Character
heart.init App.path & "\heart.bmp", Form1.hdc, 21, True, 0
heart.Top = 30
heart.Left = 3

'position the player so that he will fall into the level
Player.Top = 20
Player.Left = 20
PlayerInAir = True

Me.Show

'MAIN LOOP
Do While exitPlease = False

    If runningCutScene = 0 Then
    'respond to input
    doKeys
    
    'operate player
        Select Case doPlayer
        Case died:
        deathMessage
        exitPlease = True
        playlevel = False
        Case wonLevel:
        runningCutScene = 1
        playlevel = True
        End Select
    Else
        Select Case runningCutScene
        Case 1:
        Set bus = New Character
        bus.init App.path & "\bus.bmp", Me.hdc, 107, True, 1
        bus.Left = Player.Left - bus.width - 500
        bus.Top = Player.Top + Player.height - bus.height
        Case 101:
        bus.index = 2
        Case 105:
        Player.Visible = False
        Case 107:
        bus.index = 0
        Case 120:
        exitPlease = True
        Case Else:
            If runningCutScene <= 101 Or runningCutScene >= 107 Then
            bus.Left = bus.Left + 5
                If bus.index = 1 Then
                bus.index = 0
                Else
                bus.index = 1
                End If
            End If
        End Select
        runningCutScene = runningCutScene + 1
    End If

'operate coins and heart
DoCoins
doHeart

'operate sea
If NoSea = False Then doSea

'operate baddy
doBaddy

'paint all sprites
paintAll

'paint the text in the top left of screen
paintText

'put the buffered data to the screen
Form1.Refresh

'wait for 0 miliseconds
'Sleep 0

'allow form events to trigger
DoEvents
Loop
'LEVEL ENDS
exitPlease = False
deleteAllDCs
End Function

Private Sub Form_Load()
Dim level As Integer
Randomize
level = 1
lives = 10
    Do While lives > 0
        If playlevel(level) = True Then
        level = level + 1
        Else
        lives = lives - 1
        End If
        If level = 7 Then lives = 0
    Loop
Unload Me
End Sub

Sub doHeart()
Static C As Integer
C = C + 1

If C = 3 Then
heart.index = heart.index + 1
    If heart.index = 4 Then heart.index = 0
C = 0
End If

End Sub
Sub doKeys()
If PlayerMode = 6 Then Exit Sub

If keyDown(vbKeyUp) = True Then
    If keyDown(vbKeyRight) = True And keyDown(vbKeyLeft) = False Then
    PlayerMode = 4 'jump right
    ElseIf keyDown(vbKeyLeft) = True And keyDown(vbKeyRight) = False Then
    PlayerMode = 5 'jump left
    Else
    PlayerMode = 3 'verticle jump
    End If
Else
    If keyDown(vbKeyRight) = True And keyDown(vbKeyLeft) = False Then
    PlayerMode = 1 'move right
    ElseIf keyDown(vbKeyLeft) = True And keyDown(vbKeyRight) = False Then
    PlayerMode = 2 'move left
    Else
    PlayerMode = 0 'still
    End If
End If
End Sub


Function doPlayer(Optional resetVars As Boolean = False) As spPlayerConstants
Static C As Integer
Static yv As Single
Dim i As Integer

    If resetVars = True Then
    C = 0
    yv = 0
    i = 0
    Exit Function
    End If
    
C = C + 1
Select Case PlayerMode
Case 0: 'stand still
    Player.index = 4
    If yv < 0 Then Player.index = 11
Case 1: 'move right
    If Player.index > 3 Then Player.index = 0
    If C Mod 5 = 0 Then
    Player.index = Player.index + 1
    If Player.index = 4 Then Player.index = 0
    End If
    Player.Left = Player.Left + 10
    If Player.Left - offset > Form1.ScaleWidth * 2 / 3 Then offset = offset + 10
    If offset > maxOffset Then offset = maxOffset
Case 2: 'move left
    If Player.index < 5 Then Player.index = 5
    If C Mod 5 = 0 Then
    Player.index = Player.index + 1
    If Player.index >= 9 Then Player.index = 5
    End If
    Player.Left = Player.Left - 10
    If Player.Left - offset < Form1.ScaleWidth / 3 Then offset = offset - 10
    If offset < 0 Then offset = 0
Case 3: 'jump
    Static tyv As Integer
    
    If PlayerInAir = False Then
    Player.index = 10
    tyv = tyv - 3.5
        If tyv <= -25 Then
        yv = tyv
        tyv = 0
        PlayerInAir = True
        Player.index = 11
        End If
    Else
    End If
Case 4: 'jump right
    If PlayerInAir = False Then yv = -25
    PlayerInAir = True
    Player.index = 9
    Player.Left = Player.Left + 10
    If Player.Left - offset > Form1.ScaleWidth * 2 / 3 Then offset = offset + 10
    If offset > maxOffset Then offset = maxOffset
Case 5: 'jump left
    If PlayerInAir = False Then yv = -25
    PlayerInAir = True
    Player.index = 12
    Player.Left = Player.Left - 10
    If Player.Left - offset < Form1.ScaleWidth / 3 Then offset = offset - 10
    If offset < 0 Then offset = 0
Case 6: 'die on land
    If C Mod 25 = 0 Then Player.index = Player.index + 1
    If Player.index = 16 Then
    doPlayer = died
    End If
    GoTo gravity
End Select

If PlayerMode <> 3 And tyv <> 0 Then
    yv = tyv
    tyv = 0
    PlayerInAir = True
    Player.index = 11
End If


gravity:
'gravity ==================
yv = yv + 2.5 'force of gravity
If yv > 20 Then yv = 20 'terminal velocity

PlayerInAir = True

For i = 0 To UBound(road) - 1
    If Player.LeftF + Player.WidthF > road(i).Left And Player.LeftF < road(i).Left + road(i).width Then
        If Player.Top + Player.height <= road(i).Top And Player.height + Player.Top + yv >= road(i).Top Then
        yv = 0
        PlayerInAir = False
        Player.Top = road(i).Top - Player.height
        End If
    End If
Next i

For i = 0 To UBound(block) - 1
If block(i).Visible = True Then
    If Player.LeftF + Player.WidthF > block(i).LeftF And Player.LeftF < block(i).LeftF + block(i).width Then
        If Player.Top + Player.height <= block(i).Top And Player.height + Player.Top + yv >= block(i).Top Then
        yv = 0
        PlayerInAir = False
        Player.Top = block(i).Top - Player.height
        End If
    End If
End If
Next i

Player.speachText = ""
If PlayerInAir = True Then
Player.Top = Player.Top + yv
    If yv > 18 Then Player.speachText = "Ahh!"
End If
'=========================================
If PlayerMode = 6 Then Exit Function

'death from falling=======================
If Player.Top > 500 Then
doPlayer = died
End If
'======================================

'collect coins=========================
'coin 0 is the display coin
For i = 1 To UBound(coin)
If coin(i).Visible = True Then
    If Player.overlapsRectangle(coin(i).Left, coin(i).Top, coin(i).width, coin(i).height) = True Then
    coin(i).Visible = False
    cCount = cCount + 1
    End If
End If
Next i
'======================================

'do block smashing=====================
For i = 0 To UBound(block) - 1
    If block(i).index = 0 Then
    If yv < 0 Then
        If block(i).overlapsRectangle(Player.Left, Player.Top, Player.width, Player.height) = True Then
        block(i).index = 1
        yv = -yv
        cCount = cCount + 1
        End If
    End If
    ElseIf block(i).Visible = True Then
    If C Mod 3 = 0 Then block(i).index = block(i).index + 1
        If block(i).index = 5 Then block(i).Visible = False
    End If
Next i
'======================================

'kill or die from baddies==============
For i = 0 To UBound(baddy) - 1
    If baddy(i).Visible = False Or baddyPause(i) >= 90 Then GoTo missBaddy

    If PlayerMode <> 6 And Player.overlapsRectangle(baddy(i).Left, baddy(i).Top, baddy(i).width, baddy(i).height) = True Then
        If yv > 0 Then Player.Top = Player.Top - yv

        If Player.overlapsRectangle(baddy(i).Left, baddy(i).Top, baddy(i).width, baddy(i).height) = True Then
        Player.index = 13
        PlayerMode = 6
        baddyPause(i) = 1
        baddy(i).index = 4
        Else
        baddyPause(i) = 90
        cCount = cCount + 5
        yv = -yv
        End If
        
    Player.Top = Player.Top + yv
    End If
    
missBaddy:
Next i
'======================================



'End level alive=======================
If Player.Left > FinishLine And PlayerInAir = False Then
doPlayer = wonLevel
End If
'======================================

End Function

Sub DoCoins(Optional resetVars As Boolean = False)
Static C As Integer
Static dir() As Boolean 'false = forwads
ReDim Preserve dir(0 To UBound(coin)) As Boolean

    If resetVars = True Then
    C = 0
    Exit Sub
    End If
    
C = C + 1
If C Mod 2 <> 0 Then Exit Sub
Dim i As Integer

For i = 0 To UBound(coin)
If coin(i).index = 7 Then dir(i) = True
If coin(i).index = 0 Then dir(i) = False
If dir(i) = True Then coin(i).index = coin(i).index - 1
If dir(i) = False Then coin(i).index = coin(i).index + 1
Next i
End Sub

Sub paintText()
Dim myF As Long
Dim s As Long
SetTextColor Me.hdc, vbYellow
SetTextAlign Form1.hdc, 0
myF = CreateFont(30, 10, 0, 0, 100, False, False, False, 1, 0, 0, 1, 0, "")
s = SelectObject(Form1.hdc, myF)
TextOut Form1.hdc, 30, 5, "" & cCount, Len("" & cCount)
TextOut Form1.hdc, 30, 30, "" & lives, Len("" & lives)
SelectObject Form1.hdc, s
DeleteObject myF
SetTextColor Me.hdc, vbBlack
End Sub

Sub doSea(Optional resetVars As Boolean = False)
On Error Resume Next
Dim i As Integer
Static C As Integer
Static sxv As Integer
    
    If resetVars = True Then
    C = 0
    sxv = 0
    End If
    
    If sxv = 0 Then
    sxv = 50
    End If
shark.Left = shark.Left + sxv
If shark.Left > sea(UBound(sea)).Left Then
sxv = -sxv
shark.index = 1
End If
If shark.Left < 0 Then
sxv = -sxv
shark.index = 0
End If
    If C Mod 10 = 0 Then
        If sxv > 0 Then sxv = Rnd * 10
        If sxv < 0 Then sxv = Rnd * -10
    End If

C = C + 1

If C Mod 3 <> 0 Then Exit Sub
For i = 0 To UBound(sea)
sea(i).index = sea(i).index + 1
If sea(i).index = 7 Then sea(i).index = 0
Next i

End Sub


Sub loadLevel(address As String)
Dim inputGroup As Integer
Dim s As String
Dim i As Integer
Dim a As String
Dim h As Integer
Dim b As Boolean
ReDim road(0) As Character
ReDim block(0) As Character
ReDim coin(0) As Character
ReDim baddy(0) As Character
ReDim baddyPause(0) As Integer
NoSea = False

    Open address For Input As #1
    Input #1, s
        Do While b = False
        b = EOF(1)
            Select Case s
            Case "[BACKGROUND]":
            inputGroup = 1
            Case "[MIDGROUND]":
            inputGroup = 2
            Case "[ROAD]":
            inputGroup = 3
            a = ""
            Case "[BLOCK]":
            inputGroup = 4
            a = ""
            Case "[SEA]":
            inputGroup = 5
            Case "[COIN]":
            inputGroup = 6
            a = ""
            Case "[PLAYER]":
            inputGroup = 7
            Case "[BADDY]":
            inputGroup = 8
            a = ""
            Case "[FINISH]":
            inputGroup = 9
            a = ""
            Case Else:
                Select Case inputGroup
                Case 1:
                Input #1, i
                Set Background = New Character
                Background.init App.path & s, Form1.hdc, i, False, 0.5
                maxOffset = (Background.width - Form1.ScaleWidth) / Background.Distance
                Case 2:
                Input #1, i
                Set signPost = New Character
                signPost.init App.path & s, Form1.hdc, i, True, 0.7
                Input #1, i
                signPost.Left = i
                Input #1, i
                signPost.Top = i
                Case 3:
                    If a = "" Then
                    a = s
                    Input #1, h
                    Set road(UBound(road)) = New Character
                    road(UBound(road)).init App.path & a, Form1.hdc, h, True, 1
                    Input #1, i
                    road(UBound(road)).Left = i
                    Input #1, i
                    road(UBound(road)).Top = i
                    ReDim Preserve road(0 To UBound(road) + 1) As Character
                    Else
                    Set road(UBound(road)) = New Character
                    road(UBound(road)).init App.path & a, Form1.hdc, h, True, 1
                    i = Val(s)
                    road(UBound(road)).Left = i
                    Input #1, i
                    road(UBound(road)).Top = i
                    ReDim Preserve road(0 To UBound(road) + 1) As Character
                    End If
                    
                Case 4:
                    If a = "" Then
                    a = s
                    Input #1, h
                    Set block(UBound(block)) = New Character
                    block(UBound(block)).init App.path & a, Form1.hdc, h, True, 1
                    Input #1, i
                    block(UBound(block)).Left = i
                    Input #1, i
                    block(UBound(block)).Top = i
                    ReDim Preserve block(0 To UBound(block) + 1) As Character
                    Else
                    Set block(UBound(block)) = New Character
                    block(UBound(block)).init App.path & a, Form1.hdc, h, True, 1
                    i = Val(s)
                    block(UBound(block)).Left = i
                    Input #1, i
                    block(UBound(block)).Top = i
                    ReDim Preserve block(0 To UBound(block) + 1) As Character
                    End If
                    
                Case 5:
                    If s = "T" Then
                    Set shark = New Character
                    shark.init App.path & "\shark.bmp", Form1.hdc, 35, True, 1
                    shark.Top = Form1.ScaleHeight - 40
                    ReDim sea(0) As Character
                    Set sea(0) = New Character
                    sea(0).init App.path & "\waves.bmp", Form1.hdc, 45, True, 1
                    sea(0).Top = Form1.ScaleHeight - sea(0).height
                    ReDim Preserve sea(0 To (maxOffset + Form1.ScaleWidth) / sea(0).width) As Character
                        For i = 1 To UBound(sea)
                        Set sea(i) = New Character
                        sea(i).init App.path & "\waves.bmp", Form1.hdc, 45, True, 1
                        sea(i).Left = sea(i).width * i
                        sea(i).Top = sea(0).Top
                        Next i
                    Else
                    NoSea = True
                    End If
                Case 6:
                    If a = "" Then
                    a = s
                    Input #1, h
                    ReDim Preserve coin(0 To UBound(coin) + 1) As Character
                    Set coin(UBound(coin)) = New Character
                    coin(UBound(coin)).init App.path & a, Form1.hdc, h, True, 1
                    Input #1, i
                    coin(UBound(coin)).Left = i
                    Input #1, i
                    coin(UBound(coin)).Top = i
                    Else
                    ReDim Preserve coin(0 To UBound(coin) + 1) As Character
                    Set coin(UBound(coin)) = New Character
                    coin(UBound(coin)).init App.path & a, Form1.hdc, h, True, 1
                    i = Val(s)
                    coin(UBound(coin)).Left = i
                    Input #1, i
                    coin(UBound(coin)).Top = i
                    End If
                Case 7:
                Set Player = New Character
                Input #1, i
                Player.init App.path & s, Form1.hdc, i, True, 1
                Input #1, s
                playerName = UCase(s)
                Case 8:
                    If a = "" Then
                    a = s
                    Input #1, h
                    Input #1, s
                    End If
                    Set baddy(UBound(baddy)) = New Character
                    baddy(UBound(baddy)).init App.path & a, Form1.hdc, h, True, 1
                    baddy(UBound(baddy)).Left = Val(s)
                    Input #1, i
                    baddy(UBound(baddy)).Top = i
                    ReDim Preserve baddy(0 To UBound(baddy) + 1) As Character
                    ReDim Preserve baddyPause(0 To UBound(baddyPause) + 1) As Integer
                Case 9:
                FinishLine = Val(s)
                End Select
            End Select
            If EOF(1) = False Then
            Input #1, s
            Else
            b = True
            End If
        Loop
        
    Close #1
End Sub


Sub doBaddy(Optional resetVars As Boolean = False)
Dim i As Integer

Static xv() As Integer
Static C As Integer
Static yv() As Single
Static redimedYet As Boolean
Dim inAir As Boolean
Dim j As Integer
Dim tried As Boolean

    If resetVars = True Then
    redimedYet = False
    C = 0
    Exit Sub
    End If

    If redimedYet = False Then
    redimedYet = True
    ReDim xv(0 To UBound(baddy))
    ReDim yv(0 To UBound(baddy))
    End If
C = C + 1

For i = 0 To UBound(baddy) - 1
If baddy(i).Visible = False Then GoTo nextPlease
If xv(i) = 0 Then xv(i) = 4
tried = False
baddy(i).speachText = ""
If baddyPause(i) > 0 Then
baddyPause(i) = baddyPause(i) + 1
    If baddyPause(i) = 20 Then baddy(i).index = 0
    If baddyPause(i) = 40 Then baddy(i).index = 5
    If baddyPause(i) = 60 Then baddy(i).index = 4
    If baddyPause(i) = 80 Then baddyPause(i) = 0
    If baddyPause(i) = 92 Then baddy(i).index = 9
    If baddyPause(i) = 95 Then baddy(i).index = 10
    If baddyPause(i) = 98 Then baddy(i).index = 11
    If baddyPause(i) = 120 Then baddy(i).Visible = False
    If baddyPause(i) < 90 Then baddy(i).speachText = "Erghm.."
Else
tryAgain:
    baddy(i).Left = baddy(i).Left + xv(i)

    If C Mod 9 = 0 Then baddy(i).index = baddy(i).index + 1
    If xv(i) > 0 And baddy(i).index >= 4 Then baddy(i).index = 0
    If xv(i) < 0 And baddy(i).index >= 9 Then baddy(i).index = 5
    
    If Rnd < 0.01 Then
    baddyPause(i) = 1
    baddy(i).index = 4
    End If
End If

yv(i) = yv(i) + 2.5 'force of gravity
If yv(i) > 20 Then yv(i) = 20 'terminal velocity



inAir = True

For j = 0 To UBound(road) - 1
    If baddy(i).LeftF + baddy(i).WidthF > road(j).Left And baddy(i).LeftF < road(j).Left + road(j).width Then
        If baddy(i).Top + baddy(i).height <= road(j).Top And baddy(i).height + baddy(i).Top + yv(i) >= road(j).Top Then
        yv(i) = 0
        inAir = False
         baddy(i).Top = road(j).Top - baddy(i).height
        End If
    End If
Next j

For j = 0 To UBound(block) - 1
If block(j).Visible = True Then
    If baddy(i).LeftF + baddy(i).WidthF > block(j).Left And baddy(i).LeftF < block(j).Left + block(j).width Then
        If baddy(i).Top + baddy(i).height <= block(j).Top And baddy(i).height + baddy(i).Top + yv(i) >= block(j).Top Then
        yv(i) = 0
        inAir = False
        baddy(i).Top = block(j).Top - baddy(i).height
        End If
    End If
End If
Next j

If inAir = True And tried = False Then
xv(i) = -xv(i)
If xv(i) < 0 Then baddy(i).index = 5
If xv(i) > 0 Then baddy(i).index = 0
tried = True
GoTo tryAgain
End If

baddy(i).Top = baddy(i).Top + yv(i)

If baddy(i).Top > Form1.height Then baddy(i).Visible = False
nextPlease:
Next i
End Sub


Sub deathMessage()
Dim myF As Long
myF = CreateFont(40, 20, 0, 0, 600, False, False, False, 1, 0, 0, 1, 0, "")
SelectObject Form1.hdc, myF
SetTextColor Form1.hdc, RGB(50, 0, 0)
SetTextAlign Form1.hdc, 6
TextOut Form1.hdc, Form1.ScaleWidth / 2, 80, "YOU KILLED " & playerName, Len("YOU KILLED " & playerName)
Form1.Refresh
Sleep 1000

SetTextColor Form1.hdc, vbRed
TextOut Form1.hdc, Form1.ScaleWidth / 2, 180, "YOU BASTARD!", Len("YOU BASTARD!")
Form1.Refresh
Sleep 1000
DeleteObject myF
SetTextColor Form1.hdc, vbBlack
End Sub
