VERSION 5.00
Begin VB.Form FThreeBowls 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Three Bowls"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "FThreeBowls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBet 
      Caption         =   "Bet $25"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Tag             =   "25"
      Top             =   2580
      Width           =   915
   End
   Begin VB.CommandButton cmdBet 
      Caption         =   "Bet $10"
      Height          =   315
      Index           =   1
      Left            =   2340
      TabIndex        =   5
      Tag             =   "10"
      Top             =   2580
      Width           =   915
   End
   Begin VB.CommandButton cmdBet 
      Caption         =   "Bet $5"
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Tag             =   "5"
      Top             =   2580
      Width           =   915
   End
   Begin VB.TextBox txtSaving 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "10"
      Top             =   2580
      Width           =   615
   End
   Begin VB.PictureBox pctShow 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   0
      MouseIcon       =   "FThreeBowls.frx":0442
      ScaleHeight     =   2415
      ScaleWidth      =   4185
      TabIndex        =   1
      Top             =   0
      Width           =   4245
   End
   Begin VB.PictureBox pctBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C0C0C0&
      Height          =   2835
      Left            =   4560
      ScaleHeight     =   2775
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   2820
      Width           =   4275
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   7
      Left            =   3180
      Picture         =   "FThreeBowls.frx":074C
      Top             =   4860
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   6
      Left            =   2220
      Picture         =   "FThreeBowls.frx":15A2
      Top             =   4860
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   5
      Left            =   1260
      Picture         =   "FThreeBowls.frx":23F8
      Top             =   4860
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   4
      Left            =   240
      Picture         =   "FThreeBowls.frx":324E
      Top             =   4860
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   3
      Left            =   3120
      Picture         =   "FThreeBowls.frx":40A4
      Top             =   4140
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   2
      Left            =   2220
      Picture         =   "FThreeBowls.frx":4EFA
      Top             =   4140
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   1
      Left            =   1320
      Picture         =   "FThreeBowls.frx":5D50
      Top             =   4140
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Saving"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image imgRealObject 
      Height          =   240
      Left            =   2700
      Picture         =   "FThreeBowls.frx":6BA6
      Top             =   3720
      Width           =   240
   End
   Begin VB.Image imgObjectMask 
      Height          =   240
      Left            =   2700
      Picture         =   "FThreeBowls.frx":70E8
      Top             =   3360
      Width           =   240
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   0
      Left            =   300
      Picture         =   "FThreeBowls.frx":71EA
      Top             =   4140
      Width           =   855
   End
   Begin VB.Image imgMask 
      Height          =   645
      Left            =   480
      Picture         =   "FThreeBowls.frx":8040
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "FThreeBowls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type CoverObject
  Top As Integer
  Left As Integer
  Possition As Integer
  dX As Integer
  Length As Integer
End Type

Private Cover(2) As CoverObject
Private GameLevel As Integer
Private OnBet As Boolean
Private BetValue As Integer
Private OnTheShow As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Change(Index1 As Integer, Index2 As Integer, Index3 As Integer, ExDelay As Integer)
Dim dX As Integer
Dim dY As Integer
Dim X As Integer
Dim Y1 As Integer
Dim Y2 As Integer

Dim strPattern As String
Dim Length As Integer
Dim Divider As Integer
Dim MoveMode As Integer

Dim Cheat As Boolean
Dim CheatMode As Integer
Dim Index() As Integer

Dim I As Integer
Dim J As Integer
Dim K As Integer

  ReDim Index(UBound(Cover))
  If Cover(Index1).Left > Cover(Index2).Left Then
    Length = Cover(Index1).Left - Cover(Index2).Left
    dX = -1
  Else
    Length = Cover(Index2).Left - Cover(Index1).Left
    dX = 1
  End If
  
  If Rnd > 0.5 Then
    dY = -1
  Else
    dY = 1
  End If
  
  If Abs(Cover(Index1).Left - Cover(Index2).Left) > 2000 Then
    Divider = 400
    MoveMode = 0
  Else
    Divider = 240
    MoveMode = CInt(Rnd * 3)
  End If
  
  Cheat = (Rnd > 0.8)
  If Cheat Then
    CheatMode = CInt(Rnd * 3)
  End If
  
  For X = 1 To Length \ Divider
    If X <= (Length \ Divider) / 2 Then
      If X < 2 Then
        strPattern = strPattern & X
      Else
        strPattern = strPattern & 3
      End If
    Else
      If (((Length \ Divider)) - X) < 2 Then
        strPattern = strPattern & (((Length \ Divider)) - X)
      Else
        strPattern = strPattern & 3
      End If
    End If
  Next X
  
  Y1 = Cover(Index1).Top
  Y2 = Cover(Index1).Top
  
  For X = 1 To Length \ Divider
    Cover(Index1).Left = Cover(Index1).Left + dX * Divider
    Select Case MoveMode
      Case 0, 1, 3
        Cover(Index1).Top = Y1 + dY * (Mid(strPattern, X, 1) * 150)
    End Select
    
    Cover(Index2).Left = Cover(Index2).Left - dX * Divider
    Select Case MoveMode
      Case 0, 2, 3
        Cover(Index2).Top = Y1 - dY * (Mid(strPattern, X, 1) * 180)
    End Select
    
    For I = 0 To UBound(Cover)
      Index(I) = I
    Next I
    For I = 0 To 2
      For J = I + 1 To 2
        If Cover(Index(I)).Top > Cover(Index(J)).Top Then
          K = Index(I)
          Index(I) = Index(J)
          Index(J) = K
        End If
      Next J
    Next I
          
    pctBack.Cls
    For I = 0 To 2
      pctBack.PaintPicture imgMask.Picture, Cover(Index(I)).Left, Cover(Index(I)).Top, opcode:=vbSrcAnd
      If Index(I) <> Index3 Then
        If Cover(Index(I)).dX = 0 Then
          If Rnd > 0.5 Then
            Cover(Index(I)).dX = 1
          Else
            Cover(Index(I)).dX = -1
          End If
          Cover(Index(I)).Length = Int(Rnd * 8) + 2
        End If
        
        If Cover(Index(I)).dX = 1 Then
          Cover(Index(I)).Possition = (Cover(Index(I)).Possition + 1) Mod 8
        Else
          Cover(Index(I)).Possition = (Cover(Index(I)).Possition - 1 + 8) Mod 8
        End If
        Cover(Index(I)).Length = Cover(Index(I)).Length - 1
        
        If Cover(Index(I)).Length = 0 Then
          If Rnd > 0.5 Then
            Cover(Index(I)).dX = 1
          Else
            Cover(Index(I)).dX = -1
          End If
          Cover(Index(I)).Length = Int(Rnd * 8) + 2
        End If
      End If
      pctBack.PaintPicture imgRealCover(Cover(Index(I)).Possition).Picture, Cover(Index(I)).Left, Cover(Index(I)).Top, opcode:=vbSrcPaint
    Next I
    pctShow.PaintPicture pctBack.Image, 0, 0, opcode:=vbSrcCopy
    Sleep ExDelay
    DoEvents
  Next X
End Sub

Private Sub OpenCover(Index As Integer)
Dim I As Integer
Dim Y As Integer
Dim Y1 As Integer
Dim X As Integer
Dim strPattern As String

Dim Index1 As Integer
Dim Index2 As Integer

  Index1 = 0
  If Index1 = Index Then
    Index1 = 1
    Index2 = 2
  ElseIf Index = 1 Then
    Index2 = 2
  Else
    Index2 = 1
  End If
    
  strPattern = "12222210"
  Y = Cover(Index).Top
  
  For I = 1 To Len(strPattern)
    Cover(Index).Top = Y - (Mid(strPattern, I, 1) * 240)
    DoEvents
    
    pctBack.Cls
    pctBack.Circle (Cover(Index).Left + Me.imgMask.Width / 2 + Mid(strPattern, I, 1) * 30, Y + Me.imgMask.Height - 200 + Mid(strPattern, I, 1) * 30), imgMask.Width \ 2 + Mid(strPattern, I, 1) * 30, , , , 0.4
    If Index = 0 Then
      pctBack.PaintPicture imgObjectMask.Picture, Cover(Index).Left + 100, Y + 360, opcode:=vbSrcAnd
      pctBack.PaintPicture imgRealObject.Picture, Cover(Index).Left + 100, Y + 360, opcode:=vbSrcPaint
    End If
    
    pctBack.PaintPicture imgMask.Picture, Cover(Index1).Left, Cover(Index1).Top, opcode:=vbSrcAnd
    pctBack.PaintPicture imgRealCover(Cover(Index1).Possition).Picture, Cover(Index1).Left, Cover(Index1).Top, opcode:=vbSrcPaint
    
    pctBack.PaintPicture imgMask.Picture, Cover(Index2).Left, Cover(Index2).Top, opcode:=vbSrcAnd
    pctBack.PaintPicture imgRealCover(Cover(Index2).Possition).Picture, Cover(Index2).Left, Cover(Index2).Top, opcode:=vbSrcPaint
    
    pctBack.PaintPicture imgMask.Picture, Cover(Index).Left, Cover(Index).Top, opcode:=vbSrcAnd
    pctBack.PaintPicture imgRealCover(Cover(Index).Possition).Picture, Cover(Index).Left, Cover(Index).Top, opcode:=vbSrcPaint
    
    pctShow.PaintPicture pctBack.Image, 0, 0, opcode:=vbSrcCopy
    Sleep 40
  Next I
End Sub

Private Sub Play(ByVal Level As Integer)
Dim Possition1 As Integer
Dim Possition2 As Integer
Dim I As Integer
Dim Possition3 As Integer
Dim Pattern As Single
Dim Delay As Integer
Dim strTemp As String
Dim ExtDelay As Integer

  OnTheShow = True
  DoEvents
  
  If Level >= 5 Then
    Delay = 20
    ExtDelay = 40 - (Level - 4) * 5
  Else
    ExtDelay = 40
    Delay = 800 - ((Level + 1) * 150)
  End If
  
  For I = 1 To Level + 3
    GetRandomPossition Possition1, Possition2, Possition3
    
    Pattern = Rnd
    Change Possition1, Possition2, Possition3, ExtDelay
    If Pattern > 0.2 Then Change Possition3, Possition1, Possition2, ExtDelay
    If Pattern > 0.4 Then Change Possition2, Possition3, Possition1, ExtDelay
    If Pattern > 0.6 Then Change Possition1, Possition3, Possition2, ExtDelay
    If Pattern > 0.8 Then Change Possition3, Possition2, Possition1, ExtDelay
    If Pattern > 0.95 Then Change Possition2, Possition1, Possition3, ExtDelay
    Sleep Delay
  Next
  pctShow.MousePointer = vbDefault
  OnTheShow = False
End Sub

Private Sub cmdBet_Click(Index As Integer)
Dim I As Integer
Dim Possition1 As Integer
Dim Possition2 As Integer
Dim Possition3 As Integer

  Enabled = False
  
  GetRandomPossition Possition1, Possition2, Possition3
  Change Possition1, Possition2, Possition3, 50
  OpenCover 0
  OnBet = True
  For I = 0 To cmdBet.Count - 1
    cmdBet(I).Enabled = False
  Next I
  
  BetValue = cmdBet(Index).Tag
  txtSaving = txtSaving - BetValue
  Caption = "Three bowls Level: " & GameLevel
  DoEvents
  
  Play GameLevel
  Enabled = True
End Sub

Private Sub Form_Load()
Dim I As Integer

  Randomize
  For I = 0 To UBound(Cover)
    Cover(I).Left = 480 + 1200 * I
    Cover(I).Top = 960
    If Rnd > 0.5 Then
      Cover(I).dX = -1
    Else
      Cover(I).dX = 1
    End If
    Cover(I).Length = Int(Rnd * 8) + 2
    Cover(I).Possition = Int(Rnd * 8)
  Next I
  
  OpenCover 0
End Sub

Private Sub pctShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If OnTheShow Then
    pctShow.MousePointer = vbCustom
  End If
End Sub

Private Sub pctShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
Dim Found As Boolean
Static LoseCount As Integer
Dim ClickOnBowl As Boolean

  If Not OnBet Then Exit Sub
  Enabled = False
  OnBet = False
  ClickOnBowl = False
  Found = False
  
  For I = 0 To UBound(Cover)
    If X >= Cover(I).Left And X <= Cover(I).Left + Me.imgMask.Width And _
       Y >= Cover(I).Top And Y <= Cover(I).Top + Me.imgMask.Height Then
      ClickOnBowl = True
      Found = (I = 0)
      OpenCover I
      Enabled = True
      Exit For
    End If
  Next I
  
  If ClickOnBowl Then
    If Not Found Then
      OpenCover 0
      LoseCount = LoseCount + 1
      If LoseCount > 3 Then
        If GameLevel <> 0 Then GameLevel = GameLevel - 1
        LoseCount = 0
      End If
      
    Else
      txtSaving = txtSaving + BetValue * 2 + 3 * GameLevel
      GameLevel = GameLevel + 1
      LoseCount = 0
    End If
    For I = 0 To cmdBet.Count - 1
      cmdBet(I).Enabled = (Val(txtSaving) >= Val(cmdBet(I).Tag))
    Next I
  
  Else
    OnBet = True
  End If
  Enabled = True
End Sub

Private Sub GetRandomPossition(Possition1 As Integer, _
                               Possition2 As Integer, _
                               Possition3 As Integer)
Dim strTemp As String

  strTemp = "012"
  Possition1 = Int(Rnd * 3)
  Do
    Possition2 = Int(Rnd * 3)
  Loop Until Possition2 <> Possition1
  
  Mid(strTemp, InStr(strTemp, Possition1), 1) = " "
  Mid(strTemp, InStr(strTemp, Possition2), 1) = " "
  Possition3 = Val(strTemp)
End Sub
