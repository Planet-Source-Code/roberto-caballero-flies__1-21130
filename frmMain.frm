VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flies"
   ClientHeight    =   7635
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton About 
      Caption         =   "Ab&out"
      Height          =   495
      Left            =   9960
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton ClearScreen 
      Caption         =   "Cle&ar screen"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer TimerRandomDot 
      Interval        =   5000
      Left            =   10560
      Top             =   6000
   End
   Begin VB.Timer TimerEraseTrails 
      Interval        =   5000
      Left            =   10080
      Top             =   6000
   End
   Begin VB.Frame FrameNumFlies 
      Caption         =   "Number of flies"
      Height          =   855
      Left            =   480
      TabIndex        =   23
      Top             =   5520
      Width           =   2535
      Begin VB.HScrollBar ScrollNumFlies 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   300
         Min             =   3
         TabIndex        =   3
         Top             =   360
         Value           =   3
         Width           =   2295
      End
   End
   Begin VB.Frame FrameTrails 
      Caption         =   "Trails"
      Height          =   855
      Left            =   480
      TabIndex        =   22
      Top             =   6480
      Width           =   2295
      Begin VB.CheckBox CheckEraseTrails 
         Caption         =   "Erase trails e&very 5 secs."
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox CheckLeaveTrails 
         Caption         =   "L&eave trails"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FrameSpeed 
      Caption         =   "Speed"
      Height          =   855
      Left            =   6240
      TabIndex        =   21
      Top             =   6480
      Width           =   2415
      Begin VB.HScrollBar ScrollSpeed 
         Height          =   255
         LargeChange     =   2
         Left            =   120
         Max             =   1
         Min             =   20
         TabIndex        =   15
         Top             =   360
         Value           =   1
         Width           =   2175
      End
   End
   Begin VB.Frame FrameFliesFollow 
      Caption         =   "The flies follow"
      Height          =   1095
      Left            =   3240
      TabIndex        =   20
      Top             =   6480
      Width           =   2295
      Begin VB.OptionButton OptIndDot 
         Caption         =   "&Individual random dot"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton OptMouse 
         Caption         =   "&Mouse pointer"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton OptCommonDot 
         Caption         =   "Common &random dot"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame FrameShape 
      Caption         =   "Shape of the flies"
      Height          =   855
      Left            =   3480
      TabIndex        =   19
      Top             =   5520
      Width           =   1695
      Begin VB.OptionButton OptLine 
         Caption         =   "&Line"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptTriangle 
         Caption         =   "&Triangle"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame FrameColor 
      Caption         =   "Color of the flies and color of the background"
      Height          =   855
      Left            =   5640
      TabIndex        =   18
      Top             =   5520
      Width           =   3735
      Begin VB.OptionButton OptBlackOnWhite 
         Caption         =   "&Black on White"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton OptColorOnWhite 
         Caption         =   "&Colored on White"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptWhiteOnBlack 
         Caption         =   "&White on Black"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton OptColorOnBlack 
         Caption         =   "&Colored on Black"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   9960
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox FlyArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5175
      Left            =   240
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   629
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   9495
   End
   Begin VB.CommandButton StartAndPause 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   9960
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumFlies As Integer  'Number of flies
Dim RepelDistance As Integer 'Minimum distance squared that the flies can get to without repelling
Dim Acc As Integer 'Acceleration amount
Dim Acc2Center As Integer 'Acceleration towards the center point
Dim MaxSpeed As Integer
Dim BounceSpeed As Integer
Private Type TypeFlies
   X As Integer
   Y As Integer
   VX As Integer 'VX & VY are the velocities in x and y directions
   VY As Integer
   CenterX As Integer
   CenterY As Integer
   LeftNeighbor As Integer
   RightNeighbor As Integer
   lpX As Integer 'lpX & lpY are the last positions of the fly for drawing purposes
   lpY As Integer
   mColor As Long
End Type
Const MaxNumFlies = 300
Dim Fly(1 To MaxNumFlies) As TypeFlies
Dim mWidth As Integer
Dim mHeight As Integer
Dim SystemState As Byte '1=Not yet started, 2=Running, 3=Paused
Public Function Rand(Low As Integer, High As Integer) As Integer
   Dim fake As Integer
   fake = 0
   While fake = 0
      fake = Int((High - Low + 1) * Rnd + Low)
   Wend
   Rand = fake
End Function
Public Sub Init()
   Dim i As Integer
   Dim mCX As Integer
   Dim mCY As Integer
   mWidth = FlyArea.ScaleWidth
   mHeight = FlyArea.ScaleHeight
   mCX = Rand(10, mWidth - 10)
   mCY = Rand(10, mHeight - 10)
   For i = 1 To NumFlies
      Fly(i).X = Rand(10, mWidth - 10)
      Fly(i).lpX = Rand(10, mWidth - 10)
      Fly(i).VX = Rand(10, 100)
      Fly(i).Y = Rand(10, mHeight - 10)
      Fly(i).lpY = Rand(10, mHeight - 10)
      Fly(i).VY = Rand(10, 100)
      Fly(i).CenterX = mCX
      Fly(i).CenterY = mCY
   Next
   RandNeighbors
End Sub
Public Sub ProcessFly(Tick As Integer)
   ' Rev determines if the neighbor attracts or repels
   ' Tick is the fly number being processed
   Dim Rev As Integer
   Fly(Tick).lpX = Fly(Tick).X
   Fly(Tick).lpY = Fly(Tick).Y
   Rev = 1
   If Dist(Tick, Fly(Tick).LeftNeighbor) < RepelDistance Then Rev = -1
   If Fly(Tick).X < Fly(Fly(Tick).LeftNeighbor).X Then
      Fly(Tick).VX = Fly(Tick).VX + Acc * Rev
   Else
      Fly(Tick).VX = Fly(Tick).VX - Acc * Rev
   End If
   If Fly(Tick).Y < Fly(Fly(Tick).LeftNeighbor).Y Then
      Fly(Tick).VY = Fly(Tick).VY + Acc * Rev
   Else
      Fly(Tick).VY = Fly(Tick).VY - Acc * Rev
   End If
   Rev = 1
   If Dist(Tick, Fly(Tick).RightNeighbor) < RepelDistance Then Rev = -1
   If Fly(Tick).X < Fly(Fly(Tick).RightNeighbor).X Then
      Fly(Tick).VX = Fly(Tick).VX + Acc * Rev
   Else
      Fly(Tick).VX = Fly(Tick).VX - Acc * Rev
   End If
   If Fly(Tick).Y < Fly(Fly(Tick).RightNeighbor).Y Then
      Fly(Tick).VY = Fly(Tick).VY + Acc * Rev
   Else
      Fly(Tick).VY = Fly(Tick).VY - Acc * Rev
   End If
   If Fly(Tick).VX > MaxSpeed Then Fly(Tick).VX = MaxSpeed
   If Fly(Tick).VX < -MaxSpeed Then Fly(Tick).VX = -MaxSpeed
   If Fly(Tick).VY > MaxSpeed Then Fly(Tick).VY = MaxSpeed
   If Fly(Tick).VY < -MaxSpeed Then Fly(Tick).VY = -MaxSpeed
   If Fly(Tick).X < 0 Then Fly(Tick).VX = BounceSpeed
   If Fly(Tick).X > mWidth Then Fly(Tick).VX = -BounceSpeed
   If Fly(Tick).Y < 0 Then Fly(Tick).VY = BounceSpeed
   If Fly(Tick).Y > mHeight Then Fly(Tick).VY = -BounceSpeed
   If Fly(Tick).X < Fly(Tick).CenterX Then
      Fly(Tick).VX = Fly(Tick).VX + Acc2Center
   Else
      Fly(Tick).VX = Fly(Tick).VX - Acc2Center
   End If
   If Fly(Tick).Y < Fly(Tick).CenterY Then
      Fly(Tick).VY = Fly(Tick).VY + Acc2Center
   Else
      Fly(Tick).VY = Fly(Tick).VY - Acc2Center
   End If
   Fly(Tick).X = Fly(Tick).X + Fly(Tick).VX
   Fly(Tick).Y = Fly(Tick).Y + Fly(Tick).VY
End Sub
Public Sub DoNeighbors(Tick As Integer)
   ' This re-assigns neighbors if the neighbor's neighbors are closer to
   ' the fly than the neighbors.
   Dim m As Integer

   m = Fly(Tick).LeftNeighbor
   If Dist(Fly(m).LeftNeighbor, Tick) < Dist(Fly(Tick).LeftNeighbor, Tick) Then
      If Fly(m).LeftNeighbor <> Fly(Tick).RightNeighbor Then
         Fly(Tick).LeftNeighbor = Fly(m).LeftNeighbor
      End If
   End If
   If Dist(Fly(m).LeftNeighbor, Tick) < Dist(Fly(Tick).RightNeighbor, Tick) Then
      If Fly(m).LeftNeighbor <> Fly(Tick).LeftNeighbor Then
         Fly(Tick).RightNeighbor = Fly(m).LeftNeighbor
      End If
   End If
   
   If Dist(Fly(m).RightNeighbor, Tick) < Dist(Fly(Tick).LeftNeighbor, Tick) Then
      If Fly(m).RightNeighbor <> Fly(Tick).RightNeighbor Then
         Fly(Tick).LeftNeighbor = Fly(m).RightNeighbor
      End If
   End If
   If Dist(Fly(m).RightNeighbor, Tick) < Dist(Fly(Tick).RightNeighbor, Tick) Then
      If Fly(m).RightNeighbor <> Fly(Tick).LeftNeighbor Then
         Fly(Tick).RightNeighbor = Fly(m).RightNeighbor
      End If
   End If
      
   m = Fly(Tick).RightNeighbor
   If Dist(Fly(m).LeftNeighbor, Tick) < Dist(Fly(Tick).LeftNeighbor, Tick) Then
      If Fly(m).LeftNeighbor <> Fly(Tick).RightNeighbor Then
         Fly(Tick).LeftNeighbor = Fly(m).LeftNeighbor
      End If
   End If
   If Dist(Fly(m).LeftNeighbor, Tick) < Dist(Fly(Tick).RightNeighbor, Tick) Then
      If Fly(m).LeftNeighbor <> Fly(Tick).LeftNeighbor Then
         Fly(Tick).RightNeighbor = Fly(m).LeftNeighbor
      End If
   End If
   
   If Dist(Fly(m).RightNeighbor, Tick) < Dist(Fly(Tick).LeftNeighbor, Tick) Then
      If Fly(m).RightNeighbor <> Fly(Tick).RightNeighbor Then
         Fly(Tick).LeftNeighbor = Fly(m).RightNeighbor
      End If
   End If
   If Dist(Fly(m).RightNeighbor, Tick) < Dist(Fly(Tick).RightNeighbor, Tick) Then
      If Fly(m).RightNeighbor <> Fly(Tick).LeftNeighbor Then
         Fly(Tick).RightNeighbor = Fly(m).RightNeighbor
      End If
   End If
   
   ' This bit randomises the fly's neighbor if it has been assigned itself as a neighbor
   If Fly(Tick).LeftNeighbor = Tick Then
      Fly(Tick).LeftNeighbor = Rand(1, NumFlies)
   End If
   If Fly(Tick).RightNeighbor = Tick Then
      Fly(Tick).RightNeighbor = Rand(1, NumFlies)
   End If
End Sub
Public Function Dist(Tick1 As Integer, Tick2 As Integer) As Double
   ' Distance between flies
   Dim d1 As Double
   Dim d2 As Double
   Dim a As Double
   Dim b As Double
   d1 = Fly(Tick1).X - Fly(Tick2).X
   d2 = Fly(Tick1).Y - Fly(Tick2).Y
   a = d1 * d1
   b = d2 * d2
   Dist = Int(Sqr(a + b))
End Function
Public Sub RandNeighbors()
   ' Randomises the neighbors
   Dim k As Integer
   For k = 1 To NumFlies
      Fly(k).LeftNeighbor = Rand(1, NumFlies)
      Fly(k).RightNeighbor = Rand(1, NumFlies)
   Next
End Sub
Public Sub Run()
   Dim i As Integer
   Dim j As Long
   While True
      If SystemState = 2 Then
         For i = 1 To NumFlies
            DoNeighbors (i)
            ProcessFly (i)
         Next
         If Rand(1, 10) > 8 Then RandNeighbors
         DrawFlies
         DoEvents
         For j = 1 To ScrollSpeed.Value * 150000: Next
      End If
      DoEvents
   Wend
End Sub
Public Sub DrawFlies()
   Dim k As Integer
   Dim mDrawColor As Byte '1=Colored, 2=Black, 3=White
   If OptColorOnBlack Or OptColorOnWhite Then mDrawColor = 1
   If OptBlackOnWhite Then mDrawColor = 2
   If OptWhiteOnBlack Then mDrawColor = 3
   If CheckLeaveTrails.Value = 0 Then FlyArea.Cls
   For k = 1 To NumFlies
      If OptTriangle Then
         FlyArea.Line (Fly(k).lpX + 5, Fly(k).lpY)-(Fly(k).X, Fly(k).Y), IIf(mDrawColor = 1, Fly(k).mColor, IIf(mDrawColor = 2, vbBlack, vbWhite))
         FlyArea.Line (Fly(k).lpX - 5, Fly(k).lpY)-(Fly(k).X, Fly(k).Y), IIf(mDrawColor = 1, Fly(k).mColor, IIf(mDrawColor = 2, vbBlack, vbWhite))
         FlyArea.Line (Fly(k).lpX - 5, Fly(k).lpY)-(Fly(k).lpX + 5, Fly(k).lpY), IIf(mDrawColor = 1, Fly(k).mColor, IIf(mDrawColor = 2, vbBlack, vbWhite))
      End If
      If OptLine Then
         FlyArea.Line (Fly(k).lpX, Fly(k).lpY)-(Fly(k).X, Fly(k).Y), IIf(mDrawColor = 1, Fly(k).mColor, IIf(mDrawColor = 2, vbBlack, vbWhite))
      End If
   Next
End Sub

Private Sub About_Click()
   If SystemState = 2 Then frmMain.Caption = "Flies - Paused"
   frmAbout.Show 1
   If SystemState = 1 Then frmMain.Caption = "Flies"
   If SystemState = 2 Then frmMain.Caption = "Flies - Running ..."
End Sub

Private Sub CheckLeaveTrails_Click()
   If CheckLeaveTrails.Value = 1 Then
      CheckEraseTrails.Enabled = True
      If SystemState = 2 Then ClearScreen.Enabled = True
   Else
      CheckEraseTrails.Enabled = False
      CheckEraseTrails.Value = 0
      ClearScreen.Enabled = False
   End If
End Sub
Private Sub FlyArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Integer
   If OptMouse Then
      For i = 1 To NumFlies
         Fly(i).CenterX = X
         Fly(i).CenterY = Y
      Next
   End If
End Sub
Private Sub Form_Load()
   Dim i As Integer
   Randomize Timer
   NumFlies = 100 '100
   Acc = 1 '1
   RepelDistance = 50 '50
   Acc2Center = 1 '1
   BounceSpeed = 3 '3
   MaxSpeed = 10 '10
   For i = 1 To MaxNumFlies
      Fly(i).mColor = RGB(Rand(100, 255), Rand(100, 255), Rand(100, 255))
   Next
   ScrollNumFlies.Value = NumFlies
   ScrollNumFlies.Max = MaxNumFlies
   FrameNumFlies.Caption = "Number of flies (3 to " & MaxNumFlies & "): " & ScrollNumFlies.Value
   FrameSpeed.Caption = "Speed (1 to 20): " & 21 - ScrollSpeed.Value
   SystemState = 1 'Not yet started
End Sub
Private Sub OptIndDot_Click()
   Dim i As Integer
   For i = 1 To NumFlies
      Fly(i).CenterX = Rand(10, mWidth - 10)
      Fly(i).CenterY = Rand(10, mHeight - 10)
   Next
End Sub
Private Sub ScrollNumFlies_Scroll()
   FrameNumFlies.Caption = "Number of flies (3 to " & MaxNumFlies & "): " & ScrollNumFlies.Value
   NumFlies = ScrollNumFlies.Value
End Sub
Private Sub ScrollSpeed_Change()
   FrameSpeed.Caption = "Speed (1 to 20): " & 21 - ScrollSpeed.Value
End Sub
Private Sub ScrollSpeed_Scroll()
   FrameSpeed.Caption = "Speed (1 to 20): " & 21 - ScrollSpeed.Value
End Sub
Private Sub StartAndPause_Click()
   If SystemState = 1 Then
      StartAndPause.Caption = "&Pause"
      frmMain.Caption = "Flies - Running ..."
      SystemState = 2
      Init
      Run
      Exit Sub
   End If
   If StartAndPause.Caption = "&Start" Then
      StartAndPause.Caption = "&Pause"
      frmMain.Caption = "Flies - Running ..."
      SystemState = 2
      TimerEraseTrails.Enabled = True
      If CheckLeaveTrails.Value = 1 Then
         ClearScreen.Enabled = True
      Else
         ClearScreen.Enabled = False
      End If
   Else
      StartAndPause.Caption = "&Start"
      frmMain.Caption = "Flies - Paused"
      SystemState = 3
      TimerEraseTrails.Enabled = False
      ClearScreen.Enabled = False
   End If
End Sub
Private Sub TimerEraseTrails_Timer()
   If CheckEraseTrails.Value = 1 Then FlyArea.Cls
End Sub
Private Sub TimerRandomDot_Timer()
   Dim i As Integer
   Dim mCX As Integer
   Dim mCY As Integer
   mCX = Rand(15, mWidth - 15)
   mCY = Rand(15, mHeight - 15)
   For i = 1 To NumFlies
      If OptIndDot Then
         Fly(i).CenterX = Rand(15, mWidth - 15)
         Fly(i).CenterY = Rand(15, mHeight - 15)
      End If
      If OptCommonDot Then
         Fly(i).CenterX = mCX
         Fly(i).CenterY = mCY
      End If
   Next
End Sub
Private Sub OptCommonDot_Click()
   Dim i As Integer
   Dim mCX As Integer
   Dim mCY As Integer
   mCX = Rand(15, mWidth - 15)
   mCY = Rand(15, mHeight - 15)
   For i = 1 To NumFlies
      Fly(i).CenterX = mCX
      Fly(i).CenterY = mCY
   Next
End Sub
Private Sub ScrollNumFlies_Change()
   Dim i As Integer
   FrameNumFlies.Caption = "Number of flies (3 to " & MaxNumFlies & "): " & ScrollNumFlies.Value
   NumFlies = ScrollNumFlies.Value
   RandNeighbors
   For i = NumFlies + 1 To MaxNumFlies
      Fly(i).X = 0
      Fly(i).Y = 0
   Next
End Sub
Private Sub OptLine_Click()
   FlyArea.Cls
   If SystemState = 3 Then DrawFlies
End Sub
Private Sub OptTriangle_Click()
   FlyArea.Cls
   If SystemState = 3 Then DrawFlies
End Sub
Private Sub OptBlackOnWhite_Click()
   FlyArea.BackColor = vbWhite
   If SystemState = 3 Then DrawFlies
End Sub
Private Sub OptColorOnBlack_Click()
   FlyArea.BackColor = vbBlack
   If SystemState = 3 Then DrawFlies
End Sub
Private Sub OptColorOnWhite_Click()
   FlyArea.BackColor = vbWhite
   If SystemState = 3 Then DrawFlies
End Sub
Private Sub OptWhiteOnBlack_Click()
   FlyArea.BackColor = vbBlack
   If SystemState = 3 Then DrawFlies
End Sub
Private Sub ClearScreen_Click()
   FlyArea.Cls
End Sub
Private Sub Exit_Click()
   End
End Sub
