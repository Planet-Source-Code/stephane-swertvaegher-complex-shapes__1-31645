VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CDFig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complex Shapes"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   2970
   Icon            =   "CDFig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   Begin VB.Frame Frame2 
      Caption         =   "Only on repeating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2445
      Left            =   0
      TabIndex        =   39
      Top             =   5130
      Width           =   3030
      Begin VB.CheckBox Check1 
         Caption         =   "Symmetric repeats"
         Height          =   240
         Left            =   90
         TabIndex        =   58
         Top             =   2115
         Width           =   1725
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   8
         LargeChange     =   5
         Left            =   1080
         Max             =   50
         Min             =   -50
         TabIndex        =   45
         Top             =   360
         Width           =   1455
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   9
         LargeChange     =   5
         Left            =   1080
         Max             =   50
         Min             =   -50
         TabIndex        =   44
         Top             =   630
         Width           =   1455
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   10
         LargeChange     =   5
         Left            =   1080
         Max             =   50
         Min             =   -50
         TabIndex        =   43
         Top             =   900
         Width           =   1455
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   11
         LargeChange     =   5
         Left            =   1080
         Max             =   50
         Min             =   -50
         TabIndex        =   42
         Top             =   1170
         Width           =   1455
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   16
         LargeChange     =   5
         Left            =   1080
         Max             =   100
         Min             =   -100
         TabIndex        =   41
         Top             =   1575
         Width           =   1455
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   17
         LargeChange     =   5
         Left            =   1080
         Max             =   100
         Min             =   -100
         TabIndex        =   40
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   240
         Index           =   8
         Left            =   2565
         TabIndex        =   57
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add X1"
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   56
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   240
         Index           =   9
         Left            =   2565
         TabIndex        =   55
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Y1"
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   54
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   240
         Index           =   10
         Left            =   2565
         TabIndex        =   53
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add X2"
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   52
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   240
         Index           =   11
         Left            =   2565
         TabIndex        =   51
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Y2"
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   50
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   240
         Index           =   16
         Left            =   2565
         TabIndex        =   49
         Top             =   1575
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Ampl. 1"
         Height          =   285
         Index           =   16
         Left            =   90
         TabIndex        =   48
         Top             =   1530
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   240
         Index           =   17
         Left            =   2565
         TabIndex        =   47
         Top             =   1845
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Ampl. 2"
         Height          =   285
         Index           =   17
         Left            =   90
         TabIndex        =   46
         Top             =   1800
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   765
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "é"
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   15
      LargeChange     =   5
      Left            =   1080
      Max             =   30
      Min             =   -30
      TabIndex        =   36
      Top             =   4095
      Value           =   10
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   14
      LargeChange     =   5
      Left            =   1080
      Max             =   30
      Min             =   -30
      TabIndex        =   33
      Top             =   3825
      Value           =   10
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   13
      LargeChange     =   5
      Left            =   1080
      Max             =   30
      Min             =   -30
      TabIndex        =   30
      Top             =   2610
      Value           =   10
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   12
      LargeChange     =   5
      Left            =   1080
      Max             =   30
      Min             =   -30
      TabIndex        =   27
      Top             =   2340
      Value           =   10
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   7
      LargeChange     =   5
      Left            =   1080
      Max             =   150
      TabIndex        =   24
      Top             =   4770
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   6
      LargeChange     =   10
      Left            =   1080
      Max             =   600
      TabIndex        =   19
      Top             =   3555
      Value           =   300
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   5
      LargeChange     =   10
      Left            =   1080
      Max             =   800
      TabIndex        =   18
      Top             =   3285
      Value           =   400
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   4
      LargeChange     =   10
      Left            =   1080
      Max             =   600
      TabIndex        =   15
      Top             =   2070
      Value           =   300
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   3
      LargeChange     =   10
      Left            =   1080
      Max             =   800
      TabIndex        =   12
      Top             =   1800
      Value           =   400
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   2
      Left            =   1080
      Max             =   23
      TabIndex        =   9
      Top             =   4500
      Value           =   6
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   1
      LargeChange     =   10
      Left            =   1080
      Max             =   300
      Min             =   1
      TabIndex        =   6
      Top             =   3015
      Value           =   128
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show me"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   7650
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Index           =   0
      LargeChange     =   10
      Left            =   1080
      Max             =   300
      Min             =   1
      TabIndex        =   1
      Top             =   1530
      Value           =   200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1140
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   2940
      Begin VB.CommandButton Command2 
         Caption         =   "Swap colors"
         Height          =   330
         Left            =   1620
         TabIndex        =   64
         Top             =   765
         Width           =   1140
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Gradient color"
         Height          =   195
         Left            =   1530
         TabIndex        =   63
         Top             =   495
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Equal color"
         Height          =   195
         Left            =   1530
         TabIndex        =   62
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color 2"
         Height          =   240
         Index           =   1
         Left            =   495
         TabIndex        =   61
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   60
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color 1"
         Height          =   240
         Index           =   0
         Left            =   495
         TabIndex        =   59
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   2940
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   15
      Left            =   2565
      TabIndex        =   38
      Top             =   4095
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aspect Y2"
      Height          =   285
      Index           =   15
      Left            =   90
      TabIndex        =   37
      Top             =   4050
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   14
      Left            =   2565
      TabIndex        =   35
      Top             =   3825
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aspect X2"
      Height          =   285
      Index           =   14
      Left            =   90
      TabIndex        =   34
      Top             =   3780
      Width           =   960
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aspect Y1"
      Height          =   285
      Index           =   13
      Left            =   90
      TabIndex        =   32
      Top             =   2565
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   13
      Left            =   2565
      TabIndex        =   31
      Top             =   2610
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   12
      Left            =   2565
      TabIndex        =   29
      Top             =   2340
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aspect X1"
      Height          =   285
      Index           =   12
      Left            =   90
      TabIndex        =   28
      Top             =   2295
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   7
      Left            =   2565
      TabIndex        =   26
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Repeats"
      Height          =   285
      Index           =   7
      Left            =   90
      TabIndex        =   25
      Top             =   4725
      Width           =   960
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offset Y2"
      Height          =   285
      Index           =   6
      Left            =   90
      TabIndex        =   23
      Top             =   3510
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   6
      Left            =   2565
      TabIndex        =   22
      Top             =   3555
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offset X2"
      Height          =   285
      Index           =   5
      Left            =   90
      TabIndex        =   21
      Top             =   3240
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   5
      Left            =   2565
      TabIndex        =   20
      Top             =   3285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   4
      Left            =   2565
      TabIndex        =   17
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offset Y1"
      Height          =   285
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   2025
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   3
      Left            =   2565
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offset X1"
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   1755
      Width           =   960
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Step"
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   11
      Top             =   4455
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   2
      Left            =   2565
      TabIndex        =   10
      Top             =   4500
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   1
      Left            =   2565
      TabIndex        =   8
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amplitude2"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   2970
      Width           =   960
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amplitude1"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   1485
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      Height          =   240
      Index           =   0
      Left            =   2565
      TabIndex        =   3
      Top             =   1530
      Width           =   375
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "File"
      Begin VB.Menu mnuFile 
         Caption         =   "Load shape"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save shape"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save as bitmap"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "Picture"
      Begin VB.Menu mnuFull 
         Caption         =   "Full screen"
      End
   End
End
Attribute VB_Name = "CDFig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Px1!(720), Px2!(720), Py1!(720), Py2!(720), Temp$
Dim xx%, yy%, Ptt%, Stp%(23)
Dim Vr!, Vg!, Vb!, Wr!, Wg!, Wb! 'colors
Dim Sr!, Sg!, Sb!
Dim AmX1%, AmY1% 'amplitude
Dim Ox1%, Ox2%, Oy1%, Oy2% 'offset x & y
Dim Rep% 'repeats
Dim Ax1%, Ay1%, Ax2%, Ay2% 'add X & Y
Dim Cx1!, Cx2!, Cy1!, Cy2! 'circle 1 & 2
Dim Aamp1%, Aamp2%
Dim Bg%, Ed%
Private Const Pi = 3.1415927
Private Declare Function SetWindowPos Lib "User32" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer

Private Sub AlwaysOnTop(frmID As Form, OnTop As Integer)
' Pass any non-zero value to Place on top
' Pass zero to remove top-mostness

    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If OnTop Then
        OnTop = SetWindowPos(frmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(frmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Private Sub Command1_Click() 'show
Screen.MousePointer = 11
SetFigure
Screen.MousePointer = 1
'figure.SetFocus
End Sub

Private Sub Command2_Click() 'swap colors
Sr = Label5(0).BackColor
Label5(0).BackColor = Label5(1).BackColor
Label5(1).BackColor = Sr
Command1_Click
End Sub

Private Sub Form_Load()
Dim t%
On Error Resume Next
Figure.Move 0, 0, 12000, 9000
Figure.BackColor = 0
CDFig.Move 0, 0
Check1.Enabled = False
For xx = 1 To 360
If 360 / xx = Int(360 / xx) Then
Stp(t) = 360 / xx
t = t + 1
End If
Next xx
For xx = 0 To 719
Px1(xx) = Sin(xx / 180 * Pi) 'radials to degrees
Px2(xx) = Sin(xx / 180 * Pi) 'radials to degrees
Py1(xx) = Cos(xx / 180 * Pi) 'radials to degrees
Py2(xx) = Cos(xx / 180 * Pi) 'radials to degrees
Next xx
For xx = 0 To 11
Label3(xx).Caption = Format(HScroll2(xx).Value, "000")
Next xx
For xx = 12 To 15
Label3(xx).Caption = Format(HScroll2(xx).Value / 10, "0.0")
Next xx
Label5(0).BackColor = RGB(255, 0, 0)
Label5(1).BackColor = RGB(0, 0, 255)
AlwaysOnTop CDFig, 1
Figure.Show
Label2.Caption = "No Title"
Command1_Click
CDFig.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Temp = MsgBox("Quit the complex shapes?", vbQuestion + vbYesNo, "Complex Shapes")
If Temp = vbNo Then Exit Sub
AlwaysOnTop CDFig, 0
End
End Sub


Private Sub HScroll2_Change(Index As Integer)
If Index < 12 Then
Label3(Index).Caption = Format(HScroll2(Index).Value, "000")
End If
If Index = 12 Or Index = 13 Or Index = 14 Or Index = 15 Then
Label3(Index).Caption = Format(HScroll2(Index).Value / 10, "0.0")
End If
If Index = 16 Or Index = 17 Then
Label3(Index).Caption = Format(HScroll2(Index).Value, "000")
End If

If HScroll2(7).Value = 0 Then
Check1.Enabled = False
Else
Check1.Enabled = True
End If
End Sub

Private Sub SetFigure()
On Error Resume Next
AmX1 = HScroll2(0).Value 'amplitude 1
AmY1 = HScroll2(1).Value 'amplitude 2
Ox1 = HScroll2(3).Value 'offset x1
Oy1 = HScroll2(4).Value 'offset y1
Ox2 = HScroll2(5).Value 'offset x2
Oy2 = HScroll2(6).Value 'offset y2
Rep = HScroll2(7).Value 'repeats
Ax1 = HScroll2(8).Value 'add x1
Ay1 = HScroll2(9).Value 'add y1
Ax2 = HScroll2(10).Value 'add x2
Ay2 = HScroll2(11).Value 'add y2
Cx1 = HScroll2(12).Value / 10
Cy1 = HScroll2(13).Value / 10
Cx2 = HScroll2(14).Value / 10
Cy2 = HScroll2(15).Value / 10
Aamp1 = HScroll2(16).Value 'add ampl. 1
Aamp2 = HScroll2(17).Value ' add ampl. 2
Ptt = Stp(HScroll2(2).Value)
MakeColors Label5(0).BackColor, Label5(1).BackColor, Ptt, Option1.Value
Figure.Cls
For yy = 0 To Rep
For xx = 0 To 360 Step Ptt
Figure.Line ((Px1(xx) * AmX1 * Cx1) + Ox1, (Py1(xx) * (AmX1 * Cy1)) + Oy1)-((Px2(xx - (Ptt / 2)) * AmY1 * Cx2) + Ox2, (Py2(xx - (Ptt / 2)) * AmY1 * Cy2) + Oy2), RGB(Vr, Vg, Vb)
If xx <> 0 Then
Figure.Line -((Px2(xx + (Ptt / 2)) * AmY1 * Cx2) + Ox2, (Py2(xx + (Ptt / 2)) * AmY1 * Cy2) + Oy2), RGB(Vr, Vg, Vb)
End If
Figure.Line ((Px1(xx) * AmX1 * Cx1) + Ox1, (Py1(xx) * AmX1 * Cy1) + Oy1)-((Px2(xx + (Ptt / 2)) * AmY1 * Cx2) + Ox2, (Py2(xx + (Ptt / 2)) * AmY1 * Cy2) + Oy2), RGB(Vr, Vg, Vb)
                If xx < 180 Then
                    Vr = Vr + Sr
                    Vg = Vg + Sg
                    Vb = Vb + Sb
                Else
                    Vr = Vr - Sr
                    Vg = Vg - Sg
                    Vb = Vb - Sb
                End If
                    If Vr < 0 Then Vr = 0
                    If Vr > 255 Then Vr = 255
                    If Vg < 0 Then Vg = 0
                    If Vg > 255 Then Vg = 255
                    If Vb < 0 Then Vb = 0
                    If Vb > 255 Then Vb = 255
Next xx
Ox1 = Ox1 + Ax1
Oy1 = Oy1 + Ay1
Ox2 = Ox2 + Ax2
Oy2 = Oy2 + Ay2
AmX1 = AmX1 + Aamp1
AmY1 = AmY1 + Aamp2
MakeColors Label5(0).BackColor, Label5(1).BackColor, Ptt, Option1.Value
Next yy
If Check1.Value = 0 Then Exit Sub
'if symetric
Ox1 = HScroll2(3).Value 'offset x1
Oy1 = HScroll2(4).Value 'offset y1
Ox2 = HScroll2(5).Value 'offset x2
Oy2 = HScroll2(6).Value 'offset y2
AmX1 = HScroll2(0).Value 'amplitude 1
AmY1 = HScroll2(1).Value 'amplitude 2
MakeColors Label5(0).BackColor, Label5(1).BackColor, Ptt, Option1.Value
For yy = 0 To Rep
For xx = 0 To 360 Step Ptt
Figure.Line ((Px1(xx) * AmX1 * Cx1) + Ox1, (Py1(xx) * (AmX1 * Cy1)) + Oy1)-((Px2(xx - (Ptt / 2)) * AmY1 * Cx2) + Ox2, (Py2(xx - (Ptt / 2)) * AmY1 * Cy2) + Oy2), RGB(Vr, Vg, Vb)
If xx <> 0 Then
Figure.Line -((Px2(xx + (Ptt / 2)) * AmY1 * Cx2) + Ox2, (Py2(xx + (Ptt / 2)) * AmY1 * Cy2) + Oy2), RGB(Vr, Vg, Vb)
End If
Figure.Line ((Px1(xx) * AmX1 * Cx1) + Ox1, (Py1(xx) * AmX1 * Cy1) + Oy1)-((Px2(xx + (Ptt / 2)) * AmY1 * Cx2) + Ox2, (Py2(xx + (Ptt / 2)) * AmY1 * Cy2) + Oy2), RGB(Vr, Vg, Vb)
                If xx < 180 Then
                    Vr = Vr + Sr
                    Vg = Vg + Sg
                    Vb = Vb + Sb
                Else
                    Vr = Vr - Sr
                    Vg = Vg - Sg
                    Vb = Vb - Sb
                End If
                    If Vr < 0 Then Vr = 0
                    If Vr > 255 Then Vr = 255
                    If Vg < 0 Then Vg = 0
                    If Vg > 255 Then Vg = 255
                    If Vb < 0 Then Vb = 0
                    If Vb > 255 Then Vb = 255
Next xx
Ox1 = Ox1 - Ax1
Oy1 = Oy1 - Ay1
Ox2 = Ox2 - Ax2
Oy2 = Oy2 - Ay2
AmX1 = AmX1 + Aamp1
AmY1 = AmY1 + Aamp2
MakeColors Label5(0).BackColor, Label5(1).BackColor, Ptt, Option1.Value
Next yy
End Sub

Private Sub Label5_Click(Index As Integer)
CD1.FLAGS = 3
CD1.Color = Label5(Index).BackColor
CD1.ShowColor
Label5(Index).BackColor = CD1.Color
Command1_Click
End Sub

Private Sub mnuFile_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0 'load figure
CD1.FileName = ""
CD1.DialogTitle = "Load shape"
CD1.FLAGS = &H400
CD1.DefaultExt = ".fig"
CD1.Filter = "Shapes |*.fig"
CD1.InitDir = App.Path & "\Shapes"
CD1.ShowOpen
If Err = 32755 Then Exit Sub
LoadFigure
Label2.Caption = CD1.FileTitle
Command1_Click
Case 1 'save figure
CD1.FileName = Label2.Caption
CD1.DialogTitle = "Save as shape"
CD1.FLAGS = &H2
CD1.DefaultExt = ".fig"
CD1.Filter = "Shapes |*.fig"
CD1.InitDir = App.Path & "\Shapes"
CD1.ShowSave
If Err = 32755 Then Exit Sub
SaveFigure
Label2.Caption = CD1.FileTitle
Case 3 'save as bmp
CD1.FileName = "Shape.bmp"
CD1.DialogTitle = "Save as bitmap"
CD1.FLAGS = &H2
CD1.DefaultExt = ".bmp"
CD1.Filter = "Bitmap |*.bmp"
CD1.InitDir = App.Path
CD1.ShowSave
If Err = 32755 Then Exit Sub
SavePicture Figure.Image, CD1.FileName
Case 5 'exit
Form_QueryUnload 0, 0

End Select
End Sub

Private Sub LoadFigure()
ff = FreeFile
Dim kk!
On Error GoTo Ld_error
Check1.Enabled = True
Open CD1.FileName For Input As #ff
Input #ff, kk
Label5(0).BackColor = kk
Input #ff, kk
Label5(1).BackColor = kk
Input #ff, kk
If kk = 0 Then
Option1.Value = False
Option2.Value = True
Else
Option1.Value = True
Option2.Value = False
End If
For xx = 0 To 17
Input #ff, kk
HScroll2(xx).Value = kk
Next xx
Input #ff, kk
Check1.Value = kk
Close #ff
If HScroll2(7).Value = 0 Then
Check1.Enabled = False
Else
Check1.Enabled = True
End If
Exit Sub
Ld_error:
MsgBox "Can't load the file...", vbCritical, cdtitle
Close #ff
End Sub

Private Sub SaveFigure()
ff = FreeFile
On Error GoTo Sv_error
Open CD1.FileName For Output As #ff
Print #ff, Label5(0).BackColor
Print #ff, Label5(1).BackColor
    If Option1.Value = True Then
    Print #ff, 1
    Else
    Print #ff, 0
    End If
For xx = 0 To 17
Print #ff, HScroll2(xx).Value
Next xx
Print #ff, Check1.Value
Close #ff
Exit Sub
Sv_error:
MsgBox "Can't save the file...", vbCritical, cdtitle
Close #ff
End Sub

Private Sub MakeColors(Col1!, Col2!, St%, Grad As Boolean)
Dim St2%
Sr = 0
Sg = 0
Sb = 0
Vr = Col1 Mod 256&
Vg = ((Col1 And &HFF00) / 256&) Mod 256&
Vb = (Col1 And &HFF0000) / 65536
Wr = Col2 Mod 256&
If Grad = False Then
Wg = ((Col2 And &HFF00) / 256&) Mod 256&
Wb = (Col2 And &HFF0000) / 65536
St2 = 360 / St
Sr = (Wr - Vr) / (St2 / 2)
Sg = (Wg - Vg) / (St2 / 2)
Sb = (Wb - Vb) / (St2 / 2)
End If
End Sub

Private Sub mnuFull_Click()
CDFig.Hide
End Sub

Private Sub Option1_Click()
Command1_Click
End Sub

Private Sub Option2_Click()
Command1_Click
End Sub
