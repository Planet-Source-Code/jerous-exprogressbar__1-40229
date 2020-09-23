VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "*\AProject1.vbp"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin Project1.exProgressBar prg 
      Height          =   315
      Left            =   15
      TabIndex        =   21
      Top             =   75
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      BackColor       =   0
      Value           =   50
      PercentColor    =   16777215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   495
      Left            =   4020
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4185
      Top             =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test!"
      Height          =   495
      Left            =   2820
      TabIndex        =   19
      Top             =   2625
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1125
      TabIndex        =   18
      Top             =   2940
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   1125
      TabIndex        =   17
      Top             =   2685
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Enabled"
      Height          =   285
      Left            =   30
      TabIndex        =   15
      Top             =   1890
      Value           =   1  'Checked
      Width           =   4050
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ShowPercent"
      Height          =   285
      Left            =   30
      TabIndex        =   13
      Top             =   1650
      Value           =   1  'Checked
      Width           =   4050
   End
   Begin VB.CheckBox Check1 
      Caption         =   "RelativeScroll"
      Height          =   285
      Left            =   30
      TabIndex        =   12
      Top             =   1380
      Width           =   4050
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   390
      Left            =   60
      Max             =   100
      TabIndex        =   11
      Top             =   810
      Value           =   100
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Endcolor"
      Height          =   1290
      Left            =   4725
      TabIndex        =   5
      Top             =   1305
      Width           =   3420
      Begin VB.HScrollBar HScroll4 
         Height          =   210
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   990
         Width           =   2640
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   210
         Left            =   720
         Max             =   255
         TabIndex        =   7
         Top             =   570
         Width           =   2640
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   210
         Left            =   720
         Max             =   255
         TabIndex        =   6
         Top             =   165
         Width           =   2640
      End
      Begin VB.Label Label2 
         Height          =   900
         Left            =   75
         TabIndex        =   9
         Top             =   225
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Startcolor"
      Height          =   1290
      Left            =   4710
      TabIndex        =   0
      Top             =   30
      Width           =   3420
      Begin VB.HScrollBar scrGr 
         Height          =   210
         Left            =   720
         Max             =   255
         TabIndex        =   3
         Top             =   165
         Width           =   2640
      End
      Begin VB.HScrollBar scrBl 
         Height          =   210
         Left            =   720
         Max             =   255
         TabIndex        =   2
         Top             =   570
         Width           =   2640
      End
      Begin VB.HScrollBar scrRd 
         Height          =   210
         Left            =   720
         Max             =   255
         TabIndex        =   1
         Top             =   990
         Width           =   2640
      End
      Begin VB.Label Label1 
         Height          =   900
         Left            =   75
         TabIndex        =   4
         Top             =   225
         Width           =   525
      End
   End
   Begin MSComctlLib.ProgressBar prg2 
      Height          =   390
      Left            =   60
      TabIndex        =   10
      Top             =   450
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "SegmentWidth  SegmentSize"
      Height          =   420
      Left            =   15
      TabIndex        =   16
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Borderstyle"
      Height          =   285
      Left            =   30
      TabIndex        =   14
      Top             =   2265
      Width           =   4050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
prg.RelativeScroll = -Check1
End Sub

Private Sub Check2_Click()
prg.ShowPercent = -Check2
End Sub

Private Sub Check3_Click()
prg.Enabled = -Check3
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
prg.Value = prg.Min
End Sub

Private Sub Command2_Click()
MsgBox "exProgressBar v1.0" + vbCrLf + vbCrLf + "Made by GR Productions", vbInformation
End Sub

Private Sub Form_Load()
HScroll1_Change
End Sub

Private Sub HScroll1_Change()
prg.Value = HScroll1
prg2.Value = HScroll1
End Sub
Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub Label3_Click()
Label3.BorderStyle = Abs(Label3.BorderStyle - 1)
prg.BorderStyle = Label3.BorderStyle
End Sub

Private Sub scrBl_Change()
Label1.BackColor = RGB(scrRd, scrGr, scrBl)
prg.StartColor = Label1.BackColor
End Sub
Private Sub scrBl_Scroll()
scrBl_Change
End Sub
Private Sub scrGr_Change()
scrBl_Change
End Sub
Private Sub scrGr_Scroll()
scrBl_Scroll
End Sub
Private Sub scrRd_Change()
scrBl_Scroll
End Sub
Private Sub scrRd_Scroll()
scrBl_Scroll
End Sub

Private Sub hscroll2_Change()
Label2.BackColor = RGB(HScroll4, HScroll2, HScroll3)
prg.endColor = Label2.BackColor
End Sub
Private Sub hscroll2_Scroll()
hscroll2_Change
End Sub
Private Sub hscroll3_Change()
hscroll2_Change
End Sub
Private Sub hscroll3_Scroll()
hscroll2_Scroll
End Sub
Private Sub hscroll4_Change()
hscroll2_Change
End Sub
Private Sub hscroll4_Scroll()
hscroll2_Scroll
End Sub

Private Sub Text1_Change()
prg.SegmentWidth = Val(Text1)
End Sub

Private Sub Text2_Change()
prg.Segmentsize = Val(Text2)
End Sub

Private Sub Timer1_Timer()
prg.Value = prg.Value + 1
prg.ToolTipText = Str(prg.Value) + "%"
Caption = Str(prg.Percent) + " procent completed ..."
If prg.Value >= prg.Max Then
  Timer1.Enabled = False
End If
End Sub
