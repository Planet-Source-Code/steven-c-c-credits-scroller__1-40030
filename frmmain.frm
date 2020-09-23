VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scrolling Credits"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox creditsbox 
      Height          =   3495
      Left            =   2280
      ScaleHeight     =   3435
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2520
         Top             =   3000
      End
      Begin VB.PictureBox creditsscroll 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   240
         ScaleHeight     =   6375
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   240
         Width           =   2535
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Project1 Team"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   2535
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   2520
            Y1              =   6255
            Y2              =   6255
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   2520
            Y1              =   6240
            Y2              =   6240
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Jeremy Mills"
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Daniel Hodegins"
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Lindsay Spain"
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Nathan Brown"
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Development"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Andrew whatson"
            Height          =   255
            Left            =   480
            TabIndex        =   16
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Amanda Green"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "David Jones"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Paul Taylor"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Project Leaders"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "joe bloggs"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Richard Mansfield"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Matthew Panter"
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Steven Crowdy"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Senior Management"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   2520
            Y1              =   310
            Y2              =   310
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   2520
            Y1              =   300
            Y2              =   300
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Duncan Leggitt"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   5040
            Width           =   1095
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Genna Kidman"
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Laura Stevenson"
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   5280
            Width           =   1215
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Julie Rutter"
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   5760
            Width           =   855
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Special Thanks to"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   4800
            Width           =   1335
         End
      End
   End
   Begin VB.Label lblabout2 
      BackStyle       =   0  'Transparent
      Caption         =   "You can also add code so the mouseover or click of a name can trigger an event. try clicking Steven Crowdy as an example."
      Height          =   975
      Left            =   240
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblabout1 
      BackStyle       =   0  'Transparent
      Caption         =   "This is a scrolling credits box, that will scroll the credits flicker free."
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
creditsscroll.Top = creditsbox.Height + 100
End Sub

Private Sub Label2_Click()
MsgBox "You clicked steven Crowdy", , "Click"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlue
End Sub


Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000012
End Sub

Private Sub Timer1_Timer()
If creditsscroll.Top = -creditsscroll.Height Then
creditsscroll.Top = creditsbox.Height + 100
Else
creditsscroll.Top = creditsscroll.Top - 5
End If
End Sub



