VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CylOx"
   ClientHeight    =   1320
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3945
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Main.frx":030A
   ScaleHeight     =   1320
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton GasType 
      Caption         =   "He/O2"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   11
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton GasType 
      Caption         =   "O2/CO2"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton GasType 
      Caption         =   "O2,O2/N2,air"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton timeUnit 
      Caption         =   "minutes"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3600
      Top             =   840
   End
   Begin VB.TextBox Result 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Rate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox Cyl 
      Height          =   315
      ItemData        =   "Main.frx":0875
      Left            =   1920
      List            =   "Main.frx":0894
      TabIndex        =   0
      Text            =   "Type"
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label dprt 
      Caption         =   "debug"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Flow Rate (Liters/min)"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "PSI"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'calculate duration of oxygen cylinders
' 1.22, .796

Option Explicit
Dim factor!(8), timescale!, gtype!

Private Sub Form_Load()
factor!(0) = 0.04   'AA
factor!(1) = 0.068  'BB/M6
factor!(2) = 0.11   'C
factor!(3) = 0.16   'D
factor!(4) = 0.28   'E
factor!(5) = 2.41   'G
factor!(6) = 3.14   'H/K
factor!(7) = 344    'LOX lbs
factor!(8) = 756.8  'LOX kgs
timescale! = 1
gtype! = 1
End Sub

Private Sub Command1_Click()
    Result.Text = Format(Text1.Text * gtype! * factor!(Cyl.ListIndex) / Rate.Text / timescale!, "###.000")
End Sub

Private Sub GasType_Click(Index As Integer)
    Select Case Index
        Case 0
            gtype! = 1
        Case 1
            gtype! = 1.22
        Case 2
            gtype! = 0.796
    End Select
    Command1_Click
End Sub

Private Sub mAbout_Click()
    Main.Enabled = False
    About.Show
End Sub

Private Sub Timer1_Timer()
Dim calcEnb%

'    dprt.Caption = Cyl.ListIndex
    If Cyl.ListIndex = 7 Then
        Label1.Caption = "pounds"
    Else
        Label1.Caption = "PSI"
    End If
    If Cyl.ListIndex = 8 Then Label1.Caption = "kilos"
    calcEnb% = 1
    If Cyl.ListIndex < 0 Then calcEnb% = 0
    If IsNumeric(Text1.Text) = False Then calcEnb% = 0
    If IsNumeric(Rate.Text) = False Then calcEnb% = 0
    If Val(Rate.Text) < 0.1 Then calcEnb% = 0
    If calcEnb% Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub

Private Sub timeUnit_Click()
    If timescale! = 1 Then
        timeUnit.Caption = "hours": timescale! = 60
    Else
        timeUnit.Caption = "minutes": timescale! = 1
    End If
    If Command1.Enabled = True Then Command1_Click
End Sub
