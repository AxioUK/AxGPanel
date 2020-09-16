VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Visible=True"
      Height          =   360
      Left            =   2100
      TabIndex        =   7
      Top             =   195
      Width           =   1350
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Cross Visible"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   5
      Top             =   690
      Width           =   1860
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Moveable"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   1125
   End
   Begin Proyecto1.AxGPanel AxGPanel1 
      Height          =   2745
      Left            =   2730
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4842
      Color2          =   8454143
      Angulo          =   45
      BorderColor     =   12582912
      BorderWidth     =   2
      CornerCurve     =   1
      CrossPosition   =   2
      Begin Proyecto1.AxGPanel AxGPanel2 
         Height          =   645
         Left            =   2535
         TabIndex        =   4
         Top             =   1755
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1138
         BorderColor     =   0
      End
      Begin Proyecto1.AxGPanel AxGPanel4 
         Height          =   1140
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   2011
         Angulo          =   45
         BorderWidth     =   3
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   300
            Left            =   465
            TabIndex        =   2
            Top             =   540
            Width           =   570
         End
      End
   End
   Begin Proyecto1.AxGPanel AxGPanel3 
      Height          =   1140
      Left            =   480
      TabIndex        =   8
      Top             =   2115
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   2011
      Angulo          =   45
      BorderWidth     =   3
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   300
         Left            =   465
         TabIndex        =   9
         Top             =   540
         Width           =   570
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AxGPanel1_CrossClick()
AxGPanel1.Visible = False
End Sub

Private Sub AxGPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Caption = "X:" & X & " Y:" & Y
End Sub

Private Sub Check1_Click()
AxGPanel1.Moveable = Check1.Value
AxGPanel2.Moveable = Check1.Value
AxGPanel3.Moveable = Check1.Value
AxGPanel4.Moveable = Check1.Value

End Sub

Private Sub Check2_Click()
AxGPanel1.CrossVisible = Check2.Value
AxGPanel3.CrossVisible = Check2.Value
AxGPanel4.CrossVisible = Check2.Value
End Sub

Private Sub Command1_Click()
Dim cObj As Object
For Each cObj In Form1.Controls
  If TypeOf cObj Is AxGPanel Then cObj.Visible = True
Next

End Sub

Private Sub Form_Load()
With List1
  .AddItem "cTopRight"
  .AddItem "cMiddleRight"
  .AddItem "cBottomRight"
  .AddItem "cTopLeft"
  .AddItem "cMiddleLeft"
  .AddItem "cBottomLeft"
  .AddItem "cMiddleTop"
  .AddItem "cMiddleBottom"
End With

End Sub

Private Sub List1_Click()
AxGPanel1.CrossPosition = List1.ListIndex
AxGPanel3.CrossPosition = List1.ListIndex
AxGPanel4.CrossPosition = List1.ListIndex
End Sub
