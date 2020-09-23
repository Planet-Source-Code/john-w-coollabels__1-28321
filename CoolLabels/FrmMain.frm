VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5595
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6360
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5640
      Top             =   1080
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   2040
   End
   Begin VB.Frame Fra3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Progressbar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   2880
      TabIndex        =   14
      Top             =   960
      Width           =   3255
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Lbl5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Progressbar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1845
      End
      Begin VB.Label Lbl4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progressbar Example"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label LblProgress2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   15
      End
      Begin VB.Label LblProgress1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Fra2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Command Buttons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   4455
      Begin VB.CommandButton Cmd1 
         Caption         =   "Standard Commandbutton"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblCmd1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Commandbutton1 Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LblCmd2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Commandbutton2 Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblCmd3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Commandbutton3 Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LblShadow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame Fra1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Checkboxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2655
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FF8080&
         Caption         =   "  Standard Checkbox"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label LblChkBox1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   195
      End
      Begin VB.Label LblChkBox2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   195
      End
      Begin VB.Label LblChkBox3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checkbox1 Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checkbox2 Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Lbl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checkbox3 Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cool labels"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   30
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   630
      Left            =   2040
      TabIndex        =   23
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label LblClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      ToolTipText     =   "Close Button"
      Top             =   0
      Width           =   495
   End
   Begin VB.Label LblMini 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      ToolTipText     =   "Minimize Button"
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Lbl6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Just some simple controls made with Labels instead of the standard controls used for such"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   4920
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Everything was done here by me except the code for the progress bar
'I downloaded it off of here and I am sorry to say it didnt include
'documentation as to who made it, and I cant find it any more. So if
'you recognize the code e-mail me and I will gladly add your name here :)
'Any complaints, comments or tips are greatly appreciated, as those are
'the keys to learning good programming, so you can e-mail me at
'orbitzsoft@excite.com
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Option Explicit
'Used for the formmove function
Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean

'For the progressbar example

Public min As String

'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'The followng code between these lines are for the formmove function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

    MoveMe = True
    OldX = X
    OldY = Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MoveMe = True Then
        Me.Left = Me.Left + (X - OldX)
        Me.Top = Me.Top + (Y - OldY)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

    Me.Left = Me.Left + (X - OldX)
    Me.Top = Me.Top + (Y - OldY)
    MoveMe = False

End Sub
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Private Sub LblMini_Click()

'This is to minimize the form
 Me.WindowState = vbMinimized
 
End Sub

Private Sub LblClose_Click()

 Unload Me
 
End Sub


'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Everything between these lines has the same code as below

Private Sub LblChkBox1_Click()

'This checks to see if the label is checked and if not
'when you click it it will check it and vice versa
If LblChkBox1.Caption = "P" Then
 LblChkBox1.Caption = ""
  ElseIf LblChkBox1.Caption = "" Then
 LblChkBox1.Caption = "P"
End If

End Sub

Private Sub LblChkBox1_DblClick()

If LblChkBox1.Caption = "P" Then
 LblChkBox1.Caption = ""
  ElseIf LblChkBox1.Caption = "" Then
 LblChkBox1.Caption = "P"
End If

End Sub

Private Sub LblChkBox2_Click()

If LblChkBox2.Caption = "P" Then
 LblChkBox2.Caption = ""
  ElseIf LblChkBox2.Caption = "" Then
 LblChkBox2.Caption = "P"
End If

End Sub

Private Sub LblChkBox2_DblClick()

If LblChkBox2.Caption = "P" Then
 LblChkBox2.Caption = ""
  ElseIf LblChkBox2.Caption = "" Then
 LblChkBox2.Caption = "P"
End If

End Sub

Private Sub LblChkBox3_Click()

If LblChkBox3.Caption = "P" Then
 LblChkBox3.Caption = ""
  ElseIf LblChkBox3.Caption = "" Then
 LblChkBox3.Caption = "P"
End If

End Sub

Private Sub LblChkBox3_DblClick()

If LblChkBox3.Caption = "P" Then
 LblChkBox3.Caption = ""
  ElseIf LblChkBox3.Caption = "" Then
 LblChkBox3.Caption = "P"
End If

End Sub
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""

Private Sub LblCmd1_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

'On mouseover the forecolor will change

 LblCmd1.ForeColor = &H80FF&

End Sub

Private Sub Fra2_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

'When you move the mouse off of the button the forecolor will
'change back to its original color
 
 LblCmd1.ForeColor = &H8080&

End Sub

Private Sub LblCmd2_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

'When the mouse is pressed the button will move over the shadow
'giving it a 3d pushed look

LblShadow.Visible = False
 LblCmd2.Left = LblShadow.Left
LblCmd2.Top = LblShadow.Top

End Sub

Private Sub LblCmd2_MouseUp(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

'when the mouse is let go the button will return to its original position

LblShadow.Visible = True
 LblCmd2.Left = 240
LblCmd2.Top = 1320

End Sub

Private Sub LblCmd3_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

'This will give the button a pressed look when the mouse is pressed

 LblCmd3.BorderStyle = 1

End Sub

Private Sub LblCmd3_MouseUp(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

'When the mouse is let go the button will return flat
 LblCmd3.BorderStyle = 0

End Sub

Private Sub Timer1_Timer()

'This is to set the colors after timer2 loops
 LblProgress1.BackColor = &HC0C0FF
 LblProgress2.BackColor = &H8080FF
'""""""""""""""""""""""""""""""""""""""""""""
'What this does is set the lblprogress2's width to lblprogress1's
'width during the running of the timer and the resets its width back
'to the original width once the timer is done

min = 1
 Static times As Integer, percent As Integer, amount As Integer
amount = amount + 1

If times = 0 Then times = LblProgress1.Width / (min * 60)
 percent = (100 / (min * 60)) * amount
LblProgress2.Width = times

'""""""""""""""""""""""""""""""""""""""""""""""""""""
'You could insert another and have the percent shown
'LblProgress3.Caption = percent & "%"
'""""""""""""""""""""""""""""""""""""""""""""""""""""

times = times + (LblProgress1.Width / (min * 60))
 If percent = 100 Then
 
'Once 100% is reached timer1 is disabled and
'timer2 is enabled

 Timer1.Enabled = False
Timer2.Enabled = True
times = 0
End If

End Sub

'The following code does the same as above

Private Sub Timer2_Timer()

LblProgress1.BackColor = &HFFFFC0
 LblProgress2.BackColor = &HFFFF00
 
min = 1
 Static times As Integer, percent As Integer, amount As Integer
amount = amount + 1

If times = 0 Then times = LblProgress1.Width / (min * 60)
 percent = (100 / (min * 60)) * amount
LblProgress2.Width = times

'LblProgress3.Caption = percent & "%"

times = times + (LblProgress1.Width / (min * 60))
 If percent = 100 Then

'Once 100% is reached timer1 is disabled and
'timer2 is enabled

Timer1.Enabled = True
 Timer2.Enabled = False
times = 0
End If

End Sub

'The following two timers enable the standard progressbar to function
'looping for our example

Private Sub Timer3_Timer()
 If ProgressBar1.Value < 100 Then
  ProgressBar1.Value = ProgressBar1.Value + 2
   Else
  Timer3.Enabled = False
 Timer4.Enabled = True
End If

End Sub

Private Sub Timer4_Timer()
If ProgressBar1.Value = 100 Then
 ProgressBar1.Value = 0 And ProgressBar1.Value _
 = ProgressBar1.Value + 2
  Else
  Timer3.Enabled = True
 Timer4.Enabled = False
End If

End Sub

