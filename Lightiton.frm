VERSION 5.00
Begin VB.Form frmLightItOn 
   Caption         =   "LightItOn 1.0A"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   31
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdLøsning 
      Caption         =   "Solution"
      Height          =   375
      Left            =   960
      TabIndex        =   29
      ToolTipText     =   "Need Help?"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      Height          =   1695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   5520
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clean"
      Height          =   375
      Left            =   2400
      TabIndex        =   26
      ToolTipText     =   "Clean All"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Left Click and light on all 25 boxes."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   5160
      TabIndex        =   27
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   24
      Left            =   4010
      TabIndex        =   24
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   23
      Left            =   3010
      TabIndex        =   23
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   22
      Left            =   2010
      TabIndex        =   22
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   21
      Left            =   1010
      TabIndex        =   21
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   20
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   19
      Left            =   4010
      TabIndex        =   19
      Top             =   1010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   18
      Left            =   3010
      TabIndex        =   18
      Top             =   1010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   17
      Left            =   2010
      TabIndex        =   17
      Top             =   1010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   16
      Left            =   1010
      TabIndex        =   16
      Top             =   1010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   15
      Left            =   0
      TabIndex        =   15
      Top             =   1010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   14
      Left            =   4010
      TabIndex        =   14
      Top             =   2010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   13
      Left            =   3010
      TabIndex        =   13
      Top             =   2010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   12
      Left            =   2010
      TabIndex        =   12
      Top             =   2010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   11
      Left            =   1010
      TabIndex        =   11
      Top             =   2010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Top             =   2010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   9
      Left            =   4010
      TabIndex        =   9
      Top             =   3010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   8
      Left            =   3010
      TabIndex        =   8
      Top             =   3010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   7
      Left            =   2010
      TabIndex        =   7
      Top             =   3010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   6
      Left            =   1010
      TabIndex        =   6
      Top             =   3010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   3010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   4
      Left            =   4010
      TabIndex        =   4
      Top             =   4010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   3
      Left            =   3010
      TabIndex        =   3
      Top             =   4010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   2
      Left            =   2010
      TabIndex        =   2
      Top             =   4010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   1
      Left            =   1010
      TabIndex        =   1
      Top             =   4010
      Width           =   1000
   End
   Begin VB.Label lblKor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   4010
      Width           =   1000
   End
End
Attribute VB_Name = "frmLightItOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x(26) As Integer
Dim Y(26) As Integer
Dim Tal As Integer
Dim KorX As Integer
Dim KorY(6)
Dim KorZ As Integer
Dim KorX1 As Integer
Dim KorY1 As Integer
Dim KorZ1 As Integer
Dim p As Integer
Dim hFarve(26) As Integer
Dim a As Long
Dim Tjek As Integer
Dim TjekIt As Boolean
Dim t As Long


Private Sub cmdLøsning_Click()
Me.Hide
Form1.Show


End Sub

Private Sub Command1_Click()
Dim t As Integer
On Error Resume Next
For t = 0 To 7 Step 2
For p = 0 To 7 Step 2
frmLightItOn.lblKor((t * 8) + p).BackColor = &H80000005
frmLightItOn.lblKor((t * 8) + (p + 1)).BackColor = &H80000005
Next p
Next t
For t = 1 To 7 Step 2
For p = 0 To 7 Step 2
frmLightItOn.lblKor((t * 8) + p).BackColor = &H80000005
frmLightItOn.lblKor((t * 8) + (p + 1)).BackColor = &H80000005
Next p
Next t
Text1.Text = ""
a = 0
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
a = 0
Tjek = 0
TjekIt = False


'identify the boxes
For p = 0 To 24 Step 1

If frmLightItOn.lblKor(p).BackColor = &H80000018 Then
    hFarve(p) = 1
End If

If frmLightItOn.lblKor(p).BackColor = &H80000005 Then
    hFarve(p) = 2
End If

Next p



'x
For p = 0 To 20 Step 5
x(p) = 0
Next p
For p = 1 To 21 Step 5
x(p) = 1
Next p
For p = 2 To 22 Step 5
x(p) = 2
Next p
For p = 3 To 23 Step 5
x(p) = 3
Next p
For p = 4 To 24 Step 5
x(p) = 4
Next p

'y
For p = 0 To 4 Step 1
Y(p) = 0
Next p
For p = 5 To 9 Step 1
Y(p) = 1
Next p
For p = 10 To 14 Step 1
Y(p) = 2
Next p
For p = 15 To 19 Step 1
Y(p) = 3
Next p
For p = 20 To 24 Step 1
Y(p) = 4
Next p

End Sub

Private Sub lblKor_Click(Index As Integer)

On Error Resume Next
Label1.Caption = Tjek
Tjek = 0
Tal = Index
KorX = x(Tal) + Y(Tal) * 5
KorY(0) = KorX + 5
KorY(1) = KorX - 5
KorY(2) = KorX - 1
KorY(3) = KorX + 1
KorY(4) = KorX

Label2.Caption = KorX



'next move
a = a + 1


Text1.Text = Text1.Text & a & "|" & "[" & x(Tal) + 1 & "," & Y(Tal) + 1 & "] "

'could have been done with coordinates

If Tal = 5 Or Tal = 10 Or Tal = 15 Or Tal = 20 Then

For p = 0 To 4 Step 1
If p = 2 Then
p = p + 1
End If
If frmLightItOn.lblKor(KorY(p)).BackColor = &H80000005 Then
    frmLightItOn.lblKor(KorY(p)).BackColor = &H80000018

ElseIf frmLightItOn.lblKor(KorY(p)).BackColor = &H80000018 Then
    frmLightItOn.lblKor(KorY(p)).BackColor = &H80000005
End If

Next p

GoTo done
End If

If Tal = 9 Or Tal = 14 Or Tal = 19 Or Tal = 4 Then
For p = 0 To 4 Step 1
If p = 3 Then
p = p + 1
End If
If frmLightItOn.lblKor(KorY(p)).BackColor = &H80000005 Then
    frmLightItOn.lblKor(KorY(p)).BackColor = &H80000018

ElseIf frmLightItOn.lblKor(KorY(p)).BackColor = &H80000018 Then
    frmLightItOn.lblKor(KorY(p)).BackColor = &H80000005
End If
Next p
GoTo done
End If



For p = 0 To 4 Step 1

If frmLightItOn.lblKor(KorY(p)).BackColor = &H80000005 Then
    frmLightItOn.lblKor(KorY(p)).BackColor = &H80000018
    
ElseIf frmLightItOn.lblKor(KorY(p)).BackColor = &H80000018 Then
    frmLightItOn.lblKor(KorY(p)).BackColor = &H80000005
End If
Next p


For p = 0 To 24 Step 1
If frmLightItOn.lblKor(p).BackColor = &H80000005 Then
TjekIt = True
Else
TjekIt = False
End If
If TjekIt = True Then
Tjek = Tjek + 1
End If

Next p
Label1.Caption = 25 - Tjek

If (25 - Tjek) = 25 Then
If a > 16 Then
MsgBox ("You solved'LightItOn', but you made to many 'moves'!. It could be done in fifteen moves but you did it in: " & (a) & "Better Luck next time")
ElseIf a = 15 Then
MsgBox ("Congratulation. Either you are very clever/very lucky/just tried it before. You passed it with the minimum of moves (15)")
End If
    
End If


done:
End Sub
