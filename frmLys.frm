VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Solution 1.0A"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   2550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2055
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3625
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Push the [Start] button and wait..."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t(26) As Integer
Dim x(6, 6) As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim pro As Integer

Dim fa As Double
Dim fb As Double
Dim n As Integer
Dim box


Private Sub Command1_Click()
n = 5
pro = 1
fa = 0
fb = 33554432 / 100 * pro

For b = 1 To 25 Step 1
t(b) = 0
Next b
'---------------------------------
om:
If t(1) = 1 Then
    a = 1
    Do While t(a) = 1
        a = a + 1
    Loop

    If a = 26 Then
            MsgBox (" No solution")
            GoTo skidt
    End If

    t(a) = 1
    
    
        For b = 1 To a - 1 Step 1
        t(b) = 0
        Next b
Else
    t(1) = 1
End If
For a = 1 To 5 Step 1
    For b = 1 To 5 Step 1
        x(a, b) = -1    '-1 = closed
    Next b
Next a


'--------------------------------------
'coordinates
'--------------------------------------
For a = 0 To 24 Step 1
    
 If t(a + 1) = 1 Then
        'middle
        x1 = Int(a / 5)
        x1 = x1 * 5
        x1 = a - x1
        y1 = a - x1
        y1 = Int(y1 / 5)
        y1 = y1 + 1
        x1 = x1 + 1
        x(x1, y1) = x(x1, y1) * (-1)
        
        'left
        x2 = x1 - 1
        y2 = y1
        If x2 > 0 Then
                 x(x2, y2) = x(x2, y2) * (-1)
        End If
        
        'right
        x2 = x1 + 1
        y2 = y1
        If x2 < 6 Then
            x(x2, y2) = x(x2, y2) * (-1)
        End If
        
        'upper
        x2 = x1
        y2 = y1 + 1
        If y2 < 6 Then
            x(x2, y2) = x(x2, y2) * (-1)
        End If
        
        'down
        x2 = x1
        y2 = y1 - 1
        If y2 > 0 Then
            x(x2, y2) = x(x2, y2) * (-1)
        End If
 End If
Next a



c = 0
For a = 1 To 5 Step 1
    For b = 1 To 5 Step 1
        c = c + x(a, b)
    Next b
Next a
fa = fa + 1


If fa >= fb Then

'the progressbar
    pro = pro + 1
    ProgressBar1.Value = pro
    fb = 33554432 / 100 * pro

End If


If c <> 25 Then
    GoTo om
End If

Text1.Text = ""

'print the solution
'X= Light-IT
'0= Don't Light - IT

n = 5
For a = 0 To 24 Step 1
    
    
    If t(a + 1) = 1 Then
        Text1.Text = Text1.Text & "X"
        If Len(Text1.Text) = n Then
        Text1.Text = Text1.Text & vbNewLine
        n = n + 7
        End If
    Else
        Text1.Text = Text1.Text & "0"
        If Len(Text1.Text) = n Then
        Text1.Text = Text1.Text & vbNewLine
        n = n + 7
        End If
    End If
Next a


MsgBox ("Solution Found!!!")

skidt:
End Sub

Private Sub Command2_Click()
Unload Me
frmLightItOn.Show
End Sub

Private Sub Command3_Click()
MsgBox "This program runs through all combinations (2^25= 33554432). Therefore it takes some time (depends on the pc. - at 450mHz about 2 min - ) before the solution is found. Follow the progressbar. It stops when it has found the first solution. Then the text-box will be filled with X and 0.  X means that you should click (once).", vbOKOnly, "Read this"
End Sub
