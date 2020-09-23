VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Binary clock. The time is:"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   60
      Top             =   60
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   1260
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   930
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   2
      Left            =   570
      TabIndex        =   2
      Top             =   1260
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   3
      Left            =   570
      TabIndex        =   3
      Top             =   930
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   4
      Left            =   570
      TabIndex        =   4
      Top             =   600
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   5
      Left            =   570
      TabIndex        =   5
      Top             =   270
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   6
      Left            =   1140
      TabIndex        =   6
      Top             =   1260
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   7
      Left            =   1140
      TabIndex        =   7
      Top             =   930
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   8
      Left            =   1140
      TabIndex        =   8
      Top             =   600
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   9
      Left            =   1560
      TabIndex        =   9
      Top             =   1260
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   10
      Left            =   1560
      TabIndex        =   10
      Top             =   930
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   11
      Left            =   1560
      TabIndex        =   11
      Top             =   600
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   12
      Left            =   1560
      TabIndex        =   12
      Top             =   270
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   13
      Left            =   2100
      TabIndex        =   13
      Top             =   1260
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   14
      Left            =   2100
      TabIndex        =   14
      Top             =   930
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   15
      Left            =   2100
      TabIndex        =   15
      Top             =   600
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   16
      Left            =   2520
      TabIndex        =   16
      Top             =   1260
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   17
      Left            =   2520
      TabIndex        =   17
      Top             =   930
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   18
      Left            =   2520
      TabIndex        =   18
      Top             =   600
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin BinaryClock.LED LH 
      Height          =   315
      Index           =   19
      Left            =   2520
      TabIndex        =   19
      Top             =   270
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   1
      Left            =   1890
      TabIndex        =   21
      Top             =   630
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   0
      Left            =   900
      TabIndex        =   20
      Top             =   660
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code by fclage
' email: fclage@gmail.com
' Just a binary clock showing the current time...
' Funny to show to others who don't understand binary. Just tell them the time shown.
' They will ask you: How the hell do you know? :)
' I got the idea from a real desk binary clock available at geeks.com online store.
' Project time: 10mins

Private Sub Timer1_Timer()
RedrawClock
Label1(0).Visible = Not Label1(0).Visible
Label1(1).Visible = Label1(0).Visible
End Sub

Private Sub RedrawClock()
Dim ctime As String
ctime = Format(Now, "hhmmss")


LH(0).Value = CBIT(Mid$(ctime, 1, 1), 0) ' 1st Digit H
LH(1).Value = CBIT(Mid$(ctime, 1, 1), 1)

LH(2).Value = CBIT(Mid$(ctime, 2, 1), 0) ' 2nd Digit H
LH(3).Value = CBIT(Mid$(ctime, 2, 1), 1)
LH(4).Value = CBIT(Mid$(ctime, 2, 1), 2)
LH(5).Value = CBIT(Mid$(ctime, 2, 1), 3)

LH(6).Value = CBIT(Mid$(ctime, 3, 1), 0) ' 1st Digit M
LH(7).Value = CBIT(Mid$(ctime, 3, 1), 1)
LH(8).Value = CBIT(Mid$(ctime, 3, 1), 2)

LH(9).Value = CBIT(Mid$(ctime, 4, 1), 0) ' 2nd Digit M
LH(10).Value = CBIT(Mid$(ctime, 4, 1), 1)
LH(11).Value = CBIT(Mid$(ctime, 4, 1), 2)
LH(12).Value = CBIT(Mid$(ctime, 4, 1), 3)

LH(13).Value = CBIT(Mid$(ctime, 5, 1), 0) ' 1st Digit S
LH(14).Value = CBIT(Mid$(ctime, 5, 1), 1)
LH(15).Value = CBIT(Mid$(ctime, 5, 1), 2)

LH(16).Value = CBIT(Mid$(ctime, 6, 1), 0) ' 2nd Digit S
LH(17).Value = CBIT(Mid$(ctime, 6, 1), 1)
LH(18).Value = CBIT(Mid$(ctime, 6, 1), 2)
LH(19).Value = CBIT(Mid$(ctime, 6, 1), 3)
End Sub

Private Function CBIT(v As String, n As Long) As Boolean
CBIT = (Val(v) And 2 ^ n)
End Function
