VERSION 5.00
Begin VB.UserControl LED 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "LED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Code by fclage
' email: fclage@gmail.com
' Just a binary clock showing the current time...
' Funny to show to others who don't understand binary. Just tell them the time shown.
' They will ask you: How the hell do you know? :)
' I got the idea from a real desk binary clock available at geeks.com online store.

Private iForeColor As OLE_COLOR
Private iLowColor As OLE_COLOR
Private iValue As Boolean

Private Sub UserControl_Initialize()
Value = False
iForeColor = RGB(64, 64, 255)
iLowColor = RGB(64, 64, 64)
End Sub

Property Let Value(nv As Boolean)
If nv = iValue Then Exit Property
iValue = nv
RedrawMe
End Property
Property Get Value() As Boolean
Value = iValue
End Property
Property Let ForeColor(nv As OLE_COLOR)
iForeColor = nv
RedrawMe
End Property

Private Sub RedrawMe()
Dim cx As Long, cy As Long, r As Long
cx = UserControl.ScaleWidth / 2
cy = UserControl.ScaleHeight / 2
If cx > cy Then r = cy - 2 Else r = cx - 2

UserControl.AutoRedraw = True
UserControl.Cls
If iValue Then
    UserControl.FillColor = iForeColor ' ACTIVE
    Else
    UserControl.FillColor = iLowColor  ' INACTIVE
    End If
UserControl.FillStyle = 0 ' Solid
UserControl.Circle (cx, cy), r, vbWhite
UserControl.FillColor = RGB(192, 192, 192)
UserControl.Circle (cx / 2 + cx / 3, cy / 2 + cy / 3), r / 5
End Sub

Private Sub UserControl_Resize()
RedrawMe
End Sub
