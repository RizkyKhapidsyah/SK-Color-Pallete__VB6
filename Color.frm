VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "216 Color Pallet"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "About Me"
      Height          =   345
      Left            =   1410
      MouseIcon       =   "Color.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3420
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   3210
      ScaleHeight     =   825
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Lines"
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   3000
      Width           =   825
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   13
      Cols            =   19
      GridColor       =   -2147483641
      TextStyleFixed  =   3
      FocusRect       =   2
      ScrollBars      =   0
      AllowUserResizing=   2
      PictureType     =   1
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Color.frx":0152
      _NumberOfBands  =   1
      _Band(0).Cols   =   19
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   3
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub grdSetup()

With Me.MSHFlexGrid1
 .ColWidth(0) = .RowHeight(0)
 .RowHeight(1) = .RowHeight(1)
 .Row = 0
.RowHeight(0) = 0
.ColWidth(0) = 0
For i = 1 To .Cols - 1
     .ColWidth(i) = .RowHeight(1)
Next

End With
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = vbChecked Then
Me.MSHFlexGrid1.GridLines = flexGridFlat
Else
Me.MSHFlexGrid1.GridLines = flexGridNone

End If

End Sub

Private Sub CreatePallet()
Dim i As Long
Dim j As Long

Dim r As Long
Dim g As Long
Dim b As Long

Dim x As Long
Dim y As Long
Dim one As Long
Dim two As Long

x = 1
y = 1


one = 1
two = 6

For i = 1 To Me.MSHFlexGrid1.Cols - 1
    For r = 0 To 255 Step 51
       For g = 0 To 255 Step 51
         For b = 0 To 255 Step 51
                If Me.MSHFlexGrid1.Rows = y Then Exit For
                If 0 = (y Mod 2) Then
                    MSHFlexGrid1.Row = two
                Else
                    MSHFlexGrid1.Row = one
                End If
                MSHFlexGrid1.Col = x
             MSHFlexGrid1.CellBackColor = RGB(r, g, b)
                x = x + 1
            If x = 19 Then
                x = 1
                y = y + 1
           If 0 = (y Mod 2) Then
           two = two + 1
           Else
           one = one + 1
           End If
          
                    
            End If
         Next
       Next
    Next
Next

End Sub




Private Sub Command1_Click()
Load Form2
Form2.Show vbModal
End Sub

Private Sub Form_Load()
Call CreatePallet
Me.Check1.Value = vbChecked
Call grdSetup
End Sub



Private Sub MSHFlexGrid1_Click()
Me.Picture1.BackColor = Me.MSHFlexGrid1.CellBackColor
End Sub
