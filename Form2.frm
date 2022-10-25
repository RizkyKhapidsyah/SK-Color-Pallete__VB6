VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Me"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   405
      Left            =   2400
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4110
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3555
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form2.frx":0152
      Top             =   300
      Width           =   5985
   End
   Begin VB.Label Label3 
      Caption         =   "polestar@rediffmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4350
      MouseIcon       =   "Form2.frx":0346
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3930
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "e-Mail:"
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   3930
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "This Was Developed by Mr H.R. Renuka Prasad on 7th April 2002"
      Height          =   225
      Left            =   420
      TabIndex        =   0
      Top             =   30
      Width           =   4875
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Label3_Click()
ShellExecute 0, "open", "mailto:" & Me.Label3.Caption, vbNullString, vbNullString, 1
End Sub
