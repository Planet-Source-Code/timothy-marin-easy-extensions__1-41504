VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Extensions"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Menu Item's"
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4815
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtMPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "Play File"
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdAddMenu 
         Caption         =   "Create Menu Item"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtCall 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "/Play"
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtMExt 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "App :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Call :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   975
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Text :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   615
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create Extension"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CDB 
         Left            =   3480
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdAddExt 
         Caption         =   "Create Extension"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtIcon 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Icon :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   615
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "App :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   255
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------.
'(C)2002 IntraDream.com|
'Code By Timothy Marin |
'----------------------`

Private Sub cmdAddExt_Click()
    AddExt txtExt.Text, txtIcon.Text, txtPath.Text
End Sub

Private Sub cmdAddMenu_Click()
    AddMenu txtMExt.Text, txtMPath.Text, txtText.Text, txtCall.Text
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    If Index = 0 Or Index = 1 Then
        CDB.Filter = "Your App(.exe)|*.exe"
    Else
        CDB.Filter = "Your Icon(.ico)|*.ico"
    End If
    CDB.ShowOpen
    If Index = 0 Then
        txtPath.Text = CDB.FileName
    ElseIf Index = 1 Then
        txtMPath.Text = CDB.FileName
    Else
        txtIcon.Text = CDB.FileName
    End If
End Sub

Private Sub Form_Load()
    MsgBox "To Retreave info on load use VarInfo = Command$..." & vbCrLf & "Extensions icons will not be shown until Reboot..."
End Sub
