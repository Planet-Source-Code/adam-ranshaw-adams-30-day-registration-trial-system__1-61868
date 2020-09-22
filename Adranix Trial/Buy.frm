VERSION 5.00
Begin VB.Form frmBuy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to order WinZip"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Buy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "Web Site"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      Caption         =   "Your Computer ID :"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Please note down your computer ID, it will be asked during registration process."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "If you click ""Web site"", it will launch your web browser to directly connect to our website."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "You can view the order form on X web site or view help file. You can place a credit-card order directly from web site."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
'It will open help file
'Change help.chm to your help file
ShellExecute Me.hwnd, "open", "help.chm", 0, 0, 1
Unload Me
End Sub

Private Sub cmdWebsite_Click()
'It will connect to your buy page
'Change http://www.google.com to your buy page
ShellExecute Me.hwnd, "open", "http://www.adranix.co.uk", 0, 0, 1
Unload Me
End Sub

Private Sub Form_Load()
lblID.Caption = "Your Computer ID : " & modRegCode.getComputerID
End Sub
