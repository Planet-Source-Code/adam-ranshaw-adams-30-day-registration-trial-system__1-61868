VERSION 5.00
Begin VB.Form frmRegCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Registration Code"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "RegCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtRegCode 
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtComputerID 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "&Registration Code :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Computer ID :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   $"RegCode.frx":000C
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmRegCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intEntered As Integer 'Number of times code entered

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'Increase counter
intEntered = intEntered + 1
'If trying to enter invalid code more than 3 times close down
If intEntered > 3 Then End

'Trim fields
txtComputerID.Text = Trim(txtComputerID.Text)
txtRegCode.Text = Trim(txtRegCode.Text)

'Check if username is blank/length>3
If txtComputerID.Text = "" Or Len(txtComputerID.Text) < 3 Then
    MsgBox "Please enter a valid username.", vbOKOnly, "Invalid Information"
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtComputerID.Text)
    txtUserName.SetFocus
    Exit Sub
End If

'Check if regcode is blank/length<8
If txtRegCode.Text = "" Or Len(txtRegCode.Text) < 14 Then
    MsgBox "Please enter your registration code.", vbOKOnly, "Invalid Information"
    txtRegCode.SelStart = 0
    txtRegCode.SelLength = Len(txtRegCode.Text)
    txtRegCode.SetFocus
    Exit Sub
End If

'Check RegCode
If txtRegCode.Text <> modRegCode.GenCode(txtComputerID.Text) Then
    MsgBox "Invalid registration information.", vbOKOnly, "Invalid Information"
    txtRegCode.SelStart = 0
    txtRegCode.SelLength = Len(txtRegCode.Text)
    txtRegCode.SetFocus
    Exit Sub
End If

'Proceed if correct -  make entries in registry
modRegCode.MakeRegEntries txtRegCode.Text

'Tell to restart
MsgBox "Thank you for registering with us." & vbCrLf & _
"Please re-start for registration to take effect.", vbInformation, "Information"
End
End Sub

Private Sub Form_Load()
txtComputerID.Text = modRegCode.getComputerID
End Sub
