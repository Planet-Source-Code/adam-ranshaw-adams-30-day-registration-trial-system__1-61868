VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evaluation Version Notice"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "Nag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRegCode 
      Caption         =   "Enter &Registration Code"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "Use Evaluation &Version"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&Buy Now"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Nag.frx":000C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Nag.frx":0126
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label lblTrial 
      Alignment       =   2  'Center
      Caption         =   "30 Days Left"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmNag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------
'TODO
'Change Image1 to your own logo picture
'Add these forms to your main application project
'Set Startup Object to frmNag.frm
'Change all texts in labels to your application name
'---------------------------------------

Private Sub cmdBuy_Click()
'Opens Buy dialog
frmBuy.Show vbModal
End Sub

Private Sub cmdEvaluate_Click()
'-------------------------------------
'TODO
'Put code here to launch you application's main form
'example: frmMain.show
'-------------------------------------
MsgBox "Your Form Here"
Unload Me
End Sub

Private Sub cmdRegCode_Click()
'Opens registraion code entry dialog
frmRegCode.Show vbModal
End Sub


Private Sub Form_Load()
MsgBox "I am currently giving away 20MB of FTP space for free, please goto www.adranix.co.uk or contact aranshaw@aol.com for more infomation.  THANKS!", vbInformation
'If one instance allready running, close down
If App.PrevInstance Then End

'Check Registraion
If modRegCode.IsRegistered = True Then 'Registered
    '-----------------------------------------------
    'TODO
    'Put code to show your app's main form
    'frmMain.Show 'Change it with you app's main form
    '-----------------------------------------------
    MsgBox "Registered"
    End
End If

Dim intDays As Integer
intDays = modRegCode.getTrialDays
PBar.Max = 30
PBar.Value = intDays
lblTrial = "You have " & CStr(intDays) & " Days left on Free Trail.  To keep this software on your Ccmputer and remove this screen please register."
If intDays = 0 Then 'Expired
    cmdEvaluate.Enabled = False 'Cannt evaluate
    lblTrial = "Your free Trial has Expired.  Please register this software to remove this screen and the 30 day limit.  THANK YOU!"
    lblTrial.ForeColor = vbRed
End If

'Randomize command button positions
'This stops brute force attacking
Call RandomButtons
End Sub

Private Function RandomButtons()
'This function randomizes positions of buttons
Dim j As Integer, pos(6) As Long, m As Integer
Dim n As Integer

pos(1) = cmdBuy.Left
pos(2) = cmdEvaluate.Left
pos(3) = cmdRegCode.Left

Randomize
j = Int((Rnd * 3) + 1)
cmdBuy.Left = pos(j)

n = 4
For m = 1 To 3
    If m <> j Then
        pos(n) = pos(m)
        n = n + 1
    End If
Next

Do While j < 4
    DoEvents
    Randomize
    j = Int((Rnd * 5) + 1)
Loop
cmdEvaluate.Left = pos(j)
If j = 5 Then
    cmdRegCode.Left = pos(4)
Else
    cmdRegCode.Left = pos(5)
End If
End Function


