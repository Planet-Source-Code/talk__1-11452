VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmtalk 
   Caption         =   "Talk"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Talk"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Simply Type something in the text box and hit enter."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a Basic Program using MSAGENT control 2.0"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3705
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   240
      Top             =   2640
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "frmtalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DATAPATH = "C:\Windows\Msagent\chars\peedy.acs"
Dim Peedy As IAgentCtlCharacter

Private Sub Command1_Click()
Peedy.Speak Text1.Text
End Sub

Private Sub Form_Load()
Agent1.Characters.Load "Peedy", DATAPATH
    Set Peedy = Agent1.Characters("Peedy")
    Peedy.Show
    Peedy.MoveTo 500, 285
    Peedy.Speak "Let's get this show on the road!"
End Sub
