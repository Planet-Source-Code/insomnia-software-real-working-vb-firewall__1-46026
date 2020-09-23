VERSION 5.00
Begin VB.Form frmMinimize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warning"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   Icon            =   "frmMinimize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Do not show this again"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   2520
      Width           =   3675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "CyberSentry Personal Firewall has not been closed. Instead it has been moved to the system tray."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMinimize.frx":57E2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1305
      Picture         =   "frmMinimize.frx":5893
      Top             =   1095
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   840
      Picture         =   "frmMinimize.frx":59E5
      Top             =   720
      Width           =   2280
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "mnuSysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSysTraySettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuSysTraySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysTrayClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMinimize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Check1.Value = vbChecked Then
SaveSetting App.Title, "Settings", "WarnOnExit", "False"
End If

Hide
End Sub


Private Sub mnuSysTrayClose_Click()

If GetSetting(App.Title, "Settings", "WarnOnEnd", "True") = "True" Then
frmExit.Show
Else
frmMain.SysTray.InTray = False
End
End If

End Sub

Private Sub mnuSysTrayOpen_Click()
frmMain.Show
'frmMain.SSTab1.Tab = 0
End Sub

Private Sub mnuSysTraySettings_Click()
frmMain.Show
'frmMain.SSTab1.Tab = 2
End Sub


