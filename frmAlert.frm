VERSION 5.00
Begin VB.Form frmAlertApp 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CyberSentry Personal Firewall"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   1680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Allow"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Block"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trust"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAlert.frx":000C
      ForeColor       =   &H00C0C0FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CyberSentry Alert!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAlertApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting App.Title, "Rules", Tag, "Trust"

For i = 1 To frmMain.lvMonitor.ListItems.Count
If frmMain.lvMonitor.ListItems.item(i).Text = Tag Then
frmMain.lvMonitor.ListItems.item(i).Selected = True
'frmMain.mnuSecurityAllow_Click
Exit For
End If
Next i
mdlProcess.ResumeThreads CLng(lblDesc.Tag)
CloseAlert = Timer1.Tag
Hide
End Sub

Private Sub Command2_Click()
SaveSetting App.Title, "Rules", Tag, "Terminate"

For i = 1 To frmMain.lvMonitor.ListItems.Count
If frmMain.lvMonitor.ListItems.item(i).Text = Tag Then
frmMain.lvMonitor.ListItems.item(i).Selected = True
'frmMain.mnuSecurityBlock_Click
SaveSetting App.Title, "Rules", Tag, "Terminate"
Exit For
End If
Next i

KillProcessById lblDesc.Tag
mdlProcess.ResumeThreads CLng(lblDesc.Tag)
CloseAlert = Timer1.Tag
Hide
End Sub

Private Sub Command3_Click()
SaveSetting App.Title, "Rules", Tag, "Block"

For i = 1 To frmMain.lvMonitor.ListItems.Count
If frmMain.lvMonitor.ListItems.item(i).Text = Tag Then
frmMain.lvMonitor.ListItems.item(i).Selected = True
'frmMain.mnuSecurityBlock_Click
Exit For
End If
Next i
mdlProcess.ResumeThreads CLng(lblDesc.Tag)
CloseAlert = Timer1.Tag
Hide
End Sub

Private Sub Command4_Click()
SaveSetting App.Title, "Rules", Tag, "Ask"

For i = 1 To frmMain.lvMonitor.ListItems.Count
If frmMain.lvMonitor.ListItems.item(i).Text = Tag Then
frmMain.lvMonitor.ListItems.item(i).Selected = True
'frmMain.mnuSecurityAsk
Exit For
End If
Next i
mdlProcess.ResumeThreads CLng(lblDesc.Tag)
CloseAlert = Timer1.Tag
Hide
End Sub


Private Sub Timer1_Timer()
If CloseAlert = Timer1.Tag Then Command4_Click: Timer1.Enabled = False
End Sub


