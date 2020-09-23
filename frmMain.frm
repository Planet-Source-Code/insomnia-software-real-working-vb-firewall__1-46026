VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CyberSentry Personal Firewall  [v1.00]"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin CyberSentry.ProgressCntrl pbIncoming 
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
   End
   Begin CyberSentry.ProgressCntrl pbBlocked 
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
   End
   Begin CyberSentry.cSysTray SysTray 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMain.frx":57E2
      TrayTip         =   "CyberSentry Personal Firewall"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Message Console"
      Height          =   425
      Left            =   2040
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Only show programs using the internet"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   9
      Top             =   6000
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hide windows services"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Option1 
      Caption         =   "Disable Firewall"
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   7215
      Begin CyberSentry.ProgressCntrl pbOutgoing 
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Firewall"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   600
         TabIndex        =   7
         Top             =   525
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CyberSentry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   1755
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "frmMain.frx":AFD4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6435
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10134
            Text            =   "Normal"
            TextSave        =   "Normal"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3360
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   3240
   End
   Begin MSComctlLib.ImageList ilMonitor 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":107B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":171BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2760
      Top             =   3360
   End
   Begin MSComctlLib.ListView lvMonitor 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilMonitor"
      SmallIcons      =   "ilMonitor"
      ColHdrIcons     =   "ilMonitor"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Application"
         Object.Width           =   9596
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Security"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Menu mnuSecurity 
      Caption         =   "mnuSecurity"
      Visible         =   0   'False
      Begin VB.Menu mnuSecurityTerminate 
         Caption         =   "Terminate"
      End
      Begin VB.Menu mnuSecurityBlock 
         Caption         =   "Block"
      End
      Begin VB.Menu mnuSecurityAsk 
         Caption         =   "Ask"
      End
      Begin VB.Menu mnuSecurityAllow 
         Caption         =   "Allow"
      End
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

If Check1.Value = vbChecked Then
SaveSetting App.Title, "Settings", "HideWinServices", "True"
Else
SaveSetting App.Title, "Settings", "HideWinServices", "False"
End If
Firewall.Execute True
End Sub


Private Sub Check2_Click()

If Check1.Value = vbChecked Then
SaveSetting App.Title, "Settings", "InetOnlyServices", "True"
Else
SaveSetting App.Title, "Settings", "InetOnlyServices", "False"
End If
Firewall.Execute True
End Sub


Private Sub Form_Unload(Cancel As Integer)
If GetSetting(App.Title, "Settings", "WarnOnExit", "True") = "True" Then
frmMinimize.Show
End If
Cancel = True
Hide
SysTray.InTray = True
End Sub


Private Sub lvMonitor_ItemClick(ByVal item As MSComctlLib.ListItem)
If Len(item.Text) Then
PopupMenu mnuSecurity
End If

End Sub


Private Sub mnuSecurityAllow_Click()
lvMonitor.SelectedItem.ListSubItems(1).ReportIcon = 4
lvMonitor.SelectedItem.ListSubItems(1).Text = "Allow"
SaveSetting App.Title, "Rules", lvMonitor.SelectedItem.Text, "Allow"
Firewall.Execute True
End Sub


Private Sub mnuSecurityAsk_Click()
lvMonitor.SelectedItem.ListSubItems(1).ReportIcon = 2
lvMonitor.SelectedItem.ListSubItems(1).Text = "Ask"
SaveSetting App.Title, "Rules", lvMonitor.SelectedItem.Text, "Ask"
Firewall.Execute True
End Sub


Private Sub mnuSecurityBlock_Click()
lvMonitor.SelectedItem.ListSubItems(1).ReportIcon = 3
lvMonitor.SelectedItem.ListSubItems(1).Text = "Block"
SaveSetting App.Title, "Rules", lvMonitor.SelectedItem.Text, "Block"
Firewall.Execute True
Firewall.Execute True
End Sub
Private Sub mnuSecurityTerminate_Click()
lvMonitor.SelectedItem.ListSubItems(1).ReportIcon = 3
lvMonitor.SelectedItem.ListSubItems(1).Text = "Terminate"
SaveSetting App.Title, "Rules", lvMonitor.SelectedItem.Text, "Terminate"
Firewall.Execute True
End Sub




Private Sub Option1_Click()
If Option1.Caption <> "Enable Firewall" Then
Option1.Caption = "Enable Firewall"
Timer1.Enabled = False
Else
Option1.Value = False
Option1.Caption = "Disable Firewall"
Timer1.Enabled = True
End If
End Sub

Private Sub Slider1_Change()

Select Case Slider1.Value
Case 1
lblSecurity.Caption = "Security Level: Low. Use with caution! This setting is the same as normal but does not ask about new connections."

Case 2
lblSecurity.Caption = "Security Level: Normal. This setting is ideal for most systems. Newly installed firewalls come default with this setting."

Case 3
lblSecurity.Caption = "Security Level: High. If you believe a hacker is present on the system, use this security level. It blocks all inbound connections."

End Select


End Sub

Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
Show

End Sub

Private Sub SysTray_MouseDown(Button As Integer, Id As Long)
If Button = 2 Then
PopupMenu frmMinimize.mnuSysTray
End If

End Sub


Private Sub Timer1_Timer()
Firewall.Execute
End Sub


Private Sub Timer3_Timer()
CloseAlert = ""
End Sub


Private Sub Timer4_Timer()

End Sub


