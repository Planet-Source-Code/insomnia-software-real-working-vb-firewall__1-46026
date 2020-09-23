VERSION 5.00
Begin VB.UserControl ProgressCntrl 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ScaleHeight     =   390
   ScaleWidth      =   4710
   Begin VB.PictureBox MainBox 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Progress 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   15
         Begin VB.Label Stat2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2175
            TabIndex        =   3
            Top             =   60
            Width           =   465
         End
      End
      Begin VB.Label Stat1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2175
         TabIndex        =   1
         Top             =   60
         Width           =   465
      End
   End
End
Attribute VB_Name = "ProgressCntrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ProgVal As Integer
Dim MaxNum As Long

Public Property Let Max(lngNum As Long)
MaxNum = lngNum
End Property

Public Property Get Max() As Long
Max = MaxNum
End Property
Public Property Let Value(IntValue As Long)
On Error Resume Next
If IntValue = 0 Then
Progress.Visible = False
Else
Progress.Visible = True
End If
ProgVal = IntValue

Progress.Width = MainBox.Width * (ProgVal / MaxNum)
Refresh
End Property

Public Property Get Value() As Long
ProgVal = Value
End Property

Public Property Let Caption(MyCaption As String)
Stat1 = MyCaption
Stat2 = MyCaption
End Property

Public Property Get Caption() As String
Caption = Stat1
End Property

Private Sub UserControl_Initialize()
Progress.Visible = False
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
MainBox.Width = UserControl.Width
MainBox.Height = UserControl.Height
Stat1.Left = 50 '(MainBox.Width / 2) - (Stat1.Width / 2)
Stat1.Top = (MainBox.Height / 2) - (Stat1.Height / 2) - 30
Stat2.Left = 50
Stat2.Top = Stat1.Top
Progress.Height = MainBox.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MainBox,MainBox,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Progress.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Progress.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Stat1,Stat1,-1,ForeColor
Public Property Get Forecolor1() As OLE_COLOR
Attribute Forecolor1.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    Forecolor1 = Stat1.ForeColor
End Property

Public Property Let Forecolor1(ByVal New_Forecolor1 As OLE_COLOR)
    Stat1.ForeColor() = New_Forecolor1
    PropertyChanged "Forecolor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Stat2,Stat2,-1,ForeColor
Public Property Get Forecolor2() As OLE_COLOR
Attribute Forecolor2.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    Forecolor2 = Stat2.ForeColor
End Property

Public Property Let Forecolor2(ByVal New_Forecolor2 As OLE_COLOR)
    Stat2.ForeColor() = New_Forecolor2
    PropertyChanged "Forecolor2"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    MainBox.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Stat1.ForeColor = PropBag.ReadProperty("Forecolor1", &H80000012)
    Stat2.ForeColor = PropBag.ReadProperty("Forecolor2", &HFFFFFF)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", MainBox.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Forecolor1", Stat1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Forecolor2", Stat2.ForeColor, &HFFFFFF)
End Sub

