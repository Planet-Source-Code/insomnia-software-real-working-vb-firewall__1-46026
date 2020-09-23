Attribute VB_Name = "Firewall"
Global strSysFiles As String


Private Type Stats_
LocP As String
RemP As String
RemA As String
PID As Long
Inbound As Boolean
End Type
Public stats(0 To 2000) As Stats_
Public StatsLen As Long
Public IdLen As Long
Public lvString As String

Public CloseAlert As String
Public OnTop As New clsOnTop
Public ActionAlert As String
Public Sub Alert(ApplicationName, Host, Port, Optional Inbound As Boolean, Optional PID)
Dim frmA As New frmAlertApp
mdlProcess.SuspendThreads CLng(PID)
If Host = "localhost" Then Inbound = True
If Inbound = True Then
x = "allowing a connection from"
y = "on port"
Else
x = "connecting to"
y = "to port"
End If
frmA.Timer1.Tag = ApplicationName
frmA.Tag = ApplicationName
frmA.lblDesc.Tag = CLng(PID)
frmA.lblDesc = "The program " & ApplicationName & " is " & x & " " & Host & " " & y & " " & Port & "." & " This program may or may not be using malicious activity. Click BLOCK to block that application from using the internet. Click ALLOW to allow that application to use the internet this time. Click TRUST to always allow that program to use the internet. If you believe that program is a virus or trojan, click ELIMINATE."
frmA.Show
frmA.Timer1.Enabled = True
OnTop.MakeTopMost frmA.hwnd
End Sub


Public Sub Block(LocP, RemA)
On Error Resume Next
Dim tcpt As MIB_TCPTABLE
Dim l As Long
Dim x As Long
Dim b As Boolean, a As Boolean
Dim lvl As Long
For i = 0 To 2000
If Ports(i).LocP = LocP And Ports(i).RemA = RemA Then
        l = Len(MIB_TCPTABLE)
        GetTcpTable tcpt, l, 0
    
        tcpt.table(i).dwState = 12
        SetTcpEntry tcpt.table(i)
        DoEvents
Exit Sub
End If
DoEvents
Next i

End Sub

'Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Sub Execute(Optional force As Boolean)
Dim tmpString As String, temp2 As String
Dim strRepeat As String
Dim InData As String, item As ListItem, subitem As ListSubItem
Dim v As Long, b As Long, m As Long, n As Long
'// First Check if there was a change
On Error Resume Next
If Refresh = False And force = False Then Exit Sub
'// Opens command prompt to find connections with id's
InData = Cmd.Execute("netstat -o")
'// Parce command prompt return
ParseData InData
'// Load running executables and their id number
SetUpList
'// Now Update the list!



frmMain.lvMonitor.ListItems.Clear
        frmMain.pbBlocked.Max = 1
        frmMain.pbIncoming.Max = 1
        frmMain.pbOutgoing.Max = 1
        frmMain.pbBlocked.Caption = "Blocked Traffic: " & 0
        frmMain.pbIncoming.Caption = "Incoming Traffic: " & 0
        frmMain.pbOutgoing.Caption = "Outgoing Traffic: " & 0
        frmMain.pbBlocked.Value = 0
        frmMain.pbIncoming.Value = 0
        frmMain.pbOutgoing.Value = 0
lvString = ""

Dim HideWinServices As Boolean, InetOnlyServices As Boolean
HideWinServices = frmMain.Check1.Value
InetOnlyServices = frmMain.Check2.Value


For i = 0 To IdLen
    For y = 0 To IdLen
        If stats(i).PID = Id(y).ProcessNumber Then 'And LastPID <> stats(i).PID Then
            If InStr(1, lvString, "$" & stats(i).PID & "$") Then GoTo 1
             
             'tmpString = Replace(UCase(Id(y).Filename), "\SYSTEM32", "\SYSTEM")
             strRepeat = strRepeat & stats(i).PID & vbCrLf
             
            
            lvString = lvString & "$" & stats(i).PID & "$"
            LastPID = stats(i).PID
            'ResumeThreads stats(i).PID
            
            
            
            tmpString = Id(y).Filename
            tmpString = UCase(tmpString)
            tmpString = Replace(tmpString, "\SYSTEM32", "\SYSTEM")
            temp2 = UCase(strSysFiles)
            temp2 = Replace(temp2, "\SYSTEM32", "\SYSTEM")
            
            If InStr(1, temp2, tmpString) And HideWinServices = True Then GoTo 11
            
            
            
            Set item = frmMain.lvMonitor.ListItems.Add()
11:
            item.Text = Id(y).Filename
            item.SmallIcon = 1
            item.Icon = 1
            item.ForeColor = vbBlue
            Set subitem = item.ListSubItems.Add()

            Q = GetSetting(App.Title, "Rules", Id(y).Filename, "Ask")
            Select Case Q
            Case "Ask"
            If stats(i).Inbound = True Then m = m + 1 Else n = n + 1
            subitem.Text = "Ask"
            subitem.ReportIcon = 2
            If Not stats(i).RemA = "" Then
            
            Alert Id(y).Filename, stats(i).RemA, stats(i).RemP, False, LastPID
            End If
            Case "Block"
            SuspendThreads stats(i).PID
            subitem.Text = "Block"
            subitem.ReportIcon = 3
            Block stats(i).LocP, stats(i).RemA
            ResumeThreads stats(i).PID
            b = b + 1
            Case "Terminate"
            SuspendThreads stats(i).PID
            subitem.Text = "Terminate"
            subitem.ReportIcon = 3
            Terminate Id(i).ProcessNumber
            b = b + 1
            Case "Allow"
            
            subitem.Text = "Allow"
            subitem.ReportIcon = 4
            If stats(i).Inbound = True Then m = m + 1 Else n = n + 1
            End Select
            v = v + 1
            
            
            
1:
'Debug.Print hey
        End If
        frmMain.pbBlocked.Max = v
        frmMain.pbIncoming.Max = v
        frmMain.pbOutgoing.Max = v
        frmMain.pbBlocked.Caption = "Blocked Traffic: " & b
        frmMain.pbIncoming.Caption = "Incoming Traffic: " & m
        frmMain.pbOutgoing.Caption = "Outgoing Traffic: " & n
        frmMain.pbBlocked.Value = b
        frmMain.pbIncoming.Value = m
        frmMain.pbOutgoing.Value = n
    Next y
Next i




'Now just add normal services
For i = 0 To IdLen
If Not InStr(1, lvString, "$" & Id(i).ProcessNumber & "$") Then
             'tmpString = Replace(UCase(Id(i).Filename), "\SYSTEM32", "\SYSTEM")
                 If InStr(1, strRepeat, Id(i).ProcessNumber) Then GoTo 23

            
lvString = lvString & "$" & Id(i).ProcessNumber & "$"




            tmpString = Id(i).Filename
            tmpString = UCase(tmpString)
            tmpString = Replace(tmpString, "\SYSTEM32", "\SYSTEM")
            temp2 = UCase(strSysFiles)
            temp2 = Replace(temp2, "\SYSTEM32", "\SYSTEM")
            
            If InStr(1, temp2, tmpString) And HideWinServices = True Then GoTo 12




             If InetOnlyServices = True Then GoTo 12
            Set item = frmMain.lvMonitor.ListItems.Add()
12:
            item.Text = Id(i).Filename
            item.SmallIcon = 1
            item.Icon = 1
         
            Set subitem = item.ListSubItems.Add()

            Q = GetSetting(App.Title, "Rules", Id(i).Filename, "Ask")
        Select Case Q
            Case "Ask"
                subitem.Text = "Ask"
                subitem.ReportIcon = 2
            Case "Block"
                subitem.Text = "Block"
                subitem.ReportIcon = 3
                'Block stats(i).LocP, stats(i).RemA
            Case "Terminate"
                        SuspendThreads Id(i).ProcessNumber
                subitem.Text = "Terminate"
                subitem.ReportIcon = 3
                Terminate Id(i).ProcessNumber
            Case "Allow"
                subitem.Text = "Allow"
                subitem.ReportIcon = 4
        End Select
    End If
23:
Next i

End Sub
Public Sub ParseData(Data As String)
'MsgBox Data
Data = Replace(Data, "  ", " ")
Dim SplitUp() As String, LocP As String, PID As String, RemA As String, RemP As String
Dim LineSplit() As String, y As Long
SplitUp = Split(Data, vbCrLf)
For i = 0 To UBound(SplitUp)

LineSplit = Split(SplitUp(i), " ")
On Error Resume Next
If LineSplit(1) = "TCP" Then
PID = ""
LocA = Mid(LineSplit(3), 1, InStr(1, LineSplit(3), ":"))
LocP = Mid(LineSplit(3), InStr(1, LineSplit(3), ":") + 1, Len(LineSplit(3)) - InStr(1, LineSplit(3), ":") + 1)
RemA = Mid(LineSplit(9), 1, InStr(1, LineSplit(9), ":"))
RemP = Mid(LineSplit(9), InStr(1, LineSplit(9), ":") + 1, Len(LineSplit(9)) - InStr(1, LineSplit(9), ":"))
PID = LineSplit(15)
'If Not PID + 10000 > 0 Then PID = LineSplit(18)
If PID <> "" Then
stats(y).LocP = LocP
stats(y).PID = PID
ProcessId(y) = PID
stats(y).RemA = Replace(RemA, ":", "")
stats(y).RemP = RemP

StatsLen = y

End If
y = y + 1
End If
Next i
End Sub


Private Sub SetUpList()
    'With lvwProcess
'        .ListItems.Clear
       ' With .ColumnHeaders
       '     .Clear
        '    .Add , , "Process", (lvwProcess.Width * (0.8))
        '    .Add , , "PID", (lvwProcess.Width * (0.175))
       ' End With
    
       ' .View = lvwReport
        '.HideColumnHeaders = False
    'End With
        
    Select Case getVersion()
        Case WIN98 'Windows 95/98
            MsgBox "Windows 95 and 98 are not supported!"
            End
        Case WINNT 'Windows NT
            LoadNTProcess
    End Select
End Sub







Public Sub Terminate(PID As Long)
KillProcessById PID
End Sub


