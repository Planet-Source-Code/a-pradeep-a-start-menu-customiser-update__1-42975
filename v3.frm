VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Madness For Windows V 2.1.1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Restore"
      Height          =   375
      Left            =   720
      TabIndex        =   35
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CheckBox Check28 
      Caption         =   "To Context"
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Damn the Windows !"
      Height          =   1935
      Left            =   3840
      TabIndex        =   22
      Top             =   2880
      Width           =   3015
      Begin VB.CommandButton Command3 
         Caption         =   " "
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox Check27 
         Caption         =   "Text in IE title bar"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check26 
         Caption         =   "No Windows Hotkeys"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Hide User Profiles PAge"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Hide remote Adminstration page"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Hide Change PAsswords !"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Network options"
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   3615
      Begin VB.CheckBox Check24 
         Caption         =   "No map and Disconnect Network"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox Check23 
         Caption         =   "No Entire Network !"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Hide Network Neighbourhood !"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Disable Print Sharing"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Disable File Sharing"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About/Help"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Options"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox Check25 
         Caption         =   "No Startup Disk page"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Disable change Control panel password !"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   2160
         Width           =   3495
      End
      Begin VB.CheckBox Check18 
         Caption         =   "No Hardware Profiles in System"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Disable Dos (workx in Win 95 & 98 not sure on others)"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Disable Update of drivers button"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Hide virtual memory button"
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox Check11 
         Caption         =   "No file system button in System"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox Check10 
         Caption         =   "No Display Settings"
         Height          =   195
         Left            =   2880
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check9 
         Caption         =   "No Screen Saver page"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Check8 
         Caption         =   "No back Ground in Display"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox Check7 
         Caption         =   "No Appearance in Display"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         Caption         =   "No Device manager page"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Hide Configuration page"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Disable Task Bar movement "
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         Caption         =   "No Printer Tabs "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Disable Display Control"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Disable Registy editing Toolz"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
CreateKey "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\System"
SetDWORDValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1
ElseIf Check1.Value = 0 Then
DeleteKey "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\System"
End If
End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then
system "NoDispSettingsPage", Check10
End If
End Sub

Private Sub Check11_Click()
If Check11.Value = 1 Then
system "NoFileSysPage", Check11
End If
End Sub

Private Sub Check12_Click()
If Check12.Value = 1 Then
system "NoVirtMemPage", Check12
End If
End Sub

Private Sub Check14_Click()
If Check14.Value = 1 Then
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 1
ElseIf Check14.Value = 0 Then
RegDeleteKey &H80000001, "software\microsoft\windows\currentversion\policies\winoldapp"
End If
End Sub

Private Sub Check15_Click()
If Check15.Value = 1 Then
network "NoFileSharingControl", Check15
End If
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then
network "NoPrintSharingControl", Check16
End If
End Sub

Private Sub Check17_Click()
If Check17.Value = 1 Then
Madness "NoNetHood", Check17
End If
End Sub

Private Sub Check18_Click()
If Check18.Value = 1 Then
system "NoConfigPage", Check18
End If
End Sub

Private Sub Check19_Click()
If Check19.Value = 1 Then
system "NoSecCPL", Check19
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
system "NoDispCPL", Check2
End If
End Sub

Private Sub Check20_Click()
If Check20.Value = 1 Then
system "NoPWDPage", Check20
End If
End Sub

Private Sub Check21_Click()
If Check21.Value = 1 Then
system "NoAdminPage", Check21
End If
End Sub

Private Sub Check22_Click()
If Check22.Value = 1 Then
system "NoProfilePage", Check22
End If
End Sub

Private Sub Check23_Click()
If Check23.Value = 1 Then
network "NoEntireNetwork", Check23
End If
End Sub

Private Sub Check24_Click()
If Check24.Value = 1 Then
Madness "NoNetConnectDisconnect", Check24
End If
End Sub

Private Sub Check25_Click()
If Check25.Value = 1 Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Setup", "EBDPage", ""
End If
End Sub

Private Sub Check26_Click()
If Check26.Value = 1 Then
Madness "NoWinKeys", Check26
End If
End Sub

Private Sub Check27_Click()
If Check27.Value = 1 Then
Text1.Visible = True
Command3.Value = True
Text1.Text = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Window Title")
ElseIf Check27.Value = 0 Then
Text1.Visible = False
Command3.Value = False
End If
End Sub

Private Sub Check28_Click()
If Check28.Value = 1 Then
CreateKey ("HKEY_CLASSES_ROOT\Folder\shell\Policy update")
CreateKey ("HKEY_CLASSES_ROOT\Folder\shell\Policy update\command")
SetStringValue "HKEY_CLASSES_ROOT\Folder\shell\Policy update\command", "", App.Path & "\" & App.EXEName & ".exe"
ElseIf Check28.Value = 0 Then
RegDeleteKey &H80000000, "folder\shell\policy update"
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Madness "NoPrinterTabs", Check3
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Madness "NoSetTaskBar", Check4
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
system "NoDevMgrPage", Check6
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
system "NoDispAppearancePage", Check7
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
system "NoDispBackgroundPage", Check8
End If
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
system "NoDispScrSavPage", Check9
End If
End Sub

Private Sub Command1_Click()
MsgBox "Thanks for trying my code, it comes without any warranty and    I am not responsable for your acts ,   I hope you like this application please rate both of my codes so that i can send you the final release of Madness for Windows as soon as it's done along with huge help files if you know what i mean :)", vbOKOnly
End Sub

Public Function Madness(name As String, obj As Object)
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", name, obj.Value
End Function

Public Function network(name As String, obj As Object)
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\network", name, obj.Value
End Function

Public Function system(name As String, obj As Object)
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\system", name, obj.Value
End Function

Private Sub Command3_Click()
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Window Title", Text1.Text
Text1.Visible = False
Command3.Visible = False
End Sub

Private Sub Command4_Click()
RegDeleteKey &H80000001, "software\microsoft\windows\currentversion\policies"
End Sub

Private Sub Form_Load()
Text1.Visible = False
Command3.Value = False
End Sub
