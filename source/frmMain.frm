VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorun"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctHiddenIcon 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame frmDescription 
      Caption         =   "Description Title"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmMain.frx":0442
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ListBox lstItems 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLaunch_Click()
    On Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    
    mnuexec = sGetINI(App.Path + "\autorun.inf", lstItems.Text, "Exec", "?")
    If Not mnuexec = "?" Then
        If InStr(mnuexec, "http://") = 1 Or InStr(mnuexec, "https://") = 1 Then
            foo = mnuexec
            Shell GetURLCommand(foo)
        Else
            Shell mnuexec, vbNormalFocus
        End If
    End If
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    
    lstItems.Clear
    txtDescription.Text = ""
    frmDescription.Caption = ""
    
    buttonlaunchtitle = sGetINI(App.Path + "\autorun.inf", "General", "LaunchTitle", "?")
    If Not buttonlaunchtitle = "?" Then
        cmdLaunch.Caption = buttonlaunchtitle
    End If
    buttonexittitle = sGetINI(App.Path + "\autorun.inf", "General", "ExitTitle", "?")
    If Not buttonexittitle = "?" Then
        cmdExit.Caption = buttonexittitle
    End If
    autoruncaption = sGetINI(App.Path + "\autorun.inf", "General", "Title", "?")
    If Not autoruncaption = "?" Then
        frmMain.Caption = autoruncaption
    End If
    autorunicon = sGetINI(App.Path + "\autorun.inf", "General", "Icon", "?")
    If Not autorunicon = "?" Then
        frmMain.Icon = LoadPicture(autorunicon)
    End If
    pctHiddenIcon.Picture = frmMain.Icon
    
    idx = 1
    Do
        mnutitle = sGetINI(App.Path + "\autorun.inf", "MenuItems", "Item" + Trim(Str(idx)), "?")
        If mnutitle = "?" Then
            Exit Do
        End If
        lstItems.AddItem mnutitle
        idx = idx + 1
    Loop
    
    autorunwelcometext = sGetINI(App.Path + "\autorun.inf", "General", "HelpText", "?")
    If Not autorunwelcometext = "?" Then
        frmDescription.Caption = frmMain.Caption
        txtDescription.Text = autorunwelcometext
    End If
    
End Sub


Private Sub lstItems_Click()
    cmdLaunch.Enabled = True
    frmDescription.Caption = lstItems.Text
    mnuhelptext = sGetINI(App.Path + "\autorun.inf", lstItems.Text, "HelpText", "?")
    If mnuhelptext = "?" Then
      txtDescription.Text = ""
    Else
      txtDescription.Text = Replace(mnuhelptext, "|", vbCrLf, 1, Len(mnuhelptext), vbTextCompare)
    End If
    itemicon = sGetINI(App.Path + "\autorun.inf", lstItems.Text, "Icon", "?")
    If itemicon = "?" Then
        frmMain.Icon = pctHiddenIcon.Picture
    Else
        frmMain.Icon = LoadPicture(itemicon)
    End If
End Sub


