VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAbout 
         Caption         =   "A&bout..."
      End
      Begin VB.Menu mnuFileSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bytCounter As Byte  ' Used to track Minute Cycles

Private Sub Form_Load()
    ' ----------------------------------------
    ' Use the same Icon as frmMain
    ' ----------------------------------------
    Me.Icon = frmMain.Icon
    
    ' ----------------------------------------
    ' Grab the handle to this form
    ' to use in the System Tray
    ' ----------------------------------------
    Hook Me.hwnd
    
    ' ----------------------------------------
    ' Place the Icon into the System Tray
    ' ----------------------------------------
    AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, App.Title
    
    ' ----------------------------------------
    ' Hide this form
    ' ----------------------------------------
    Me.Hide
    
    ' ----------------------------------------
    ' Run this hour's events
    ' ----------------------------------------
    RunScheduledEvents
    
    ' ----------------------------------------
    ' Set the Counter to 0
    ' ----------------------------------------
    bytCounter = 0
End Sub

Private Sub mnuExit_Click()
    ' ----------------------------------------
    ' Remove the Icon from the System Tray
    ' ----------------------------------------
    RemoveIconFromTray
    
    ' ----------------------------------------
    ' Terminate the application
    ' ----------------------------------------
    Terminate
End Sub

Public Sub SysTrayMouseEventHandler()
    ' ----------------------------------------
    ' Detect the mouse, making sure the application is
    ' Right-Clickable via the System Tray
    ' ----------------------------------------
    SetForegroundWindow Me.hwnd
    PopupMenu mnuFile, vbPopupMenuRightButton
End Sub

Private Sub mnuFileAbout_Click()
    ' ----------------------------------------
    ' Show the Splash screen in About mode
    ' ----------------------------------------
    Load frmSplash
    frmSplash.ShowForm False
End Sub

Private Sub mnuFileSettings_Click()
    ' ----------------------------------------
    ' Load/Show the frmMain
    ' ----------------------------------------
    Load frmMain
    
    frmMain.Show
End Sub

Private Sub tmrTimer_Timer()
    ' ----------------------------------------
    ' This will run the scheduled events
    ' ----------------------------------------
    If bytCounter = 59 Then
        ' ----------------------------------------
        ' Run the events
        ' ----------------------------------------
        RunScheduledEvents
        
        ' ----------------------------------------
        ' Reset Counter
        ' ----------------------------------------
        bytCounter = 0
    Else
        ' ----------------------------------------
        ' Add one to the Counter
        ' ----------------------------------------
        bytCounter = bytCounter + 1
    End If
End Sub
