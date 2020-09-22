VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer 
      Interval        =   2000
      Left            =   240
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6345
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Programmed By:  Daniel K Murphy"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   1
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   1200
         Width           =   2325
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ProductName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2835
         TabIndex        =   3
         Top             =   705
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------
' Private variables
' ----------------------------------------
Private blnSplash As Boolean

Private Sub cmdOK_Click()
    ' ----------------------------------------
    ' Unload the form
    ' ----------------------------------------
    Unload Me
End Sub

Private Sub Form_Activate()
    ' ----------------------------------------
    ' Enable/Disable the Timer
    ' ----------------------------------------
    If blnSplash = True Then
        ' In Splash Mode
        tmrTimer.Enabled = True
        cmdOK.Visible = False
    Else
        ' In About Mode
        tmrTimer.Enabled = False
        cmdOK.Visible = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' ----------------------------------------
    ' Unload the form
    ' ----------------------------------------
    Unload Me
End Sub

Private Sub Form_Load()
    ' ----------------------------------------
    ' Populate the labels with the info
    ' ----------------------------------------
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    imgLogo.Picture = frmMain.Icon
    lblCopyright.Caption = "Copyright 2002"
    
    ' ----------------------------------------
    ' Resize the labels to the same size
    ' ----------------------------------------
    lblVersion.Left = lblProductName.Left
    lblVersion.Width = lblProductName.Width
    lblCopyright.Left = lblProductName.Left
    lblCopyright.Width = lblProductName.Width
    
    ' ----------------------------------------
    ' Set default as Splash
    ' ----------------------------------------
    blnSplash = True
End Sub

Private Sub Frame1_Click()
    ' ----------------------------------------
    ' Unload the form
    ' ----------------------------------------
    Unload Me
End Sub

Public Sub ShowForm(Optional Splash As Boolean = True)
    ' ----------------------------------------
    ' Show the form
    ' ----------------------------------------
    blnSplash = Splash
    Me.Show
End Sub

Private Sub tmrTimer_Timer()
    ' ----------------------------------------
    ' Unload the form
    ' ----------------------------------------
    Unload Me
End Sub
