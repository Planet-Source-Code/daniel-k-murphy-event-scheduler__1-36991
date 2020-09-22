VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Date"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24444929
      CurrentDate     =   37131
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    ' ----------------------------------------
    ' Set the date for the calendar
    ' ----------------------------------------
    mvDate.Value = Date
    
    ' ----------------------------------------
    ' Set focus to the OK Button
    ' ----------------------------------------
    cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    ' ---------------------
    ' Set tag="Cancel"
    ' ---------------------
    Me.Tag = "Cancel"
    
    ' ---------------------
    ' Hide Form
    ' ---------------------
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    ' ---------------------
    ' Set tag="OK"
    ' ---------------------
    Me.Tag = "Ok"
    
    ' ---------------------
    ' Hide Form
    ' ---------------------
    Me.Hide
End Sub

Public Function GetDate() As String
    Dim strValue As String
    
    ' ---------------------
    ' Open the form
    ' ---------------------
    frmDate.Show vbModal
    
    ' ---------------------
    ' Check the tag to see if Cancel
    ' ---------------------
    If Trim(LCase(Me.Tag)) = "cancel" Then
        strValue = ""
    Else
        strValue = Trim(mvDate.Value)
    End If
    
    ' ---------------------
    ' Clear the tag
    ' ---------------------
    Me.Tag = ""
    
    ' ---------------------
    ' Remove the form
    ' ---------------------
    Unload Me
    
    ' ---------------------
    ' Return the value
    ' ---------------------
    GetDate = strValue
End Function

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    ' ----------------------------------------
    ' Since the Calendar doesn't work quite
    ' right after getting a click (or focus) then
    ' attempting to click on another object --
    ' you have to double-click -- This will make
    ' the clickability "windows natural" by not
    ' having to double click where you are
    ' supposed to single-click
    ' ----------------------------------------
    cmdOK.SetFocus
End Sub
