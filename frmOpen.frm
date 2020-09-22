VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select File"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFileType 
      Height          =   315
      ItemData        =   "frmOpen.frx":0000
      Left            =   3240
      List            =   "frmOpen.frx":000D
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin VB.FileListBox fileBox 
      Height          =   2625
      Left            =   3240
      Pattern         =   "*.EXE; *.COM"
      TabIndex        =   2
      Top             =   450
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.DirListBox dirBox 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.DriveListBox drvBox 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GetFile() As String
    Dim strValue As String
    
    ' ----------------------------------------
    ' Open the form
    ' ----------------------------------------
    frmOpen.Show vbModal
    
    ' ----------------------------------------
    ' Check the tag to see if Cancel
    ' ----------------------------------------
    If Trim(LCase(Me.Tag)) = "cancel" Then
        strValue = ""
    Else
        strValue = Trim(fileBox.Path) + "\" + Trim(fileBox.FileName)
    End If
    
    ' ----------------------------------------
    ' Clear the tag
    ' ----------------------------------------
    Me.Tag = ""
    
    ' ----------------------------------------
    ' Remove the form
    ' ----------------------------------------
    Unload Me
    
    ' ----------------------------------------
    ' Return the value
    ' ----------------------------------------
    GetFile = strValue
End Function

Private Sub cboFileType_Click()
    ' ----------------------------------------
    ' This will change the file type
    ' and will automatically fix the
    ' viewing type in the FileListBox
    ' ----------------------------------------
    Select Case cboFileType.ListIndex
        Case 0
            ' *.EXE; *.COM
            fileBox.Pattern = "*.EXE; *.COM"
        Case 1
            ' *.BAT
            fileBox.Pattern = "*.BAT"
        Case 2
            ' *.*
            fileBox.Pattern = "*.*"
    End Select
    
    ' ----------------------------------------
    ' Refresh to show the files
    ' ----------------------------------------
    fileBox.Refresh
End Sub

Private Sub cmdCancel_Click()
    ' ----------------------------------------
    ' Set tag="Cancel"
    ' ----------------------------------------
    Me.Tag = "Cancel"
    
    ' ----------------------------------------
    ' Hide Form
    ' ----------------------------------------
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    ' ----------------------------------------
    ' Set tag="OK"
    ' ----------------------------------------
    Me.Tag = "Ok"
    
    ' ----------------------------------------
    ' Hide Form
    ' ----------------------------------------
    Me.Hide
End Sub

Private Sub dirBox_Click()
    ' ----------------------------------------
    ' This will update the FileListBox to
    ' show the contents of the selected directory
    ' ----------------------------------------
    fileBox.Path = dirBox.Path
    
    ' ----------------------------------------
    ' Refresh the FileListBox to show the contents
    ' ----------------------------------------
    fileBox.Refresh
End Sub

Private Sub dirBox_Change()
    ' ----------------------------------------
    ' This will update the FileListBox to
    ' show the contents of the selected directory
    ' ----------------------------------------
    fileBox.Path = dirBox.Path
    
    ' ----------------------------------------
    ' Refresh the FileListBox to show the contents
    ' ----------------------------------------
    fileBox.Refresh
End Sub

Private Sub drvBox_Change()
    ' ----------------------------------------
    ' This will set the directorylistbox
    ' path for the drive selected
    ' ----------------------------------------
    dirBox.Path = drvBox.Drive
End Sub

Private Sub Form_Load()
    ' ----------------------------------------
    ' Set up the Form Defaults
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' Clear the Combobox
    ' ----------------------------------------
    cboFileType.Clear
    
    ' ----------------------------------------
    ' Populate the file type options
    ' ----------------------------------------
    cboFileType.AddItem "*.EXE; *.COM (Executable Files)"
    cboFileType.AddItem "*.BAT (Batch Files)"
    cboFileType.AddItem "*.* (All Files)"
    
    ' ----------------------------------------
    ' Set the default for the File Type
    ' ----------------------------------------
    cboFileType.SelText = "*.EXE; *.COM (Executable Files)"
End Sub
