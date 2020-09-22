VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   130
      ScaleHeight     =   0
      ScaleWidth      =   5820
      TabIndex        =   6
      Top             =   350
      Width           =   5845
   End
   Begin VB.CommandButton cmdRunNow 
      Caption         =   "Run Now"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstEvents 
      BackColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngTabStops As Long      ' Used to setup the TabStops in the Listbox

Private Sub cmdEdit_Click()
    ' ----------------------------------------
    ' Edit the specified Event
    ' ----------------------------------------
    Load frmAddEdit
    
    frmAddEdit.ShowForm "EDIT\" & Trim(Str(GetEventID(GetEventName(Trim(lstEvents.List(lstEvents.ListIndex))))))
End Sub

Private Sub cmdNew_Click()
    ' ----------------------------------------
    ' Open the Add/Edit for in Add mode
    ' ----------------------------------------
    Load frmAddEdit
    
    frmAddEdit.ShowForm "ADD"
End Sub

Private Sub cmdRemove_Click()
    ' ----------------------------------------
    ' This will remove the selected event
    ' ----------------------------------------
    Dim strSQL As String
    
    Dim lngEvent
    
    ' ----------------------------------------
    ' Ask to make sure
    ' ----------------------------------------
    If MsgBox("This will pernamently remove this Event." & Chr$(13) & Chr$(13) & "Are you sure?", vbYesNo, App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision) = vbYes Then
        ' ----------------------------------------
        ' Get Event ID
        ' ----------------------------------------
        lngEvent = GetEventID(GetEventName(Trim(lstEvents.List(lstEvents.ListIndex))))
    
        strSQL = "DELETE * FROM [tblEVENT] WHERE [EVENT_ID]=" & Trim(Str(lngEvent))
    
        dbDatabase.Execute strSQL
    
        ' ----------------------------------------
        ' Re-Load events
        ' ----------------------------------------
        LoadEvents
        
        ' ----------------------------------------
        ' Disable controls
        ' ----------------------------------------
        Reset
    End If
End Sub

Private Sub cmdRunNow_Click()
    ' ----------------------------------------
    ' This will launch the Event in Normal Mode
    ' ----------------------------------------
    Dim lngEvent As Long
    
    ' ----------------------------------------
    ' Get the Event ID
    ' ----------------------------------------
    lngEvent = GetEventID(GetEventName(Trim(lstEvents.List(lstEvents.ListIndex))))
    
    ' ----------------------------------------
    ' Run The Event
    ' ----------------------------------------
    RunEvent lngEvent, False, False
End Sub

Private Sub cmdSettings_Click()
    ' ----------------------------------------
    ' This will launch the Event in Settings Mode
    ' ----------------------------------------
    Dim lngEvent As Long
    
    ' ----------------------------------------
    ' Get the Event ID
    ' ----------------------------------------
    lngEvent = GetEventID(GetEventName(Trim(lstEvents.List(lstEvents.ListIndex))))
    
    ' ----------------------------------------
    ' Run The Event
    ' ----------------------------------------
    RunEvent lngEvent, False, True
End Sub

Private Sub Form_Activate()
    ' ----------------------------------------
    ' Load the events
    ' ----------------------------------------
    LoadEvents
    
    ' ----------------------------------------
    ' Reset the controls
    ' ----------------------------------------
    Reset
End Sub

Private Sub Form_Load()
    ' ----------------------------------------
    ' Establish defaults
    ' ----------------------------------------
    Me.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' ----------------------------------------
    ' Clear all variables/references
    ' ----------------------------------------
    Set dbDatabase = Nothing
End Sub

Private Sub LoadEvents()
    ' ----------------------------------------
    ' This will load the Event Names
    ' and populate them in the list box
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' Data variables
    ' ----------------------------------------
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim strFrequency As String
    
    ' ----------------------------------------
    ' Establish the database connection
    ' ----------------------------------------
    Set dbDatabase = New ADODB.Connection
    Set rsData = New ADODB.Recordset
    
    With dbDatabase
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\EventScheduler.mdb"
        .Open
    End With
    
    ' ----------------------------------------
    ' Clear the List box
    ' ----------------------------------------
    lstEvents.Clear

    ' ----------------------------------------
    ' Grab the LONGEST Name Length
    ' ----------------------------------------
    strSQL = "SELECT MAX(LEN([EVENT_NAME])) AS LENGTH FROM [tblEVENT]"
    
    With rsData
        .Open strSQL, dbDatabase
        
        If Not (.BOF And .EOF) Then
            ' ----------------------------------------
            ' Records found
            ' Calculate the tabstop based on
            ' # of characters times list box fontsize
            ' Then divided by One third (1.5)
            ' ----------------------------------------
            If IsNull(!LENGTH) Then
                ' Null
                lngTabStops = 40
            Else
                ' Not Null
                lngTabStops = CLng((!LENGTH * lstEvents.FontSize) / 1.5)
            End If
        Else
            ' ----------------------------------------
            ' No Records Found
            ' ----------------------------------------
            lngTabStops = 40
        End If
        
        .Close
    End With
    
    ' ----------------------------------------
    ' If the Header is LONGER than
    ' the Event Name(s), use the
    ' Header length for the tabstops
    ' ----------------------------------------
    If ((12 * lstEvents.FontSize) / 1.5) > lngTabStops Then
        lngTabStops = ((12 * lstEvents.FontSize) / 1.5)
    End If
        
    ' ----------------------------------------
    ' Clear any Tabstops
    ' ----------------------------------------
    Call SendMessage(lstEvents.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    
    ' ----------------------------------------
    ' Assign the TabStops to the List Box
    ' ----------------------------------------
    Call SendMessage(lstEvents.hwnd, LB_SETTABSTOPS, 1, lngTabStops)
    
    ' ----------------------------------------
    ' Refresh the List Box to update Tabstops
    ' ----------------------------------------
    lstEvents.Refresh
    
    ' ----------------------------------------
    ' Place the header
    ' ----------------------------------------
    lstEvents.AddItem "[EVENT NAME]" & vbTab & "| [FREQUENCY]"
    
    ' ----------------------------------------
    ' Get the records
    ' ----------------------------------------
    strSQL = "SELECT * FROM [tblEVENT] ORDER BY [EVENT_NAME]"
    
    With rsData
        .Open strSQL, dbDatabase
                            
        Do Until .EOF
            ' Reset to null
            strFrequency = ""
            
            Select Case !EVENT_FREQUENCY_ID
                Case 0
                    ' Daily
                    strFrequency = "Daily at " & Trim(Str(!EVENT_HOUR_OF_DAY)) & "00 Hours"
                Case 1
                    ' Daily (Non-Weekend)
                    strFrequency = "Week Days at " & Trim(Str(!EVENT_HOUR_OF_DAY)) & "00 Hours"
                Case 2
                    ' Weekly
                    ' Get the Day of the Week
                    Select Case !EVENT_DAY_OF_WEEK
                        Case 0
                            ' Monday
                            strFrequency = "Every Monday at "
                        Case 1
                            ' Tuesday
                            strFrequency = "Every Tuesday at "
                        Case 2
                            ' Wednesday
                            strFrequency = "Every Wednesday at "
                        Case 3
                            ' Thursday
                            strFrequency = "Every Thursday at "
                        Case 4
                            ' Friday
                            strFrequency = "Every Friday at "
                        Case 5
                            ' Saturday
                            strFrequency = "Every Saturday at "
                        Case 6
                            ' Sunday
                            strFrequency = "Every Sunday at "
                    End Select
                    
                    ' Finish out the Frequency message
                    strFrequency = strFrequency & Trim(Str(!EVENT_HOUR_OF_DAY)) & "00 Hours"
                Case 3
                    ' Monthly
                    strFrequency = "Every Month on day " & Trim(Str(!EVENT_DAY_OF_MONTH)) & " at " & Trim(Str(!EVENT_HOUR_OF_DAY)) & "00 Hours"
                Case 4
                    ' Once Only
                    strFrequency = "On " & Trim(Str(!EVENT_MONTH)) & "/" & Trim(Str(!EVENT_DAY_OF_MONTH)) & " at " & Trim(Str(!EVENT_HOUR_OF_DAY)) & "00 Hours"
                Case Else
                    ' Undefined
                    strFrequency = "Unknown"
            End Select
            
            ' Add the | before the strFrequency
            strFrequency = "| " & Trim(strFrequency)
            
            lstEvents.AddItem UCase(Trim(![EVENT_NAME]) & vbTab & strFrequency)
            
            ' Next record
            .MoveNext
        Loop
        
        .Close
    End With
    ' ----------------------------------------
    ' Adjust to the Width of the longest line
    ' ----------------------------------------
    SetHorizontalExtent
        
    ' ----------------------------------------
    ' Remove recorset from memory
    ' ----------------------------------------
    Set rsData = Nothing
End Sub

Private Sub Reset(Optional Off As Boolean = True)
    ' ----------------------------------------
    ' This will reset the controls
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' To make better sense of "Off" being
    ' True or False, it must be negated to
    ' use it for the BOOLEAN equation for
    ' properties of the controls
    '
    ' We are asking if Off is True, but it
    ' really means make the Enabled property
    ' FALSE (and vice versa)
    ' ----------------------------------------
    Off = Not Off
    
    cmdEdit.Enabled = Off
    cmdRemove.Enabled = Off
    cmdRunNow.Enabled = Off
    cmdSettings.Enabled = Off
End Sub

Private Sub lstEvents_Click()
    ' ----------------------------------------
    ' Re-activate controls if not the first Item
    ' in the list box
    ' ----------------------------------------
    If lstEvents.ListIndex > 0 Then
        ' Activate Controls
        Reset False
    Else
        ' Deactivate Controls
        Reset
    End If
End Sub

Private Function GetHorizontalExtent() As Long
    ' ----------------------------------------
    ' Return the Horizontal Extent
    ' The width of the longest line in Pixels
    ' ----------------------------------------
    
    Dim i As Integer
    Dim lngNum As Long
    Dim lngNumUse As Long
    
    ' ----------------------------------------
    ' Loop through EACH row, finding
    ' the total number of characters
    ' ----------------------------------------
    For i = 0 To GetRowCount - 1
        lngNum = SendMessage(lstEvents.hwnd, LB_GETTEXTLEN, i, 0)
        
        ' ----------------------------------------
        ' If the new number is larger, then save it
        ' ----------------------------------------
        If lngNum > lngNumUse Then lngNumUse = lngNum
    Next i
    
    GetHorizontalExtent = lngNumUse * 10 ' Converts to a base of 8 pixels wide + Padding
End Function

Private Sub SetHorizontalExtent()
    Dim lngHorizontalExtent As Long
    
    ' ----------------------------------------
    ' Get the Horizontal Extent
    ' ----------------------------------------
    lngHorizontalExtent = GetHorizontalExtent
    
    ' ----------------------------------------
    ' Set the TabStops (Columns)
    ' ----------------------------------------
    Call SendMessage(lstEvents.hwnd, LB_SETHORIZONTALEXTENT, lngHorizontalExtent, ByVal 0&)
    
    If lngHorizontalExtent > CPixelX(lstEvents.Width) Then
        ' Show the scrollbar
        ShowScrollBar lstEvents.hwnd, SB_HORZ, True
    Else
        ' Hide the scrollbar
        ShowScrollBar lstEvents.hwnd, SB_HORZ, False
    End If
End Sub

Private Function GetRowCount() As Long
    ' ----------------------------------------
    ' This will return the number of rows in
    ' the listbox
    ' ----------------------------------------
    GetRowCount = SendMessage(lstEvents.hwnd, LB_GETCOUNT, 0, ByVal 0&)
End Function
