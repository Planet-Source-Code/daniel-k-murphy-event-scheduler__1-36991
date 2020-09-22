VERSION 5.00
Begin VB.Form frmAddEdit 
   Caption         =   "Add/Edit Event"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame fmeEvent 
      Height          =   5655
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdSelectFile 
         Caption         =   "Select File"
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdSelectDate 
         Caption         =   "Select Date"
         Height          =   265
         Left            =   3600
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtNotes 
         Height          =   2295
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3240
         Width           =   4695
      End
      Begin VB.TextBox txtSettings 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2880
         Width           =   4695
      End
      Begin VB.TextBox txtCommands 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   2520
         Width           =   4695
      End
      Begin VB.ComboBox cboHour 
         Height          =   315
         ItemData        =   "frmAddEdit.frx":0000
         Left            =   1440
         List            =   "frmAddEdit.frx":005A
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cboDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboFrequency 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtProgramName 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         Caption         =   "File Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label lblSettings 
         AutoSize        =   -1  'True
         Caption         =   "Settings:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2985
         Width           =   615
      End
      Begin VB.Label lblCommands 
         AutoSize        =   -1  'True
         Caption         =   "Commands:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2600
         Width           =   825
      End
      Begin VB.Label lblHour 
         AutoSize        =   -1  'True
         Caption         =   "Hour:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   390
      End
      Begin VB.Label lblDateDay 
         AutoSize        =   -1  'True
         Caption         =   "Date/Day"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label lblFrequency 
         AutoSize        =   -1  'True
         Caption         =   "Frequency:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblProgramName 
         AutoSize        =   -1  'True
         Caption         =   "Event Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------
' Variables used for Form
' ----------------------------------------
Dim strTag As String                ' Used to identify how the form is being used
Dim blnFormActivated As Boolean     ' Used to identify if the form is activated
Dim typData As DataType             ' Data Type for the storing the information
Dim lngEvent As Long                ' Used to hold the Event ID
Dim blnIsDirty As Boolean           ' Used to see if the data has changed

' ----------------------------------------
' Data Type for Data loaded
' ----------------------------------------
Private Type DataType
    Name As String                  ' Event Name
    Program As String               ' Event Filename and Path
    Frequency As Byte               ' Frequency at which to run the event
    DateDay As String               ' Date/Day of event
    Hour As Byte                    ' Hour to launch event
    Commands As String              ' Any commands needed to run event
    Settings As String              ' Commands for the setup for the event
    Notes As String                 ' Notes about the event
End Type

Private Sub cboDate_Click()
    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.DateDay = Str(cboDate.ListIndex)
    
    blnIsDirty = True
End Sub

Private Sub cboFrequency_Click()
    ' ----------------------------------------
    ' This will update the date Combo, Text Box, and GO button
    ' ----------------------------------------
    PopulateDate

    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.Frequency = CByte(cboFrequency.ListIndex)
    Select Case typData.Frequency
        Case 2
            ' Weekly
            typData.DateDay = Str(cboDate.ListIndex)
        Case 3
            ' Monthly
            typData.DateDay = Str(cboDate.ListIndex)
    End Select
        
    blnIsDirty = True
End Sub

Private Sub cboHour_Click()
    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.Hour = CByte(cboHour.ListIndex)
    
    blnIsDirty = True
End Sub

Private Sub cmdCancel_Click()
    ' ----------------------------------------
    ' See if there were any changes made
    ' and ask to save if so.
    '
    ' Otherwise, close the form
    ' ----------------------------------------
    Dim intAnswer As Integer
    
    If blnIsDirty = True Then
        intAnswer = MsgBox("Do you want to save changes?", vbYesNoCancel, App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision)
        
        Select Case intAnswer
            Case vbYes
                ' Save the changes
                SaveData
            Case vbNo
                ' Do nothing
            Case vbCancel
                ' Skip out from closing the form
                Exit Sub
            Case Else
                ' Do nothing
        End Select
    End If
    
    Unload Me
End Sub

Private Sub cmdSelectDate_Click()
    ' ----------------------------------------
    ' Grab the date from frmDate
    ' ----------------------------------------
    Dim strDate As String
    
    strDate = Format(frmDate.GetDate, "M/D")

    If strDate <> "" Then
        ' ----------------------------------------
        ' A date was selected
        ' ----------------------------------------
        txtDate.Text = strDate
        
        ' ----------------------------------------
        ' This will save the data into the Data Type
        '
        ' Will also set blnIsDirty to True
        ' ----------------------------------------
        typData.DateDay = Trim(txtDate.Text)
        
        blnIsDirty = True
    End If
End Sub

Private Sub cmdOK_Click()
    ' ----------------------------------------
    ' Check the for required data
    ' ----------------------------------------
    If CheckData = False Then Exit Sub
    
    ' ----------------------------------------
    ' This will save the data and Close the form
    ' ----------------------------------------
    SaveData
    
    ' ----------------------------------------
    ' Release the form
    ' ----------------------------------------
    Unload Me
End Sub

Private Sub cmdSelectFile_Click()
    ' ----------------------------------------
    ' Grab the path and file name from frmOpen
    ' ----------------------------------------
    Dim strFilename As String
    
    strFilename = frmOpen.GetFile
    
    If strFilename <> "" Then
        ' ----------------------------------------
        ' A file was selected
        ' ----------------------------------------
        txtFilename.Text = strFilename
        
        ' ----------------------------------------
        ' This will save the data into the Data Type
        '
        ' Will also set blnIsDirty to True
        ' ----------------------------------------
        typData.Program = UCase(Trim(txtFilename.Text))
        
        blnIsDirty = True
    End If
End Sub

Private Sub Form_Activate()
    ' ----------------------------------------
    ' Depending upon Adding or Editing,
    ' Either populate by creating new or
    ' Filling in with info that is Event specific
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' Using the strTAG to identify
    ' the way the form is loading,
    ' ADD = Adding a new Event
    ' EDIT\ID = Editing an existing event
    ' ----------------------------------------
    
    ' Check to see if form is already activated
    If blnFormActivated = True Then Exit Sub
    
    ' Data variables
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    
    ' EVENT variables
    Dim strType As String
    
    ' ----------------------------------------
    ' Clear the data type
    ' ----------------------------------------
    ClearDataType
    
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
    ' Determine which mode the user is requesting
    ' ----------------------------------------
    strType = ParseType(Trim(strTag))
    
    Select Case LCase(strType)
        Case "add"
            ' ADD Mode.  Create new event ID
            strSQL = "SELECT [EVENT_ID] FROM [tblEVENT] ORDER BY [EVENT_ID] DESC"
            
            ' Establish recordset
            With rsData
                ' Connect recordset
                .Open strSQL, dbDatabase
                
                ' Check for records
                If Not (.BOF And .EOF) Then
                    ' Grab the event ID and add one
                    lngEvent = CLng(!EVENT_ID) + 1
                Else
                    lngEvent = 1
                End If
                
                ' Close the recordset
                .Close
            End With
            
            ' Clear the form
            Reset
        Case "edit"
            ' EDIT Mode.  Get event ID
            lngEvent = ParseEvent(Trim(strTag))
            
            ' Clear the form
            Reset
            
            ' Fill the form
            Populate lngEvent
    End Select
    
    ' Show form as activated
    blnFormActivated = True
    
    ' Set blnIsDirty = False
    blnIsDirty = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' ----------------------------------------
    ' Remove all references to variables
    ' and clean up anything that was used
    ' ----------------------------------------
    dbDatabase.Close
    Set dbDatabase = Nothing
End Sub

Public Function ShowForm(FormType As String)
    ' ----------------------------------------
    ' This will set the strTag to the value passed
    '
    ' Will also show the form in a Modal (App only) format
    ' ----------------------------------------
    blnFormActivated = False
    strTag = FormType
    Me.Show vbModal
End Function

Private Sub Reset()
    ' ----------------------------------------
    ' This will reset the form to be blank,
    ' only containing the defaults.
    ' ----------------------------------------
    With Me
        .txtProgramName.Text = ""
        .txtFilename.Text = ""
        .txtFilename.Enabled = False
        .txtCommands.Text = ""
        .txtSettings.Text = ""
        .txtNotes.Text = ""
    End With
    
    PopulateFrequency
    PopulateHour
    PopulateDate
End Sub

Private Sub Populate(EventID As Long)
    ' ----------------------------------------
    ' This will load the data from the database
    ' and populate the form.
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' Load the data
    ' ----------------------------------------
    LoadData EventID
    
    ' ----------------------------------------
    ' Populate the form
    ' ----------------------------------------
    With typData
        txtProgramName.Text = .Name
        txtFilename.Text = .Program
        txtNotes.Text = .Notes
        txtSettings.Text = .Settings
        txtCommands.Text = .Commands
        If .Frequency < 5 Then cboFrequency.ListIndex = .Frequency
        
        Select Case .Frequency
            Case 0
                ' Daily
                ' Fill the Date
                PopulateDate
            Case 1
                ' Daily (Non-Weekend)
                ' Fill the Date
                PopulateDate
            Case 2
                ' Weekly
                ' Fill the Date
                PopulateDate
                
                ' Select the date
                cboDate.ListIndex = CInt(.DateDay)
            Case 3
                ' Monthly
                ' Fill the Date
                PopulateDate
                
                ' Select the date
                cboDate.ListIndex = CInt(.DateDay)
            Case 4
                ' Once Only
                ' Fill the Date
                PopulateDate
                
                ' Select the date
                txtDate.Text = .DateDay
            Case Else
                ' Undefined
                cboFrequency.ListIndex = 3
                
                ' Fill the Date
                PopulateDate
                
                ' Select the date
                txtDate.Text = .DateDay
        End Select
        
        cboHour.ListIndex = CInt(.Hour)
    End With
End Sub

Private Sub PopulateDate()
    ' ----------------------------------------
    ' This will, based upon the Frequency,
    ' populate the date combo list
    ' ----------------------------------------
    Dim i As Byte
    
    ' Clear the combo and text boxes
    cboDate.Clear
    txtDate.Text = ""
    txtDate.Enabled = False
    
    Select Case cboFrequency.ListIndex
        Case 0
            ' Daily
            cboDate.Enabled = False
            cmdSelectDate.Enabled = False
        Case 1
            ' Daily (Non-Weekend)
            cboDate.Enabled = False
            cmdSelectDate.Enabled = False
        Case 2
            ' Weekly
            cboDate.Enabled = True
            cmdSelectDate.Enabled = False
            
            With cboDate
                .AddItem "MONDAY"
                .AddItem "TUESDAY"
                .AddItem "WEDNESDAY"
                .AddItem "THURSDAY"
                .AddItem "FRIDAY"
                .AddItem "SATURDAY"
                .AddItem "SUNDAY"
                .ListIndex = 0
                .Refresh
            End With
        Case 3
            ' Monthly
            cboDate.Enabled = True
            cmdSelectDate.Enabled = False
            
            With cboDate
                For i = 1 To 28
                    .AddItem Trim(Str(i))
                Next i
                
                .ListIndex = 0
                .Refresh
            End With
        Case 4
            ' Once Only
            cboDate.Enabled = False
            cmdSelectDate.Enabled = True
            
        Case Else
            ' Undefined
            cboDate.Enabled = False
            cmdSelectDate.Enabled = True
    End Select

End Sub

Private Sub PopulateFrequency()
    ' ----------------------------------------
    ' This will populate the Frequency combo list
    ' ----------------------------------------
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    
    ' ----------------------------------------
    ' Establish recordset
    ' ----------------------------------------
    Set rsData = New ADODB.Recordset
    
    strSQL = "SELECT * FROM [tblEVENT_FREQUENCY_TYPE_LOOKUP] ORDER BY [EVENT_FREQUENCY_ID]"
    
    rsData.Open strSQL, dbDatabase
    
    ' ----------------------------------------
    ' Populate the data into the list
    ' ----------------------------------------
    With cboFrequency
        .Clear
        Do Until rsData.EOF
            .AddItem UCase(Trim(rsData![EVENT_FREQUENCY_DESCRIPTION])), rsData![EVENT_FREQUENCY_ID]
            
            ' Next record
            rsData.MoveNext
        Loop
    End With
    
    ' ----------------------------------------
    ' Close recordset
    ' ----------------------------------------
    rsData.Close
    Set rsData = Nothing
    
    ' ----------------------------------------
    ' Set the default settings
    ' ----------------------------------------
    cboFrequency.ListIndex = 0 ' Trim(UCase("Daily"))
    cboFrequency.Refresh
End Sub

Private Sub PopulateHour()
    ' ----------------------------------------
    ' This will populate the Hour combo list
    ' ----------------------------------------
    Dim i As Byte
    
    With cboHour
        .Clear
        For i = 0 To 23
            .AddItem Trim(Str(i))
        Next i
        .ListIndex = 0
    End With
End Sub

Private Sub ClearDataType()
    ' ----------------------------------------
    ' This will clear the data from the Type
    '
    ' This will ensure that NO data is placed incorrectly
    ' onto a different event
    ' ----------------------------------------
    With typData
        .Commands = ""
        .DateDay = ""
        .Frequency = 0
        .Hour = 0
        .Name = ""
        .Notes = ""
        .Program = ""
        .Settings = ""
    End With
End Sub

Private Sub LoadData(EventID As Long)
    ' ----------------------------------------
    ' This will load the data for the form
    ' ----------------------------------------
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    
    Set rsData = New ADODB.Recordset
    
    strSQL = "SELECT * FROM [tblEvent] WHERE [EVENT_ID]=" & Trim(Str(EventID))
    
    rsData.Open strSQL, dbDatabase
    
    With rsData
        If Not (.BOF And .EOF) Then
            ' ----------------------------------------
            ' Get the data from the database
            ' and populate the data type
            ' ----------------------------------------
            typData.Name = UCase(![EVENT_NAME])
            typData.Program = UCase(![EVENT_PROGRAM])
            typData.Notes = ![EVENT_NOTES]
            typData.Commands = ![EVENT_COMMANDS]
            typData.Settings = ![EVENT_SETTING_COMMANDS]
            typData.Frequency = ![EVENT_FREQUENCY_ID]
            typData.Hour = ![EVENT_HOUR_OF_DAY]
            
            ' ----------------------------------------
            ' Capture the Date/Day according to Frequency
            '
            ' This part has redundancy (IE: 2 CASE statements
            ' instead of 5), and could have
            ' been lumped together, however, I purposely
            ' put it this way for clarity, and ease to follow.
            ' ----------------------------------------
            Select Case typData.Frequency
                Case 0
                    ' Daily
                    ' Do Nothing
                Case 1
                    ' Daily (Non-Weekend)
                    ' Do Nothing
                Case 2
                    ' Weekly
                    ' Capture the DAY
                    typData.DateDay = Trim(Str(![EVENT_DAY_OF_WEEK]))
                Case 3
                    ' Monthly
                    ' Capture the Month and Day
                    typData.DateDay = Trim(Str(![EVENT_DAY_OF_MONTH]))
                Case 4
                    ' Once Only
                    ' Capture the Month and Day
                    typData.DateDay = Trim(Str(![EVENT_MONTH])) & "/" & Trim(Str(![EVENT_DAY_OF_MONTH]))
                Case Else
                    ' Undefined
                    ' Capture the Month and Day
                    typData.DateDay = Trim(Str(![EVENT_MONTH])) & "/" & Trim(Str(![EVENT_DAY_OF_MONTH]))
            End Select
        End If
        
        .Close
    End With
    
    Set rsData = Nothing
End Sub

Private Sub SaveData()
    ' ----------------------------------------
    ' This will save all the data
    ' ----------------------------------------
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim i As Byte
    
    Set rsData = New ADODB.Recordset
    
    If strTag = "ADD" Then
        strSQL = "[tblEvent]"
    Else
        strSQL = "SELECT * FROM [tblEvent] WHERE [EVENT_ID]=" & Trim(Str(lngEvent))
    End If
    
    rsData.Open strSQL, dbDatabase, adOpenDynamic, adLockOptimistic
    
    With rsData
        If strTag = "ADD" Then
            .AddNew
        End If
            If Not (.BOF And .EOF) Then
                ' ----------------------------------------
                ' Using the data type and a
                ' couple of other variables,
                ' save the data to the database
                ' ----------------------------------------
                ![EVENT_ID] = lngEvent
                ![EVENT_NAME] = UCase(typData.Name)
                ![EVENT_PROGRAM] = UCase(typData.Program)
                ![EVENT_NOTES] = typData.Notes
                ![EVENT_COMMANDS] = typData.Commands
                ![EVENT_SETTING_COMMANDS] = typData.Settings
                ![EVENT_FREQUENCY_ID] = CByte(typData.Frequency)
                ![EVENT_HOUR_OF_DAY] = CByte(typData.Hour)
                
                ' ----------------------------------------
                ' Capture the Date/Day according to Frequency
                '
                ' This part has redundancy (IE: 2 CASE statements
                ' instead of 5), and could have
                ' been lumped together, however, I purposely
                ' put it this way for clarity, and ease to follow.
                ' ----------------------------------------
                Select Case typData.Frequency
                    Case 0
                        ' Daily
                        ' Do Nothing
                    Case 1
                        ' Daily
                        ' Do Nothing
                    Case 2
                        ' Weekly
                        ' Capture the DAY
                        ![EVENT_DAY_OF_WEEK] = CByte(typData.DateDay)
                    Case 3
                        ' Monthly
                        ' Capture the Month and Day
                        i = InStr(Trim(typData.DateDay), "/")
                        
                        ![EVENT_DAY_OF_MONTH] = CByte(Mid(Trim(typData.DateDay), i + 1, Len(Trim(typData.DateDay))))
                    Case 4
                        ' Once Only
                        ' Capture the Month and Day
                        i = InStr(Trim(typData.DateDay), "/")
                        
                        ![EVENT_MONTH] = CByte(Mid(Trim(typData.DateDay), 1, i - 1))
                        ![EVENT_DAY_OF_MONTH] = CByte(Mid(Trim(typData.DateDay), i + 1, Len(Trim(typData.DateDay))))
                    Case Else
                        ' Undefined
                        ' Capture the Month and Day
                        i = InStr(Trim(typData.DateDay), "/")
                        
                        ![EVENT_MONTH] = CByte(Mid(Trim(typData.DateDay), 1, i - 1))
                        ![EVENT_DAY_OF_MONTH] = CByte(Mid(Trim(typData.DateDay), i + 1, Len(Trim(typData.DateDay))))
                End Select
            .Update
        End If
        
        .Close
    End With
    
    Set rsData = Nothing
End Sub

Private Sub txtCommands_KeyUp(KeyCode As Integer, Shift As Integer)
    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.Commands = Trim(txtCommands.Text)
    
    blnIsDirty = True
End Sub

Private Sub txtNotes_KeyUp(KeyCode As Integer, Shift As Integer)
    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.Notes = Trim(txtNotes.Text)
    
    blnIsDirty = True
End Sub

Private Sub txtProgramName_KeyUp(KeyCode As Integer, Shift As Integer)
    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.Name = UCase(Trim(txtProgramName.Text))
    
    blnIsDirty = True
End Sub

Private Sub txtSettings_KeyUp(KeyCode As Integer, Shift As Integer)
    ' ----------------------------------------
    ' This will save the data into the Data Type
    '
    ' Will also set blnIsDirty to True
    ' ----------------------------------------
    typData.Settings = Trim(txtSettings.Text)
    
    blnIsDirty = True
End Sub

Private Function CheckData() As Boolean
    ' ----------------------------------------
    ' This will make sure the required
    ' fields are populated BEFORE
    ' attempting to add the record to
    ' the database
    ' ----------------------------------------
    Dim strMsg As String
    
    strMsg = ""
    
    ' ----------------------------------------
    ' See if user missed something needed
    ' ----------------------------------------
    With typData
        ' ----------------------------------------
        ' Check for Event Name
        ' ----------------------------------------
        If Trim(.Name) = "" Then
            strMsg = "Please enter an Event Name"
            txtProgramName.SetFocus
        End If
        
        ' ----------------------------------------
        ' Check for Filename and Path
        ' ----------------------------------------
        If Trim(.Program) = "" Then
            strMsg = "Please enter the Path and Filename"
            cmdSelectFile.SetFocus
        End If
        
        ' ----------------------------------------
        ' Check for the Frequency
        ' ----------------------------------------
        If .Frequency > 1 Then
            If Trim(.DateDay) = "" Then
                strMsg = "Please enter the Date/Day"
                
                Select Case .Frequency
                    Case 2
                        ' Weekly
                        cboDate.SetFocus
                    Case 3
                        ' Monthly
                        cboDate.SetFocus
                    Case 4
                        ' Once Only
                        cmdSelectDate.SetFocus
                End Select
            End If
        End If
    End With
    
    ' ----------------------------------------
    ' Check to see if an Error Message (Non-VB)
    ' was created, and return value
    ' ----------------------------------------
    If Trim(strMsg) <> "" Then
        MsgBox strMsg, vbOKOnly + vbCritical, App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
        CheckData = False
    Else
        CheckData = True
    End If
End Function
