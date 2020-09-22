Attribute VB_Name = "modMain"
Option Explicit

' ----------------------------------------
' These are the Global variabes used
' ----------------------------------------
Public dbDatabase As ADODB.Connection

Public Sub Main()
    ' ----------------------------------------
    ' This is where ALL preliminary loading
    ' of information is handled.
    '
    ' All initial setting up is completed here
    ' before any forms are loaded.
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' Load and Show the Splash screen
    ' ----------------------------------------
    Load frmSplash
    frmSplash.Show
    
    ' ----------------------------------------
    ' Load the main form
    ' ----------------------------------------
    Load frmMenu
    
    ' ----------------------------------------
    ' Open the main form
    ' ----------------------------------------
    'frmMenu.Show
End Sub

Public Function ParseType(ModeType As String) As String
    ' ----------------------------------------
    ' This will separate the info for the Mode Type
    '
    ' If there is an "\", then anything BEFORE
    ' the "\" is parsed as the TYPE
    ' Everything after and including the "\" is
    ' disreguarded
    ' ----------------------------------------
    Dim i As Integer
    i = InStr(ModeType, "\")
    
    ' ----------------------------------------
    ' Return value
    ' ----------------------------------------
    If i = 0 Then
        ParseType = Trim(ModeType)
    Else
        ParseType = Trim(Mid(Trim(ModeType), 1, i - 1))
    End If
End Function

Public Function ParseEvent(EventID As String) As Long
    ' ----------------------------------------
    ' This will separate the info for the Event ID
    '
    ' If there is an "\", then anything AFTER
    ' the "\" is parsed as the TYPE
    ' Everything before and including the "\" is
    ' disreguarded
    ' ----------------------------------------
    Dim i As Integer
    i = InStr(EventID, "\")
    
    ' ----------------------------------------
    ' Return value
    ' ----------------------------------------
    ParseEvent = CLng(Trim(Mid(Trim(EventID), i + 1, Len(Trim(EventID)))))
End Function

Public Function ChangeCharacter(Word As String, Optional Character As String = "", Optional Replace As Boolean = True, Optional ReplaceCharacter As String = "?") As String
    ' ----------------------------------------
    ' This will find the character specified,
    ' either change the character to a specified
    ' character, or remove the character altogether
    ' ----------------------------------------
    
    Dim lngApostrophePos As Long
    Dim strWord As String
    
    strWord = Trim(Word)
    
    ' ----------------------------------------
    ' Look for the character to be
    ' replaced or removed
    ' ----------------------------------------
    Do
        ' Get position of character
        lngApostrophePos = InStr(Trim(strWord), Character)
        
        If lngApostrophePos <> 0 Then
            ' ----------------------------------------
            ' Found character... Replace?
            ' ----------------------------------------
            If lngApostrophePos = 1 Then
                If Replace = True Then
                    ' Replace
                    strWord = ReplaceCharacter + Mid(Trim(strWord), 2, Len(Trim(strWord)))
                Else
                    ' Remove
                    strWord = Mid(Trim(strWord), 2, Len(Trim(strWord)))
                End If
            Else
                If Replace = True Then
                    ' Replace
                    strWord = Mid(Trim(strWord), 1, lngApostrophePos - 1) _
                                + ReplaceCharacter _
                                + Mid(Trim(strWord), lngApostrophePos + 1, Len(Trim(strWord)))
                Else
                    ' Remove
                    strWord = Mid(Trim(strWord), 1, lngApostrophePos - 1) _
                                + Mid(Trim(strWord), lngApostrophePos + 1, Len(Trim(strWord)))
                End If
            End If
        End If
    Loop Until lngApostrophePos = 0
    
    ' ----------------------------------------
    ' Return the value
    ' ----------------------------------------
    ChangeCharacter = strWord
End Function

Public Function GetEventID(Name As String) As Long
    ' ----------------------------------------
    ' This will return the Event ID
    ' ----------------------------------------
    Dim rsData As ADODB.Recordset
    
    Dim strEvent As String
    Dim lngEvent As Long
    Dim strSQL As String
    
    ' ----------------------------------------
    ' Make sure no Apostrophes are in the NAME
    ' ----------------------------------------
    strEvent = ChangeCharacter(Name, "'", True, "?")
    
    ' ----------------------------------------
    ' Connect to the table
    ' ----------------------------------------
    Set rsData = New ADODB.Recordset
    
    strSQL = "SELECT [EVENT_ID] FROM [tblEVENT] WHERE [EVENT_NAME] LIKE '" & Trim(strEvent) & "'"
    
    With rsData
        .Open strSQL, dbDatabase
        
        ' See if any records are there
        If Not (.BOF And .EOF) Then
            ' Grab the EVENT_ID
            lngEvent = ![EVENT_ID]
        Else
            ' Error - No records found
            lngEvent = -1
        End If
        
        .Close
    End With
    
    Set rsData = Nothing
    
    ' ----------------------------------------
    ' Return the value
    ' [-1 is an error for no record found]
    ' ----------------------------------------
    GetEventID = lngEvent
End Function

Public Function GetEventName(EventName As String) As String
    ' ----------------------------------------
    ' This will get the real event name
    ' from the string
    ' ----------------------------------------
    Dim i As Integer
    
    ' Find the tab (vbTab)
    i = InStr(Trim(EventName), vbTab)
    
    ' Return value
    GetEventName = Trim(Mid(Trim(EventName), 1, i - 1))
End Function

Public Sub RunEvent(EventID As Long, Optional RunSilent As Boolean = True, Optional Settings As Boolean = False)
    ' ----------------------------------------
    ' This will run the specified event in either
    ' SETTINGS mode or Normal mode
    ' ----------------------------------------
    Dim rsData As Recordset
    Dim strSQL As String
    Dim strCommands As String
    Dim strEvent As String
    Dim dblDouble As Double
    
    ' ----------------------------------------
    ' Establish recordset
    ' ----------------------------------------
    Set rsData = New ADODB.Recordset
    
    ' ----------------------------------------
    ' Build SQL
    ' ----------------------------------------
    strSQL = "SELECT [EVENT_PROGRAM], [EVENT_COMMANDS], [EVENT_SETTING_COMMANDS] FROM [tblEVENT] WHERE [EVENT_ID]=" & Trim(Str(EventID))
    
    ' ----------------------------------------
    ' Test connectivity and open recordset
    ' ----------------------------------------
    With rsData
        .Open strSQL, dbDatabase
        
        If Not (.BOF And .EOF) Then
            ' ----------------------------------------
            ' Grab the Program and Path
            ' ----------------------------------------
            strEvent = Trim(!EVENT_PROGRAM)
            
            ' ----------------------------------------
            ' Check to see if in SETTINGS Mode
            ' ----------------------------------------
            If Settings = False Then
                ' Normal Mode
                strCommands = Trim(!EVENT_COMMANDS)
            Else
                ' Settings Mode
                strCommands = Trim(!EVENT_SETTING_COMMANDS)
            End If
        End If
        
        .Close
    End With
    
    ' ----------------------------------------
    ' Terminate Recordset
    ' ----------------------------------------
    Set rsData = Nothing
    
    ' ----------------------------------------
    ' If in SETTINGS mode, see if user
    ' wants to run without commands
    ' ----------------------------------------
    ' ----------------------------------------
    ' Check to see in in SILENT mode
    ' ----------------------------------------
    If Settings = True Then
        If strCommands = "" Then
            If RunSilent = False Then
                If MsgBox("The event has no commands described for running in setup mode." & Chr$(13) _
                & "Do you want to continue running the event?", vbQuestion + vbYesNo, _
                App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision) = vbNo Then Exit Sub
            End If
        End If
    End If
    
    ' ----------------------------------------
    ' If there are command described,
    ' ask user how to run the Event
    ' ----------------------------------------
    If strCommands <> "" Then
        If RunSilent = True Then
            If MsgBox("The event has commands specified." & Chr$(13) _
                & "Do you want to continue running the event with the commands?", vbQuestion + vbYesNo, _
                App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision) = vbNo Then
                
                ' Continue without commands
            Else
                ' Continue
                strEvent = strEvent & " " & strCommands
            End If
        Else
            ' Continue
            strEvent = strEvent & " " & strCommands
        End If
    End If
    
    ' ----------------------------------------
    ' Run the Event
    ' ----------------------------------------
    If RunSilent = True Then
        dblDouble = Shell(strEvent, vbMinimizedNoFocus)
    Else
        dblDouble = Shell(strEvent, vbNormalFocus)
    End If
End Sub

Public Sub RunScheduledEvents()
    ' ----------------------------------------
    ' This will run the scheduled events
    ' ----------------------------------------
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim strRemoveSQL As String
    Dim bytHour As Byte
    Dim bytDayOfWeek As Byte
    Dim bytMonth As Byte
    Dim bytDayOfMonth As Byte
    
    ' ----------------------------------------
    ' Keep windows breathing
    ' ----------------------------------------
    DoEvents
    
    strRemoveSQL = ""

    ' ----------------------------------------
    ' Establish Recordset
    ' ----------------------------------------
    Set dbDatabase = New ADODB.Connection
    Set rsData = New ADODB.Recordset
    
    ' ----------------------------------------
    ' Connect to Database
    ' ----------------------------------------
    With dbDatabase
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\EventScheduler.mdb"
        .Open
    End With

    ' ----------------------------------------
    ' Get the current hour (24-Hour based)
    ' ----------------------------------------
    bytHour = GetHourOfDay
    
    ' ----------------------------------------
    ' Get Day of Week
    ' ----------------------------------------
    bytDayOfWeek = GetDayOfWeekNumber
    
    ' ----------------------------------------
    ' Get Month Number
    ' ----------------------------------------
    bytMonth = GetMonthNumber
    
    ' ----------------------------------------
    ' Get the Day of the Month
    ' ----------------------------------------
    bytDayOfMonth = GetDayOfMonth
    
    ' ----------------------------------------
    ' Build the SQL to find all Events to run
    ' ----------------------------------------
    strSQL = "SELECT * FROM [tblEVENT] WHERE [EVENT_HOUR_OF_DAY]=" & Trim(Str(bytHour))
    
    ' ----------------------------------------
    ' Create Recordset
    ' ----------------------------------------
    With rsData
        .Open strSQL, dbDatabase, adOpenDynamic, adLockOptimistic
        
        ' ----------------------------------------
        ' Records Found
        '
        ' Check to make sure event is to run
        ' ----------------------------------------
        Do Until .EOF
            ' ----------------------------------------
            ' Keep Windows Breathing
            ' ----------------------------------------
            DoEvents
            
            ' ----------------------------------------
            ' See which Frequency ID
            ' ----------------------------------------
            Select Case ![EVENT_FREQUENCY_ID]
                Case 0
                    ' Daily
                    ' Run the Event
                    RunEvent CLng(!EVENT_ID)
                Case 1
                    ' Daily (Non-Weekend)
                    If bytDayOfWeek <> 6 And bytDayOfWeek <> 7 Then
                        ' Run the Event
                        RunEvent CLng(!EVENT_ID)
                    End If
                Case 2
                    ' Weekly
                    If !EVENT_DAY_OF_WEEK = bytDayOfWeek Then
                        ' Run the Event
                        RunEvent CLng(!EVENT_ID)
                    End If
                Case 3
                    ' Monthly
                    If !EVENT_DAY_OF_MONTH = bytDayOfMonth Then
                        ' Run Event
                        RunEvent CLng(!EVENT_ID)
                    End If
                Case 4
                    ' Once Only
                    If !EVENT_MONTH = bytMonth And !EVENT_DAY_OF_MONTH = bytDayOfMonth Then
                        ' Run Event
                        RunEvent CLng(!EVENT_ID)
                        
                        If strRemoveSQL = "" Then
                            strRemoveSQL = "[EVENT_ID]=" & Trim(Str(![EVENT_ID]))
                        Else
                            strRemoveSQL = strRemoveSQL & " OR [EVENT_ID]=" & Trim(Str(![EVENT_ID]))
                        End If
                    End If
            End Select
            
            ' Next Record
            .MoveNext
        Loop
        
        .Close
    End With
    
    ' ----------------------------------------
    ' See if Any Events need to be removed
    ' [Once Only Events]
    ' ----------------------------------------
    If strRemoveSQL <> "" Then
        strSQL = "DELETE * FROM [tblEVENT] WHERE " & Trim(strRemoveSQL)
        
        dbDatabase.Execute (strSQL)
    End If
    
    ' ----------------------------------------
    ' Eliminate rsData
    ' ----------------------------------------
    Set rsData = Nothing

    ' ----------------------------------------
    ' Destroy the Database object
    ' ----------------------------------------
    Set dbDatabase = Nothing
End Sub

Public Function GetDayOfWeekNumber() As Byte
    ' ----------------------------------------
    ' This will return the Day of the Week
    ' ----------------------------------------
    Dim bytNum As Byte
    
    Select Case LCase(Trim(Format(Now, "DDDD")))
        Case "monday"
            bytNum = 1
        Case "tuesday"
            bytNum = 2
        Case "wednesday"
            bytNum = 3
        Case "thursday"
            bytNum = 4
        Case "friday"
            bytNum = 5
        Case "saturday"
            bytNum = 6
        Case "sunday"
            bytNum = 7
    End Select
    
    ' ----------------------------------------
    ' Return the value
    ' ----------------------------------------
    GetDayOfWeekNumber = bytNum
End Function

Public Function GetMonthNumber() As Byte
    ' ----------------------------------------
    ' This will return the Month Number
    ' ----------------------------------------
    GetMonthNumber = CByte(Trim(Format(Now, "M")))
End Function

Public Function GetHourOfDay() As Byte
    ' ----------------------------------------
    ' This will get the Hour of the Day
    ' ----------------------------------------
    GetHourOfDay = CByte(Trim(Hour(Now)))
End Function

Public Function GetDayOfMonth() As Byte
    ' ----------------------------------------
    ' This will return the Day of the Month
    ' ----------------------------------------
    GetDayOfMonth = CByte(Trim(Day(Now)))
End Function

Public Sub Terminate()
    Dim lngFrmCount As Long
    
    ' --------------------------------
    ' Unload all forms
    ' --------------------------------
    For lngFrmCount = (Forms.Count - 1) To 0 Step -1
        Unload Forms(lngFrmCount)
    Next lngFrmCount
    
    ' ----------------------------------------
    ' Make sure dbDatabase is nothing
    ' ----------------------------------------
    Set dbDatabase = Nothing
End Sub

