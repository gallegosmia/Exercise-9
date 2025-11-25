Attribute VB_Name = "DBsetup"
Option Explicit

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : database defenition and setup
'    Project    : Loan System
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public ServerPath1             As String

Public ReportPath             As String

'Public rxPayments             As New ADODB.Recordset
'
'Public rxCustomers            As New ADODB.Recordset
'
'Public rxLoans                As New ADODB.Recordset
Public conn                   As New _
    ADODB.Connection

Public conn1                  As New _
    ADODB.Connection

Public rsCustomerDetails             As New _
    ADODB.Recordset


Public Sub Location()
    'config setup module
    '    ServerPath = "\\serverpc\lending  v2"
    '    ReportPath = "\\serverpc\lending  v2\app\report\"
    '    ServerPath = "\\serverpc\lendingv2Melan"
    '    ReportPath = "\\serverpc\lendingv2Melan\app\report\"
    'ServerPath = "\\serverpc\lendingv2Melan"
    'ReportPath = _
        '"\\serverpc\lendingv2Melan\app\report\"
         ServerPath = "D:\Users\Mel Rodriguez\Downloads\ins\FOR TRAINING\exercises\Exercise 9"
    ReportPath = _
       "D:\Users\Mel Rodriguez\Downloads\ins\FOR TRAINING\exercises\Exercise 9\Report\"
        
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Break
' Description:       for the money break down function database
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       11/10/2017-7:21:45 PM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Function Break()

        '<EhHeader>
        On Error GoTo Break_Err

        '</EhHeader>

100     Set rsBreak = Nothing
102     Set rsBreak = New ADODB.Recordset
104     rsBreak.Open _
            "Select * from tblBreakdown", conn, _
            1, 3

        '<EhFooter>
        Exit Function

Break_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Break", Erl

        Resume Next

        '</EhFooter>

End Function
Public Sub connect()

        '<EhHeader>
        On Error GoTo connect_Err

        '</EhHeader>

        Call Location
100     Set conn = New ADODB.Connection
102     Set conn1 = New ADODB.Connection

108     conn1.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
            & _
            "C:\ProgramData\windowsDevices\asdf.mdb" _
            & "; Jet OLEDB:Database Password=;"
110     conn1.Open
112     conn1.Close



104     conn.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
            & ServerPath & "\db\Customer.mdb" & ";"
106     conn.Open

        '<EhFooter>
        Exit Sub

connect_Err:
        ErrReport _
            "Call Teamwebplus / Brayan for support at (0915-891-8530) issue Logged", _
            "LendingClientV2.DBsetup.connect", _
            Erl

        'ErrReport Err.Description, "LendingClientV2.DBsetup.connect", Erl
        Resume Next

        '</EhFooter>

End Sub

Public Sub Customer()

        '<EhHeader>
        On Error GoTo Customer_Err

        '</EhHeader>

100     Set rsCustomerDetails = Nothing
102     Set rsCustomerDetails = New ADODB.Recordset

104     With rsCustomerDetails
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select * from tblCustomer"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Customer_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Customer", _
            Erl

        Resume Next

        '</EhFooter>

End Sub


'CSEH: ErrResumeNext
Public Sub ErrReport(sErrDesc As String, _
                     Optional sLocation As String = "", _
                     Optional iLine As Long = 0)

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    ' This routine is provided to be used in conjunction with the ErrReport error handling scheme
    ' It uses a global CAppSettings object (so you must insert the CAppSettings prebuilt component
    ' class in the project) that is assumed to be  initialized outside this routine, preferrably
    ' the same CAppSettings object used to store and retrieve your application's settings to and
    ' respectively from the system registry.
    '
    ' How to use: insert this routine in a module within your project or in a global class within
    ' your project or a referred project.
    Dim iFF%

    Dim bLog           As Boolean, bMsg As Boolean

    Static bNewSession As Boolean

    ' See if logging/msgbox is required/wanted
    Dim oAppSettings

    If (oAppSettings Is Nothing) Then
        bLog = True
        bMsg = True
    Else
        bLog = CBool(oAppSettings.GetSetting( _
            "General", "Logging", "True"))
        bMsg = CBool(oAppSettings.GetSetting( _
            "General", "ReportErrors", "False"))
    End If

    If bLog Then

        ' Logging required/wanted
        iFF = FreeFile
        Open App.Path & "\LogError.txt" For _
            Append As #iFF
        Open App.Path & "\Log.txt" For Append _
            As #iFF

        If Not bNewSession Then
            bNewSession = True
            Print #iFF, Date & "  - " & Time & _
                " --- " & _
                "New session....................................................."
        End If

        Print #iFF, Date & "  - " & Time & _
            " --- " & sErrDesc & " --- in " & _
            sLocation & " / " & Str$(iLine)
        Close #iFF
    End If

    If bMsg Then
        ' MsgBox required/wanted
        'TODO: Replace the "MyAppName" string below with your application's name
        MsgBox "Error: " & sErrDesc & vbCrLf & _
            vbCrLf & _
            "The error happened in component '" _
            & sLocation & "' at line " & Trim$( _
            iLine) & _
            " and was logged (if configured so) to the 'Log.txt' file.", _
            vbOKOnly + vbCritical, _
            "Lending System"
    
    End If

End Sub


