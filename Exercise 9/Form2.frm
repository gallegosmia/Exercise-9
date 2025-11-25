VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17820
   LinkTopic       =   "Form2"
   ScaleHeight     =   9975
   ScaleWidth      =   17820
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   8775
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   17295
      _cx             =   30506
      _cy             =   15478
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   13321
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
PrintReport
End Sub

Private Sub cmdGoprint_Click()
PrintReport
End Sub

Private Sub cmdPrint_Click()
PrintReport
End Sub


Private Sub Form_Load()


    On Error GoTo Form_Load_Err

    'MaturityDate = Now
    'DTP_End_Date = Now
    Me.Show
     Me.WindowState = vbMaximized
    CRViewer.Width = Me.Width - 800
    CRViewer.Height = Me.Height - 2000

    Exit Sub

Form_Load_Err:
    ErrReport Err.Description, _
        "LendingClientV2.rep_CollectorList.Form_Load", _
        Erl

    Resume Next
  

End Sub

Sub PrintReport()

        '<EhHeader>
    On Error GoTo PrintReport_Err

        '</EhHeader>

    Dim CRApp As New CRAXDDRT.Application
    Dim Report As CRAXDDRT.Report

    Set Report = CRApp.OpenReport(ReportPath & "customerdetails.rpt")
    Report.Database.Tables(1).Location = ServerPath & "\db\Customer.mdb"
   ' Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
     CRViewer.ReportSource = Report

    'Dim startDate As Date
   ' Dim endDate As Date
    Dim strString As String
    Dim ID As String
    'MousePointer = vbHourglass

    ' Dates from controls or variables
    'startDate = MaturityDate
    'endDate = DTP_End_Date
    ID = (Form1.txtID)
    
    ' Force format to MM/dd/yyyy regardless of system locale
   ' Dim startDateStr As String
    'Dim endDateStr As String
    'startDateStr = Format$(startDate, "MM/dd/yyyy")
   ' endDateStr = Format$(endDate, "MM/dd/yyyy")
   

    ' Use US-style date string in Crystal syntax
  If ID = "" Then
        Exit Sub
    Else
        strString = "{tblCustomer.ID} = " & ID
    End If

Report.RecordSelectionFormula = strString


    Report.RecordSelectionFormula = strString

   ' MaturityDate.Enabled = False
    cmdPrint.Enabled = False
    cmdPrint.Caption = "Working"

    CRViewer.ViewReport

    Do While CRViewer.IsBusy
        DoEvents
    Loop

    CRViewer.Zoom 100
    Me.WindowState = vbMaximized

   ' MaturityDate.Enabled = True
    cmdPrint.Enabled = True
    cmdPrint.Caption = "GO PRINT"
    MousePointer = vbNormal
  
  
    '<EhFooter>
        Exit Sub

PrintReport_Err:
        ErrReport Err.Description, _
            "Please call brayan immediately 0915-891-8530 LendingClientV2.rep_LoansMaturityChecker.PrintReport", _
            Erl

        Resume Next

        '/EhFooter>

End Sub


