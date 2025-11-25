VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19110
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   19110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13080
      TabIndex        =   32
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtID 
      Height          =   495
      Left            =   1560
      TabIndex        =   31
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txtsearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   29
      Top             =   6480
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   840
      TabIndex        =   27
      Top             =   7200
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   26
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   25
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   24
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   23
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "USER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4335
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Width           =   16095
      Begin VB.ComboBox cmbGender 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "Customerdetails.frx":0000
         Left            =   2400
         List            =   "Customerdetails.frx":000D
         TabIndex        =   7
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtemailadd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12480
         TabIndex        =   9
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtoccupation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtphone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12000
         TabIndex        =   6
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   3120
         Width           =   6735
      End
      Begin VB.TextBox txtage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtbday 
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108855297
         CurrentDate     =   45974
      End
      Begin VB.TextBox txtFname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtMname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtLname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10320
         TabIndex        =   3
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "EMAIL ADDRESS:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   22
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "GENDER:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   21
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "OCCUPATION:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   20
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "MOBILE NUMBER:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   19
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ADDRESS:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "AGE:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "BIRTHDATE:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "LAST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "MIDDLE NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "FIRST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "FULLNAME:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Label Label13 
      Caption         =   "ID Num"
      Height          =   495
      Left            =   360
      TabIndex        =   30
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   28
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()

        On Error GoTo btnAdd_Click_Err


100     txtsearch.Enabled = False

102     If btnAdd.Caption = "Add" Then

104         If rsCustomerDetails.State = 1 Then rsCustomerDetails.Close
            'Sort records descending
106         rsCustomerDetails.Open "Select * from tblCustomer"
            Call autoNumber
            ClearControls
112         dtbday.Visible = True
114         btnAdd.Caption = "Save"
116         btnClose.Caption = "Cancel"
118         txtLname.Locked = False
120         txtFname.Locked = False
122         txtMname.Locked = False
124         txtaddress.Locked = False
            txtoccupation.Locked = False
132         txtemailadd.Locked = False
136         txtage.Locked = False
140         cmbGender.Locked = False
142         txtphone.Locked = False
148         txtID.Locked = False
160         DataGrid1.Enabled = False
162         txtFname.SetFocus
        Else

            'Check if any required fields is blank.
164         If Trim$(txtLname.Text) = "" Or Trim$(txtFname.Text) = "" Or Trim$(txtMname.Text) = "" Or Trim$(txtaddress.Text) = "" Then
166             MsgBox "All fields are required.", vbInformation, "Webplus Lending Corporation"
            Else

                'Ask if the you are ready to save this record to database.
168             If rsCustomerDetails.State = 1 Then rsCustomerDetails.Close
170             rsCustomerDetails.Open "Select * from tblCustomer where Middlename = '" & txtMname.Text & "' and Lastname = '" & txtLname.Text & "' and Firstname = '" & txtFname.Text & "'"

172             If rsCustomerDetails.RecordCount <> 0 Then
                    'If there is a duplicate record. This message box will show
174                 MsgBox "Account already exist", vbInformation, "Webplus Lending Corporation"
176                 txtFname.SetFocus
                Else

                    'New Record is saved to database
178                 If rsCustomerDetails.State = 1 Then rsCustomerDetails.Close
                    'Sort records descending
180                 rsCustomerDetails.Open "Select * from tblCustomer"

182                 If MsgBox("Are you sure you want to add new Record?", vbQuestion + vbYesNo, "Webplus Lending Corporation") = vbYes Then
184                     DataGrid1.Refresh
186                     Call autoNumber

188                     With rsCustomerDetails
190                         .AddNew
192                         '!ID = Val(txtID.Text)
194                         !Lastname = Trim$(txtLname.Text)
196                         !Firstname = Trim$(txtFname.Text)
198                         !Middlename = txtMname.Text
200                         !Address = Trim$(txtaddress.Text)
                            !Occupation = Trim$(txtoccupation.Text)
                            !Cellphone = Trim$(txtphone.Text)
218                         !Emailaddress = txtemailadd
222                         !Age = txtage
                            !Gender = cmbGender
224                         !Birthday = dtbday
244                         .Update
                            End With

246                     DataGrid1.Refresh

248                     If rsCustomerDetails.State = 1 Then rsCustomerDetails.Close
250                     rsCustomerDetails.Open "Select * from tblCustomer"
252                     Call autoNumber
254                     txtsearch.Enabled = True
260                     Me.Show
                    End If
                End If
            End If
        End If
        
        
        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_CustomerDetails.btnAdd_Click", Erl

        Resume Next


End Sub
Private Sub ClearControls()

        On Error GoTo ClearControls_Err

        
100     txtLname.Text = ""
104     txtFname.Text = ""
        txtMname.Text = ""
110     txtaddress.Text = ""
128     txtemailadd = ""
        txtoccupation = ""
132     txtage = ""
134     cmbGender = ""
136     txtphone = ""
142     txtID = ""


        Exit Sub

ClearControls_Err:
        ErrReport Err.Description, "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_Customer.ClearControls", Erl

        Resume Next



End Sub


Private Sub btnClose_Click()


        On Error GoTo btnClose_Click_Err

100     txtsearch.Enabled = True

102     If btnClose.Caption = "Close" Then
104         Unload Me
        Else
108         txtLname = ""
110         txtFname = ""
112         txtMname.Text = ""
114         txtaddress.Text = ""
            txtemailadd = ""
            cmbGender = ""
            txtphone = ""
            txtID = ""
148         btnClose.Caption = "Close"
150         btnEdit.Caption = "Edit"
152         btnEdit.Enabled = False
154         btnAdd.Caption = "Add"
156         btnAdd.Enabled = True
160         DataGrid1.Enabled = True
162         DataGrid1.Refresh
164         txtID.Text = Label13.Caption
        End If
        Unload Me
        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Customer.btnClose_Click", _
            Erl

        Resume Next

End Sub

Private Sub btnDelete_Click()

    On Error GoTo btnDelete_Click_Err

116     If MsgBox( _
                "Are you sure you want to delete this user?", _
                vbQuestion + vbYesNo) = vbYes _
                Then
118         rsCustomerDetails.Delete
120         rsCustomerDetails.Update
122         MsgBox _
                    "Record successfully deleted", _
                    vbInformation
124         Unload Me
126         Me.Show
        End If

    Exit Sub

btnDelete_Click_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.btnDelete_Click", _
        Erl
    Resume Next

End Sub



Private Sub btnEdit_Click()
  
  txtsearch.Enabled = False

        If btnEdit.Caption = "Edit" Then
            btnEdit.Caption = "Update"
           
104         txtLname.Enabled = True
            txtFname.Enabled = True
            txtMname.Enabled = True
            txtaddress.Enabled = True
            txtoccupation.Enabled = True
            txtage.Enabled = True
            dtbday.Enabled = True
            txtphone.Enabled = True
            cmbGender.Enabled = True
            txtemailadd.Enabled = True
            txtLname.Enabled = True
            txtLname.Locked = False
            txtFname.Locked = False
            txtMname.Locked = False
            txtaddress.Locked = False
            txtemailadd.Locked = False
            txtage.Locked = False
            dtbday.Enabled = True
            cmbGender.Locked = False
            txtphone.Locked = False
            txtID.Locked = False
            DataGrid1.Enabled = False
            dtbday.Visible = True
            txtLname.SetFocus
        
        Else
        
                Call Customer
                Set DataGrid1.DataSource = rsCustomerDetails

124        If MsgBox("Are you sure you want to update this record?", vbQuestion + vbYesNo, "Webplus Lending Corporation") = vbYes Then

                   
                    With rsCustomerDetails
                        !Lastname = txtLname.Text
                        !Firstname = txtFname.Text
138                     !Middlename = txtMname.Text
                        !Address = Trim$(txtaddress.Text)
                        !Occupation = Trim$(txtoccupation.Text)
                        !Cellphone = Trim$(txtphone.Text)
218                     !Emailaddress = txtemailadd
222                     !Age = txtage
                        !Gender = cmbGender
224                     !Birthday = dtbday
148                     .Update
                    End With
                     btnEdit.Caption = "Edit"
                    rsCustomerDetails.Close
                    End If
                    txtsearch.Enabled = True
                    Me.Show
                    btnClose_Click
                    ClearControls
                End If
 
        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_Customer.btnEdit_Click", Erl

        Resume Next

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdReport_Click()
Form2.Show
Form2.PrintReport
End Sub

Private Sub DataGrid1_Click()

       
        On Error GoTo DataGrid1_Click_Err


100     If rsCustomerDetails.RecordCount = 0 Then
        Else
102         btnAdd.Enabled = True
104         btnClose.Caption = "Cancel"
106         btnEdit.Enabled = True
108         btnDelete.Enabled = True
            

110         With rsCustomerDetails
118             txtLname.Text = !Lastname
120             txtFname.Text = !Firstname
                txtID.Text = !ID
124             txtMname.Text = !Middlename
                txtage.Text = !Age
                dtbday = !Birthday
                txtaddress.Text = !Address
                txtoccupation.Text = !Occupation
                cmbGender = !Gender
                txtphone.Text = !Cellphone
                txtemailadd.Text = !Emailaddress
126
            End With
            

        End If

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Users.DataGrid1_Click", _
            Erl

        Resume Next

 

End Sub


Private Sub Form_Load()

        On Error GoTo Form_Load_Err

      
100     Call connect
102     Call Customer
        'Calls recordset for Collector table
116     Me.Show
118     Me.SetFocus
120     Me.WindowState = vbMaximized
'LoadDataComboID1Type
122     If rsCustomerDetails.State = 1 Then _
            rsCustomerDetails.Close
        'Sort records descending
124     rsCustomerDetails.Open _
            "Select * from tblCustomer"
128     Set DataGrid1.DataSource = rsCustomerDetails
        'Adjusting the width of Fields on Datagird
130     DataGrid1.Width = 19695 ' 19695
132     DataGrid1.Columns(0).Width = 1000
134     DataGrid1.Columns(1).Width = 1000
136     DataGrid1.Columns(4).Width = 1000
138     DataGrid1.Columns(5).Width = 4000
140     DataGrid1.Columns(7).Width = 2000
142     DataGrid1.Columns(10).Width = 1100
        txtLname = ""
146     txtMname = ""
148     txtFname = ""
150     txtage = ""
        'dtbday = ""
152     cmbGender = ""
154     txtphone = ""
158     txtoccupation = ""
160     txtID = ""
164     txtemailadd = ""
        'txtLname.Locked = True
        'txtMname.Locked = True
        'txtFname.Locked = True
        'txtage.Locked = True
       ' cmbGender.Locked = True
       ' txtphone.Locked = True
      '  txtoccupation.Locked = True
       ' txtemailadd.Locked = True
        
        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, _
            "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_Customer.Form_Load", _
            Erl

        Resume Next

        '</EhFooter>
End Sub
Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        '</EhHeader>
100     If rsCustomerDetails.RecordCount = 0 Then
102         txtID.Text = "00001"
104         'lblAuCode.Caption = "00001"'
        Else
106         DataGrid1.Refresh
108         rsCustomerDetails.MoveFirst
110         txtID.Text = "" & Format$(Right$( _
                rsCustomerDetails!ID, 5) + 1, _
                "00000")
112         DataGrid1.Refresh
114         'Label13.Caption = txtID.Text
        End If

        '<EhFooter>
        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, _
            "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_Customer.autoNumber", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtSearch_Change()

        On Error GoTo txtsearch_Change_Err


100     If rsCustomerDetails.State = 1 Then rsCustomerDetails.Close
102     rsCustomerDetails.Open _
           "Select * from tblCustomer where Lastname like '" _
                & txtsearch.Text & _
                "%' or Firstname like '" & _
                txtsearch.Text & _
                "%' or ID like '" & _
                txtsearch.Text & _
                "' Order by ID desc"
             Set DataGrid1.DataSource = rsCustomerDetails


        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Users.txtsearch_Change", _
            Erl

        Resume Next
        

End Sub
