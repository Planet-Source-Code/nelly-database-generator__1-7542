VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDatabase 
   Caption         =   "Database Code Generator"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopyClip 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Left            =   12960
      TabIndex        =   62
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCloseDBase 
      Caption         =   "Close Database"
      Height          =   375
      Left            =   7800
      TabIndex        =   61
      Top             =   10320
      Width           =   2415
   End
   Begin VB.ComboBox cboRelAttrib 
      Height          =   315
      Left            =   12240
      TabIndex        =   20
      Top             =   7080
      Width           =   2655
   End
   Begin VB.ListBox lstRelAttrib 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   12240
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton cmdAppendRel 
      Caption         =   "Create Relation"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton cmdRelation 
      Caption         =   "Add Relation Information"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   7560
      Width           =   14655
   End
   Begin VB.TextBox txtRelFieldForName 
      Height          =   315
      Left            =   9840
      TabIndex        =   19
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtRelFieldKeyName 
      Height          =   315
      Left            =   7440
      TabIndex        =   18
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtRelForTable 
      Height          =   315
      Left            =   5040
      TabIndex        =   17
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtRelTableName 
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtRelKeyName 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   7080
      Width           =   2295
   End
   Begin VB.ListBox lstRelFldFor 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   9840
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ListBox lstRelFieldKey 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   7440
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ListBox lstRelForTable 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   5040
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ListBox lstRelTable 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   2640
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ListBox lstRelKey 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   240
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ComboBox cboFieldIncr 
      Height          =   315
      Left            =   8160
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.ListBox lstFieldIncr 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   8040
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "F"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdIndex 
      Caption         =   "Add Index Information"
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   3960
      Width           =   5415
   End
   Begin VB.ComboBox cboIndexU 
      Height          =   315
      Left            =   14400
      TabIndex        =   12
      Top             =   3480
      Width           =   495
   End
   Begin VB.ComboBox cboIndexP 
      Height          =   315
      Left            =   13800
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txtIndexField 
      Height          =   315
      Left            =   11640
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtIndexKey 
      Height          =   315
      Left            =   9480
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.ListBox lstIndexU 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   14280
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.ListBox lstIndexP 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   13680
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.ListBox lstIndexField 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   11520
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ListBox lstIndexKey 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   9480
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdField 
      Caption         =   "Add Field Information"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   9015
   End
   Begin VB.ListBox lstFieldSize 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   6840
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "F"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox lstFieldType 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   4320
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "F"
      Top             =   1080
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox rtBoxDataBase 
      Height          =   2175
      Left            =   240
      TabIndex        =   31
      Top             =   8040
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmDatabase.frx":0000
   End
   Begin VB.CommandButton cmdAppendTable 
      Caption         =   "Append Table info to Database"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   10320
      Width           =   2415
   End
   Begin VB.ComboBox cboFieldSize 
      Height          =   315
      Left            =   6960
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox cboFieldType 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtFieldName 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   3975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Database"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   10320
      Width           =   2415
   End
   Begin VB.TextBox txtTableName 
      Height          =   315
      Left            =   12240
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      Left            =   9360
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtDbaseName 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
   Begin VB.ListBox lstFieldName 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "F"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Relation Attributes"
      Height          =   195
      Left            =   12240
      TabIndex        =   60
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Relation Field Foreign Name"
      Height          =   195
      Left            =   9840
      TabIndex        =   58
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Relation Field Key Name"
      Height          =   195
      Left            =   7440
      TabIndex        =   57
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Relation Foreign Table"
      Height          =   195
      Left            =   5040
      TabIndex        =   56
      Top             =   4680
      Width           =   1605
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Relation Table Name"
      Height          =   195
      Left            =   2640
      TabIndex        =   55
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Relation Key Name"
      Height          =   195
      Left            =   240
      TabIndex        =   54
      Top             =   4680
      Width           =   1365
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Auto Icr Field"
      Height          =   195
      Left            =   8160
      TabIndex        =   48
      Top             =   3240
      Width           =   930
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Auto Icr Field"
      Height          =   195
      Left            =   8040
      TabIndex        =   47
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Primary   Unique"
      Height          =   195
      Left            =   13680
      TabIndex        =   45
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Index Field Name"
      Height          =   195
      Left            =   11520
      TabIndex        =   44
      Top             =   840
      Width           =   1230
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Index Key Name"
      Height          =   195
      Left            =   9480
      TabIndex        =   43
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Select Pi or Un"
      Height          =   195
      Left            =   13800
      TabIndex        =   42
      Top             =   3240
      Width           =   1065
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Enter Index Field Name"
      Height          =   195
      Left            =   11640
      TabIndex        =   41
      Top             =   3240
      Width           =   1650
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Enter Index Key Name"
      Height          =   195
      Left            =   9480
      TabIndex        =   40
      Top             =   3240
      Width           =   1590
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Field Size"
      Height          =   195
      Left            =   6840
      TabIndex        =   35
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Field Type"
      Height          =   195
      Left            =   4320
      TabIndex        =   34
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Field Size"
      Height          =   195
      Left            =   6960
      TabIndex        =   30
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Select a Field Type"
      Height          =   195
      Left            =   4440
      TabIndex        =   29
      Top             =   3240
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter Field Name"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Field Name"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter Table Name"
      Height          =   195
      Left            =   12240
      TabIndex        =   26
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select Language"
      Height          =   195
      Left            =   9360
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Database Path and Name"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   2265
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuH1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuH2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRunRDemo 
         Caption         =   "Run Demo"
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
      Begin VB.Menu mnuH3 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuClearFields 
         Caption         =   "Clear Fields"
      End
      Begin VB.Menu mnuH4 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuClearIndexes 
         Caption         =   "Clear Indexes"
      End
      Begin VB.Menu mnuH5 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuClearRelations 
         Caption         =   "Clear Relations"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHTRTP 
         Caption         =   "How to use this program"
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************************************************
'Form Events
'***********************************************************************************

'Load Form
Private Sub Form_Load()
    
    'Used for adding Field Type size
    Dim x As Byte
    
    'Add Language to cboLanguage
    cboLanguage.AddItem "dbLangGeneral"
    cboLanguage.ListIndex = 0
    
    'Add Field Type to cboFieldType
    With cboFieldType
        .AddItem "dbBinary"
        .AddItem "dbBoolean"
        .AddItem "dbByte"
        .AddItem "dbCurrency"
        .AddItem "dbDate"
        .AddItem "dbDouble"
        .AddItem "dbGUID"
        .AddItem "dbInteger"
        .AddItem "dbLong"
        .AddItem "dbLongBinary"
        .AddItem "dbMemo"
        .AddItem "dbSingle"
        .AddItem "dbText"
        .ListIndex = 12
    End With
    
    'Add Field size to cboFieldSize
    With cboFieldSize
        .AddItem "-"
        For x = 1 To 100
            .AddItem (x)
        Next x
        .ListIndex = 0
    End With
    
    'Add Database Path and Name to txtDbaseName
    txtDbaseName.Text = "D:\CreateNew.mdb"
    
    'Add Index Primary to cboIndexP
    With cboIndexP
        .AddItem "-"
        .AddItem "P"
        .ListIndex = 0
    End With
    
    'Add Index Unique to cboIndexU
    With cboIndexU
        .AddItem "-"
        .AddItem "U"
        .ListIndex = 0
    End With

    'Add AutoIncrField option to Field
    With cboFieldIncr
        .AddItem "Y"
        .AddItem "N"
        .ListIndex = 1
    End With

    'Add Relation Attributes to cboRelAttrib
    With cboRelAttrib
        .AddItem "dbRelationUnique"
        .AddItem "dbRelationDontEnforce"
        .AddItem "dbRelationInherited"
        .AddItem "dbRelationUpdateCascade"
        .AddItem "dbRelationDeleteCascade"
        .ListIndex = 4
    End With

    'Disable cmdAppendTable and cmdAppendRel and cmdCloseDbase
    cmdAppendTable.Enabled = False
    cmdAppendRel.Enabled = False
    cmdCloseDBase.Enabled = False

End Sub

'Unload Form
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frmDatabase = Nothing
End Sub

'***********************************************************************************
'Command Events
'***********************************************************************************

'Add Field Information
Private Sub cmdField_Click()

    'Error handler
    If txtFieldName.Text = Trim("") Or cboFieldType.Text = Trim("") Or cboFieldSize.Text = Trim("") Or cboFieldIncr.Text = Trim("") Then
        MsgBox "All Fields must be completed", vbCritical + vbOKOnly, App.EXEName
    Exit Sub
    End If

    'Add Field information
    lstFieldName.AddItem txtFieldName.Text
    lstFieldType.AddItem cboFieldType.Text
    lstFieldSize.AddItem cboFieldSize.Text
    lstFieldIncr.AddItem cboFieldIncr.Text
    
End Sub

'Add Index Information
Private Sub cmdIndex_Click()
    
    'Error handler
    If txtIndexKey.Text = Trim("") Or txtIndexField.Text = Trim("") Or cboIndexP.Text = Trim("") Or cboIndexU.Text = Trim("") Then
        MsgBox "All Index Fields must be completed", vbCritical + vbOKOnly, App.EXEName
    Exit Sub
    End If
        
    'Add Index information
    lstIndexKey.AddItem txtIndexKey.Text
    lstIndexField.AddItem txtIndexField.Text
    lstIndexP.AddItem cboIndexP.Text
    lstIndexU.AddItem cboIndexU.Text
           
End Sub

'Add Relation Information
Private Sub cmdRelation_Click()

    'Error handler
    If txtRelKeyName.Text = Trim("") Or txtRelTableName.Text = Trim("") Or txtRelForTable.Text = Trim("") Or txtRelFieldKeyName.Text = Trim("") Or txtRelFieldForName.Text = Trim("") Or cboRelAttrib.Text = Trim("") Then
        MsgBox "All Relation Fields must be completed", vbCritical + vbOKOnly, App.EXEName
    Exit Sub
    End If
    
    'Add Relation Information
    lstRelKey.AddItem txtRelKeyName.Text
    lstRelTable.AddItem txtRelTableName.Text
    lstRelForTable.AddItem txtRelForTable.Text
    lstRelFieldKey.AddItem txtRelFieldKeyName.Text
    lstRelFldFor.AddItem txtRelFieldForName.Text
    lstRelAttrib.AddItem cboRelAttrib.Text

End Sub

'***********************************************************************************
'Command Database Creation Events
'***********************************************************************************

'Command Create Database
Private Sub cmdCreate_Click()
    basDataBase.subCreateDB
    cmdAppendTable.Enabled = True
End Sub

'Command Append Table, Field and Index
Private Sub cmdAppendTable_Click()
    
    'Error handler
    If txtTableName.Text = Trim("") Then
        MsgBox "Please enter a Table Name", vbCritical + vbOKOnly, App.EXEName
    Exit Sub
    End If
    'Create Table
    basDataBase.subCreateTA
    'Enable cmdAppendRel and cmdCloseDBase
    cmdAppendRel.Enabled = True
    cmdCloseDBase.Enabled = True
    
End Sub

'Command Create Relation
Private Sub cmdAppendRel_Click()
    basDataBase.subCreateRL
End Sub

'Command close Database
Private Sub cmdCloseDBase_Click()
    basDataBase.CloseDatabase
End Sub

'Command copy to Clipboard
Private Sub cmdCopyClip_Click()
    
    Dim strClip As String
    
    'Copy the Database Text into the Variable strClip
    With rtBoxDataBase
        .SelStart = 0
        .SelLength = Len(rtBoxDataBase.Text)
        strClip$ = .SelText
    End With
    
    'Clear the Clipboard before transfering the Database from the Variable strClip to the Clipboard
    Clipboard.Clear
    Clipboard.SetText strClip$, 1

    MsgBox "Database copied to Clipboard", vbInformation + vbOKOnly, App.EXEName

End Sub

'***********************************************************************************
'Menu Commands
'***********************************************************************************

'Menu Exit
Private Sub mnuFileExit_Click()
    Form_Unload (2)
    frmHowTo.mnuFileExit_Click
End Sub

'Menu Run Demo
Private Sub mnuRunRDemo_Click()
    basDemoDBase.CreateDemoDB
End Sub

'Menu Clear Fields
Private Sub mnuClearFields_Click()
    'Clear Field Lists
    txtTableName.Text = ""
    lstFieldName.Clear
    lstFieldType.Clear
    lstFieldSize.Clear
    lstFieldIncr.Clear
End Sub

'Menu Clear Indexes
Private Sub mnuClearIndexes_Click()
    'Clear Index Lists
    lstIndexKey.Clear
    lstIndexField.Clear
    lstIndexP.Clear
    lstIndexU.Clear
End Sub

'Menu Clear Relations
Private Sub mnuClearRelations_Click()
    'Clear Relation Lists
    lstRelKey.Clear
    lstRelTable.Clear
    lstRelForTable.Clear
    lstRelFieldKey.Clear
    lstRelFldFor.Clear
    lstRelAttrib.Clear
End Sub

'Menu How to run this program
Private Sub mnuHelpHTRTP_Click()
    frmHowTo.Show
End Sub

