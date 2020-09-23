VERSION 5.00
Begin VB.Form frmHowTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database HowTo:"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHowTo.frx":0000
      Top             =   120
      Width           =   15015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuh1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmHowTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Text1.Text = " How to work this Program:" & vbCrLf & vbCrLf _
    & "1. Enter the Path and Name of the Database under the Label -Enter Database Path and Name-." & vbCrLf _
    & "2. Select the only Language thats available under the Label -Select Language_.  Depress the Command Button -Create Database-." & vbCrLf _
    & "3. Enter the first Table Name under the Label -Enter Table Name-." & vbCrLf _
    & "4. Enter all the Fields associated with this Table using the Field controls and depress the Command Button -Add Field Information- on each Field to add." & vbCrLf _
    & "5. Enter the Indexes associated with this Table using the Index controls. Add each Index using the Command Button -Add Index Information-." & vbCrLf _
    & "6. Once you have entered all the Fields and Indexes associated with this Table, Depress the Command Button -Append Table info to Database-. repeat the process with the second Table. Leaving the Database Path,Name and Language as is." & vbCrLf _
    & "7. Now all you have to do is use the Relationship controls to set the Relations between each Table. Enter each Relation into the Relationship controls and depress the Command Button -Add Relation Information-" & vbCrLf _
    & "8. When all Relations are added into the ListBoxes depress the Command Button -Create Relation-." & vbCrLf _
    & "9. Nearly their. Depress the Command button -Close Database-. If you have entered all the information correctly, just copy and Paste the code created in the RichTextBox into a new VBProject. Make the reference to Microsoft DAO 3.51 Object Library, all thats required then is a reference to this Procedure eg. Command_Click"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Text1.Text = ""
    Unload Me
    Set frmHowTo = Nothing
End Sub

Public Sub mnuFileExit_Click()
    Form_Unload 1
End Sub
