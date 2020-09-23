Attribute VB_Name = "basDemoDBase"
Option Explicit

'Create a Demo Database
Public Sub CreateDemoDB()

    'I`ll be using a MsgBox to indicate whats happening as we go through this
    'Demo of creating a Database

    'Used for returning response from MsgBox
    Dim strRet As String
        

    'Message displays description of Database we are creating
    strRet$ = MsgBox("Before we write any code, here is a description of the Database we are going to create." & vbCrLf _
    & "The Database will have two Tables, one called -Employee- and the second called -CompanyData-." & vbCrLf _
    & "The Employee Table will have the Fields -EmployeeID-, which we will make our Primary Index with a Unique" & vbCrLf _
    & "Value, AutoIncrNumber and in which will make the Relationship between this and the -CompanyData- Table." & vbCrLf _
    & "The following Fields will also reside in this Table, FirstName with Index, LastName with Index," & vbCrLf _
    & "Address, BirthDay, HomePhone, Mobile No., and E.Mail Address." & vbCrLf _
    & "The second Table will contain the Fields ContactID, identical to the Employee Table, Clock Card No," & vbCrLf _
    & "Days absent from work and Job description." & vbCrLf & vbCrLf _
    & "Table Names --" & vbTab & vbTab & "Employee" & vbTab & vbTab & vbTab & "CompanyData" & vbCrLf & vbCrLf _
    & "Field Names --" & vbTab & vbTab & "EmployeeID" & vbTab & vbTab & "EmployeeID" & vbCrLf _
    & vbTab & vbTab & vbTab & "FirstName" & vbTab & vbTab & vbTab & "Clock Card No" & vbCrLf _
    & vbTab & vbTab & vbTab & "LastName" & vbTab & vbTab & "Days Absent from Work" & vbCrLf _
    & vbTab & vbTab & vbTab & "Address" & vbTab & vbTab & vbTab & "Job Description" & vbCrLf _
    & vbTab & vbTab & vbTab & "BirthDay" & vbCrLf _
    & vbTab & vbTab & vbTab & "HomePhoone No" & vbCrLf _
    & vbTab & vbTab & vbTab & "Mobile Phone No" & vbCrLf _
    & vbTab & vbTab & vbTab & "E.Mail Address" & vbCrLf & vbCrLf _
    & "Continue ?", vbInformation + vbYesNo)
    
    'Wait for response from MsgBox
    If strRet$ = vbNo Then
    Exit Sub
    Else
    End If

    'Create Database Path and Name
    basDataBase.subCreateDB

    'MsgBox displays information on current state of Database
    strRet$ = MsgBox("We have just created the Variables, set the Variables to represent the Database Objects," & vbCrLf _
    & "checked to see if the Database exists and write the code that creates the DataBase in the specified Path" & vbCrLf & vbCrLf _
    & "Continue ?", vbInformation + vbYesNo)
    
    'Wait for response from MsgBox
    If strRet$ = vbNo Then
    Exit Sub
    Else
    End If

    'Enter Table and Field information
    With frmDatabase
    
        'Add Table name
        .txtTableName.Text = "Employee"
    
        'Add Fields
        .lstFieldName.AddItem "EmployeeID"
        .lstFieldName.AddItem "FirstName"
        .lstFieldName.AddItem "LastName"
        .lstFieldName.AddItem "Address"
        .lstFieldName.AddItem "BirthDay"
        .lstFieldName.AddItem "HomePhone"
        .lstFieldName.AddItem "MobilePhoneNo"
        .lstFieldName.AddItem "EMail"
        
        'Add Field Type
        .lstFieldType.AddItem "dbLong"
        .lstFieldType.AddItem "dbText"
        .lstFieldType.AddItem "dbText"
        .lstFieldType.AddItem "dbText"
        .lstFieldType.AddItem "dbDate"
        .lstFieldType.AddItem "dbText"
        .lstFieldType.AddItem "dbText"
        .lstFieldType.AddItem "dbText"
        
        'Add Field Size
        .lstFieldSize.AddItem "-"
        .lstFieldSize.AddItem "15"
        .lstFieldSize.AddItem "20"
        .lstFieldSize.AddItem "40"
        .lstFieldSize.AddItem "-"
        .lstFieldSize.AddItem "15"
        .lstFieldSize.AddItem "20"
        .lstFieldSize.AddItem "30"
        
        'Add AutoIncrField, True or False
        .lstFieldIncr.AddItem "Y"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        
        'Add Listindex EmployeeID
        .lstIndexKey.AddItem "PrimaryKey"
        .lstIndexField.AddItem "EmployeeID"
        .lstIndexP.AddItem "P"
        .lstIndexU.AddItem "U"
        
        'Add Listindex FirstName
        .lstIndexKey.AddItem "SecondaryA"
        .lstIndexField.AddItem "FirstName"
        .lstIndexP.AddItem "-"
        .lstIndexU.AddItem "-"
        
        'Add Listindex LastName
        .lstIndexKey.AddItem "SecondaryB"
        .lstIndexField.AddItem "LastName"
        .lstIndexP.AddItem "-"
        .lstIndexU.AddItem "-"
        
    End With

    'Append Richtextbox Data with new Table
    basDataBase.subCreateTA
    
    strRet$ = MsgBox("We have now entered the Table Name and all the Fields associated with the Table, including all" & vbCrLf _
    & "Parameters etc. We will now create the second Table and associated Fields." & vbCrLf & vbCrLf _
    & "Continue ?", vbInformation + vbYesNo)
    
    'Wait for response from MsgBox
    If strRet$ = vbNo Then
    Exit Sub
    Else
    End If
        
    'Clear Listboxes for next Table
    ClearListBoxes

    'Enter second Table and Field information
    With frmDatabase
    
        'Add Table name
         .txtTableName.Text = "CompanyData"
    
        'Add Fields
        .lstFieldName.AddItem "EmployeeID"
        .lstFieldName.AddItem "ClockCardNo"
        .lstFieldName.AddItem "DaysAbsent"
        .lstFieldName.AddItem "JobDescription"
        
        'Add Field Type
        .lstFieldType.AddItem "dbLong"
        .lstFieldType.AddItem "dbText"
        .lstFieldType.AddItem "dbInteger"
        .lstFieldType.AddItem "dbText"
        
        'Add Field Size
        .lstFieldSize.AddItem "-"
        .lstFieldSize.AddItem "15"
        .lstFieldSize.AddItem "-"
        .lstFieldSize.AddItem "40"
        
        'Add AutoIncrField, True or False
        .lstFieldIncr.AddItem "Y"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        .lstFieldIncr.AddItem "N"
        
        'Add Listindex EmployeeID
        .lstIndexKey.AddItem "PrimaryKey"
        .lstIndexField.AddItem "EmployeeID"
        .lstIndexP.AddItem "P"
        .lstIndexU.AddItem "U"

    End With

    'Append Richtextbox Data with new Table
    basDataBase.subCreateTA
    
    'Second Table completed, all we have to do now is make the Relation between the the two Tables
    strRet$ = MsgBox("We have now created the two Tables and all associated Fields and Indexes. All their is to do" & vbCrLf _
    & "now is add the Relation between Table Employee and Table CompanyData" & vbCrLf & vbCrLf _
    & "Continue ?", vbInformation + vbYesNo)
    
    'Wait for response from MsgBox
    If strRet$ = vbNo Then
    Exit Sub
    Else
    End If

    With frmDatabase
        .lstRelKey.AddItem "NewRelation"
        .lstRelTable.AddItem "Employee"
        .lstRelForTable.AddItem "CompanyData"
        .lstRelFieldKey.AddItem "EmployeeID"
        .lstRelFldFor.AddItem "EmployeeID"
        .lstRelAttrib.AddItem "dbRelationDeleteCascade"
    End With

    basDataBase.subCreateRL
    
    basDataBase.CloseDatabase

    MsgBox "Database created sucessfully, Copy and paste this code into" & vbCrLf _
    & "a new VBProject, make the reference to Microsoft DAO 3.51 Object Library," & vbCrLf _
    & "all thats required then is a reference to this Procedure eg. Command_Click", vbInformation + vbOKOnly, App.EXEName

End Sub

'Clear ListBoxes
Private Sub ClearListBoxes()

    'Clear Listboxes for next Table
    With frmDatabase
        .lstFieldName.Clear
        .lstFieldType.Clear
        .lstFieldSize.Clear
        .lstFieldIncr.Clear
        .lstIndexKey.Clear
        .lstIndexField.Clear
        .lstIndexP.Clear
        .lstIndexU.Clear
    End With

End Sub
