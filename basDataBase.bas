Attribute VB_Name = "basDataBase"
Option Explicit
    
'***********************************************************************************
'Create Database Variables, Language and set Path
'***********************************************************************************

'Create Database and Variables
Public Sub subCreateDB()

    'Set rtBox as RichTextBox
    Dim rtBox As RichTextBox
        Set rtBox = frmDatabase.rtBoxDataBase

    With rtBox
    
        'Add comments
        .Text = .Text & "'This code has been generated using the  ..Database Code Generator.." & vbCrLf
        .Text = .Text & "'Copyright(c) Neil Etherington 2000" & vbCrLf & vbCrLf
        
        'Add Varibles
        .Text = .Text & "'Create Variables" & vbCrLf
        .Text = .Text & "Dim newDBase as DataBase" & vbCrLf
        .Text = .Text & "Dim newTable as TableDef" & vbCrLf
        .Text = .Text & "Dim newField as Field" & vbCrLf
        .Text = .Text & "Dim newIndex as Index" & vbCrLf
        .Text = .Text & "Dim newRelation as Relation" & vbCrLf & vbCrLf
        
        'Add Sub CreateDatabase()
        .Text = .Text & "'Create Database" & vbCrLf
        'Public Sub CreateDatabase()
        .Text = .Text & "Public Sub CreateDatabase()" & vbCrLf & vbCrLf
        'Check to see if Database already exists"
        .Text = .Text & "Dim strReturn as String" & vbCrLf
        .Text = .Text & "Dim strPath as String" & vbCrLf & vbCrLf
        .Text = .Text & vbTab & "Me.MousePointer = vbHourglass" & vbCrLf
        .Text = .Text & vbTab & "'Check to see if Database already exists" & vbCrLf
        .Text = .Text & vbTab & "strPath$ = " & Chr$(34) & "D:\CreateNew.mdb" & Chr$(34) & vbCrLf
        .Text = .Text & vbTab & "If Len(Dir(strPath$)) Then" & vbCrLf
        .Text = .Text & vbTab & vbTab & "strReturn$ =  Msgbox (" & Chr$(34) & "Database already exists. OverWrite.?" & Chr$(34) & "," & "vbCritical" & "+" & "vbYesNo" & ")" & vbCrLf
        .Text = .Text & vbTab & "End If" & vbCrLf
        .Text = .Text & vbTab & "'If confirmation = Yes to Delete, then remove File" & vbCrLf
        .Text = .Text & vbTab & "If strReturn$ = vbYes Then" & vbCrLf
        .Text = .Text & vbTab & vbTab & "Kill (strPath$)" & vbCrLf
        .Text = .Text & vbTab & "Else" & vbCrLf
        .Text = .Text & vbTab & "Exit Sub" & vbCrLf
        .Text = .Text & vbTab & "End If" & vbCrLf & vbCrLf
        'Set newDBase = DBEngine.Workspaces(0).CreateDatabase(" ",)
        .Text = .Text & vbTab & "'Create Database" & vbCrLf
        .Text = .Text & vbTab & "Set newDBase = DBEngine.Workspaces(0).CreateDatabase(" & Chr$(34) & "D:\CreateNew.mdb" & Chr$(34) & "," & frmDatabase.cboLanguage.Text & ")" & vbCrLf & vbCrLf
            
    End With

End Sub
    
'***********************************************************************************
'Create Database Tables, Fields and Indexes
'***********************************************************************************

'Create Table, Field and Index information
Public Sub subCreateTA()

    'Set rtBox as RichTextBox
    Dim rtBox As RichTextBox
        Set rtBox = frmDatabase.rtBoxDataBase
    'Set Field Object
    Dim lstFldName As ListBox: Dim lstFldType As ListBox: Dim lstFldSize As ListBox: Dim lstFldAuto As ListBox
        Set lstFldName = frmDatabase.lstFieldName
        Set lstFldType = frmDatabase.lstFieldType
        Set lstFldSize = frmDatabase.lstFieldSize
        Set lstFldAuto = frmDatabase.lstFieldIncr
    'Set Index Object
    Dim lstIdxKey As ListBox: Dim lstIdxField As ListBox: Dim lstIdxP As ListBox: Dim lstIdxU As ListBox
        Set lstIdxKey = frmDatabase.lstIndexKey
        Set lstIdxField = frmDatabase.lstIndexField
        Set lstIdxU = frmDatabase.lstIndexU
        Set lstIdxP = frmDatabase.lstIndexP
    
    'Loop through Fields
    Dim x As Integer
    'Loop through Index
    Dim y As Integer
    
    With rtBox
        
        'Set newTable = newDBase.CreateTable(" ")
        .Text = .Text & vbTab & "'Create Table" & vbCrLf
        .Text = .Text & vbTab & "Set newTable = newDBase.CreateTableDef(" & Chr$(34) & frmDatabase.txtTableName.Text & Chr$(34) & ")" & vbCrLf & vbCrLf
                
        'Loop through Fields
        For x = 0 To lstFldName.ListCount - 1
            
            'Current ListIndex
            lstFldName.ListIndex = (x)
            lstFldType.ListIndex = (x)
            lstFldSize.ListIndex = (x)
            lstFldAuto.ListIndex = (x)
            
            'Add Field to Table
            If lstFldSize.Text <> "-" Then
                .Text = .Text & vbTab & "'Create Field" & vbCrLf
                .Text = .Text & vbTab & "Set newField = newTable.CreateField(" & Chr$(34) & lstFldName.Text & Chr$(34) & "," & lstFldType.Text & "," & lstFldSize.Text & ")" & vbCrLf
                .Text = .Text & vbTab & vbTab & "'Append Field to Table" & vbCrLf
                .Text = .Text & vbTab & vbTab & "newTable.Fields.Append newField" & vbCrLf & vbCrLf
            Else
                .Text = .Text & vbTab & "'Create Field" & vbCrLf
                .Text = .Text & vbTab & "Set newField = newTable.CreateField(" & Chr$(34) & lstFldName.Text & Chr$(34) & "," & lstFldType.Text & ")" & vbCrLf
                .Text = .Text & vbTab & vbTab & "'Append Field to Table" & vbCrLf
                .Text = .Text & vbTab & vbTab & "newTable.Fields.Append newField" & vbCrLf
            End If
        
            'Add vbAutoIncrField if selected
            If lstFldAuto.Text = "Y" Then
                .Text = .Text & vbTab & vbTab & "'Set Attributes" & vbCrLf
                .Text = .Text & vbTab & vbTab & "newField.Attributes = vbAutoIncrField" & vbCrLf & vbCrLf
            Else
                .Text = .Text & vbCrLf
            End If
            
            'Loop through Index
            For y = 0 To lstIdxKey.ListCount - 1
            
                'Current ListIndex
                lstIdxKey.ListIndex = (y)
                lstIdxField.ListIndex = (y)
                lstIdxP.ListIndex = (y)
                lstIdxU.ListIndex = (y)
                
               'If Index Field matches Field Name
                If lstIdxField.Text = lstFldName.Text Then
                    .Text = .Text & vbTab & "'Create Index Key" & vbCrLf
                    .Text = .Text & vbTab & "Set newIndex = newTable.CreateIndex(" & Chr$(34) & lstIdxKey.Text & Chr$(34) & ")" & vbCrLf
                    .Text = .Text & vbTab & "'Create Index Field" & vbCrLf
                    .Text = .Text & vbTab & "Set newField = newIndex.CreateField(" & Chr$(34) & lstIdxField.Text & Chr$(34) & ")" & vbCrLf
                    .Text = .Text & vbTab & vbTab & "'Append Field" & vbCrLf
                    .Text = .Text & vbTab & vbTab & "newIndex.Fields.Append newField" & vbCrLf
                    
                    'Add Index Parameters
                    If lstIdxP.Text = "P" Then
                        .Text = .Text & vbTab & vbTab & "'Set Index as Primary" & vbCrLf
                        .Text = .Text & vbTab & vbTab & "newIndex.Primary = True" & vbCrLf
                    End If
                    If lstIdxU.Text = "U" Then
                        .Text = .Text & vbTab & vbTab & "'Set Index as Unique" & vbCrLf
                        .Text = .Text & vbTab & vbTab & "newIndex.Unique = True" & vbCrLf
                    End If
                    'Append Index
                    .Text = .Text & vbTab & vbTab & "'Append Index to Table" & vbCrLf
                    .Text = .Text & vbTab & vbTab & "newTable.Indexes.Append newIndex" & vbCrLf & vbCrLf
                End If
                        
                        
            DoEvents
            Next y
        
        DoEvents
        Next x
    
    End With
            
    'Append Table to Database
    With rtBox
        .Text = .Text & vbTab & "'Append Table to Database" & vbCrLf
        .Text = .Text & vbTab & "newDBase.TableDefs.Append newTable" & vbCrLf & vbCrLf
        .Text = .Text & "'***********************************************************************" & vbCrLf & vbCrLf
    End With

End Sub
    
'***********************************************************************************
'Create Database Relationships
'***********************************************************************************

'Create Relationship Information
Public Sub subCreateRL()

    'Set Field Objects
    Dim lstRKey As ListBox: Dim lstRTable As ListBox: Dim lstRForTable As ListBox: Dim lstRFieldKey As ListBox: Dim lstRFldFor As ListBox: Dim lstRAttrib As ListBox
        Set lstRKey = frmDatabase.lstRelKey
        Set lstRTable = frmDatabase.lstRelTable
        Set lstRForTable = frmDatabase.lstRelForTable
        Set lstRFieldKey = frmDatabase.lstRelFieldKey
        Set lstRFldFor = frmDatabase.lstRelFldFor
        Set lstRAttrib = frmDatabase.lstRelAttrib
    Dim rtBox As RichTextBox
        Set rtBox = frmDatabase.rtBoxDataBase
    Dim x As Integer
    
    With rtBox
    
        For x = 0 To lstRKey.ListCount - 1
            
            'Current ListIndex
            lstRKey.ListIndex = (x)
            lstRTable.ListIndex = (x)
            lstRForTable.ListIndex = (x)
            lstRFieldKey.ListIndex = (x)
            lstRFldFor.ListIndex = (x)
            lstRAttrib.ListIndex = (x)
            
            'Create Relation Key
            .Text = .Text & vbTab & "'Create Relation Key" & vbCrLf
            .Text = .Text & vbTab & "Set newRelation = newDBase.CreateRelation(" & Chr$(34) & lstRKey.Text & Chr$(34) & ")" & vbCrLf
            'Select Primary Table
            .Text = .Text & vbTab & vbTab & "'Select Primary Table" & vbCrLf
            .Text = .Text & vbTab & vbTab & "newRelation.Table = (" & Chr$(34) & lstRTable.Text & Chr$(34) & ")" & vbCrLf
            'Select Foreign Table
            .Text = .Text & vbTab & vbTab & "'Select Foreign Table" & vbCrLf
            .Text = .Text & vbTab & vbTab & "newRelation.ForeignTable = (" & Chr$(34) & lstRForTable.Text & Chr$(34) & ")" & vbCrLf & vbCrLf
            'Create Field Key
            .Text = .Text & vbTab & "'Create Field Key" & vbCrLf
            .Text = .Text & vbTab & "Set newField = newRelation.CreateField(" & Chr$(34) & lstRFieldKey.Text & Chr$(34) & ")" & vbCrLf
            'Select Field Foreign Name
            .Text = .Text & vbTab & vbTab & "'Select Field Foreign Name" & vbCrLf
            .Text = .Text & vbTab & vbTab & "newField.ForeignName = (" & Chr$(34) & lstRFldFor.Text & Chr$(34) & ")" & vbCrLf
            'Append Fields
            .Text = .Text & vbTab & vbTab & "'Append Fields" & vbCrLf
            .Text = .Text & vbTab & vbTab & "newRelation.Fields.Append newField" & vbCrLf
            'Set Attributes of Relation
            .Text = .Text & vbTab & vbTab & "'Set Relation Attributes" & vbCrLf
            .Text = .Text & vbTab & vbTab & "newRelation.Attributes = " & lstRAttrib.Text & vbCrLf
            'Append Relation to Database
            .Text = .Text & vbTab & vbTab & "'Append Relation to Database" & vbCrLf
            .Text = .Text & vbTab & vbTab & "newDBase.Relations.Append newRelation" & vbCrLf & vbCrLf
        DoEvents
        Next x
      
    End With

End Sub
    
'***********************************************************************************
'Close Database
'***********************************************************************************

'Close Database
Public Sub CloseDatabase()
    
    Dim rtBox As RichTextBox
        Set rtBox = frmDatabase.rtBoxDataBase
    
    With rtBox
        'Finish Database
        .Text = .Text & vbTab & "Me.MousePointer = vbDefault" & vbCrLf
        .Text = .Text & vbTab & "MsgBox" & Chr$(34) & "Database sucessfully created" & Chr$(34) & "," & "vbInformation " & "+ " & "vbOkOnly" & vbCrLf & vbCrLf
        .Text = .Text & "End Sub"
    End With
      
End Sub
