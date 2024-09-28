Attribute VB_Name = "ScriptAccessForSqlServer"
Option Compare Database
Option Explicit

' Add the following references:
'    Microsoft ADO Ext X.X For DDL and Security
'    Microsoft Office XX.X Object Library

'USAGE: Import as a module into your Access project and run 'ScriptDatabase' ... that's it.

Sub ScriptDatabase()
'PURPOSE: Scripts an entire Access database for SQL Server and saves it in the file selected

    Dim fd As FileDialog
    Dim intFile As Integer

    'Create a FileDialog object as a Save As Dialog to choose where to save the script to
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    'Setup the File Dialog Box
    With fd
        .Title = "Choose where to save the script"
        .ButtonName = "Save"
        .InitialFileName = "*.sql"   'Forces the filename to have a .sql at the end
        .AllowMultiSelect = False
    End With
    
    'Use the Show method to display the File Picker dialog box and return the user's action.
    'If the user pressed the action button then continue.
    If fd.Show = -1 Then
        
        intFile = FreeFile
        Open fd.SelectedItems(1) For Binary Access Write As #intFile
        Put #intFile, , GenerateCreateTableSQL
        Put #intFile, , GeneratePrimaryKeySQL
        Put #intFile, , GenerateDefaultValueSQL
        Put #intFile, , GenerateUniqueIndexSQL
        Put #intFile, , GenerateNonUniqueIndexSQL
        Put #intFile, , GenerateInsertIntoSQL
        Put #intFile, , GenerateForeignKeySQL
        Close #intFile
        
        MsgBox "Script Saved.", vbOKOnly, "Complete"
        
    End If

End Sub


Private Function GenerateCreateTableSQL() As String
'PURPOSE: Scripts Access Tables for SQL Server
'RETURNS: SQL Server script as a string
    
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim key As DAO.Field
    
    Dim cat As ADOX.Catalog
    Dim col As ADOX.Column
    
    Dim strSQL As String
    Dim blnNOT As Boolean
    
    Dim i As Integer
    
    'Start the meter
    SysCmd acSysCmdInitMeter, "Scripting Tables", CurrentDb.TableDefs.Count
    
    'Start strSQL
    strSQL = vbCrLf & "/****** TABLES as of " & Now() & " ******/" & vbCrLf & vbCrLf
    
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection

    'loop through each table in the database.  Start with DAO because it retrieves the tables and columns
    'in the order that they are designed (ADOX retrieves columns alphabetically only)
    For Each tdf In CurrentDb.TableDefs
        'Omits system tables
        If tdf.Attributes = 0 Then
            strSQL = strSQL & "CREATE TABLE [dbo].[" & tdf.Name & "] (" & vbCrLf
            'cycle through each field
            For Each fld In tdf.Fields
                'Switch over to ADOX for the column because these are better properties to help us
                Set col = cat.Tables(tdf.Name).Columns(fld.Name)
                blnNOT = False
                strSQL = strSQL + "    [" & col.Name & "] "
                Select Case col.Type
                    
                    'Text
                    Case adChar
                        strSQL = strSQL + "char(" & col.DefinedSize & ")"
                    Case adWChar
                        strSQL = strSQL + "nchar(" & col.DefinedSize & ")"
                    Case adVarChar
                        strSQL = strSQL + "varchar(" & col.DefinedSize & ")"
                    Case adVarWChar
                        strSQL = strSQL + "nvarchar(" & col.DefinedSize & ")"
                    Case adLongVarChar
                        strSQL = strSQL + "varchar(max)"
                    Case adLongVarWChar
                        strSQL = strSQL + "nvarchar(max)"
                    
                    'Binary
                    Case adBinary
                        strSQL = strSQL + "binary(" & col.DefinedSize & ")"
                    Case adVarBinary
                        strSQL = strSQL + "varbinary(" & col.DefinedSize & ")"
                    Case adLongVarBinary
                        strSQL = strSQL + "varbinary(max)"
                    
                    'DateTime
                    Case adDate
                        strSQL = strSQL + "datetime"
                    Case adDBDate
                        strSQL = strSQL + "date"
                    Case adDBTime
                        strSQL = strSQL + "time"
                    Case adDBTimeStamp
                        strSQL = strSQL + "timestamp"
                    
                    'Numeric
                    Case adBoolean
                        strSQL = strSQL + "bit"
                    Case adTinyInt, adUnsignedTinyInt
                        strSQL = strSQL + "tinyint"
                    Case adSmallInt, adUnsignedSmallInt
                        strSQL = strSQL + "smallint"
                    Case adInteger, adUnsignedInt
                        strSQL = strSQL + "int"
                    Case adBigInt, adUnsignedBigInt
                        strSQL = strSQL + "bigint"
                    Case adSingle
                        strSQL = strSQL + "real"
                    Case adDouble
                        strSQL = strSQL + "float"
                    Case adDecimal
                        strSQL = strSQL + "decimal(" & col.Precision & "," & col.NumericScale & ")"
                    Case adNumeric, adVarNumeric
                        strSQL = strSQL + "numeric(" & col.Precision & "," & col.NumericScale & ")"
                    Case adCurrency
                        strSQL = strSQL + "money"
                    
                    'Other
                    Case adGUID
                        strSQL = strSQL + "uniqueidentifier"
                    
                    'Anything else ... shouldn't happen but just in case
										Case Else
                        strSQL = strSQL + "sql_variant"
                    
                End Select
                
                'Check if the field is part of the primary key
                For Each idx In tdf.Indexes
                    'If the index is a primary key
                    If idx.Primary Then
                        For Each key In idx.Fields
                            If key.Name = col.Name Then blnNOT = True
                        Next key
                    End If
                Next idx
                
                'Check if the field should not be null for other reasons
                If col.Type = adBoolean Or col.Properties!Nullable = False Or col.Properties!Autoincrement Then blnNOT = True
                If blnNOT Then strSQL = strSQL & " NOT"
                strSQL = strSQL & " NULL"
                If col.Properties!Autoincrement Then
                    strSQL = strSQL & " IDENTITY(" & col.Properties!Seed & "," & col.Properties!Increment & ")"
                End If
                strSQL = strSQL & "," & vbCrLf
                
            Next fld
            strSQL = Left(strSQL, Len(strSQL) - 3) & vbCrLf & ")" & vbCrLf & "GO" & vbCrLf
        End If
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    Next tdf
    
    GenerateCreateTableSQL = strSQL
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set tdf = Nothing
    Set fld = Nothing
    
    Set cat = Nothing
    Set col = Nothing
    Set key = Nothing
    Set idx = Nothing
End Function


Private Function GeneratePrimaryKeySQL() As String
'PURPOSE: Scripts Primary Keys from Access Tables for SQL Server
'RETURNS: SQL Server script as a string

    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strKeys As String
    Dim strSQL As String
    Dim intCount As Integer
    Dim i As Integer
    
    'Start the meter
    SysCmd acSysCmdInitMeter, "Scripting Primary Keys", CurrentDb.TableDefs.Count
    
    'Start strSQL
    strSQL = vbCrLf & "/****** PRIMARY KEYS as of " & Now() & " ******/" & vbCrLf & vbCrLf
    
    'loop through each table in the database
    For Each tdf In CurrentDb.TableDefs
        'Omits system tables
        If tdf.Attributes = 0 Then
            strSQL = strSQL & ""
            'Cycle through each index
            For Each idx In tdf.Indexes
                strKeys = ""
                'If the index is a primary key
                If idx.Primary Then
                    For Each fld In idx.Fields
                        strKeys = strKeys & "[" & fld.Name & "],"
                    Next fld
                    strKeys = Left(strKeys, Len(strKeys) - 1)
                    strSQL = strSQL & "ALTER TABLE dbo.[" & tdf.Name & "]" & vbCrLf & _
                                      "    ADD CONSTRAINT PK_" & Replace(tdf.Name, " ", "_") & " " & _
                                      "PRIMARY KEY (" & strKeys & ");" & vbCrLf & "GO" & vbCrLf
                End If
            Next idx
        End If
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    Next tdf
    
    GeneratePrimaryKeySQL = strSQL
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set tdf = Nothing
    Set idx = Nothing
    Set fld = Nothing
End Function

Private Function GenerateDefaultValueSQL() As String
'PURPOSE: Scripts Default Value Constraints from Access Tables for SQL Server
'RETURNS: SQL Server script as a string
    
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strKeys As String
    Dim strSQL As String
    Dim i As Integer
    
    'Start the meter
    SysCmd acSysCmdInitMeter, "Scripting Default Values", CurrentDb.TableDefs.Count
    
    'Start strSQL
    strSQL = vbCrLf & "/****** DEFAULT VALUES as of " & Now() & " ******/" & vbCrLf & vbCrLf
    
    'loop through each table in the database
    For Each tdf In CurrentDb.TableDefs
        'Omits system tables
        If tdf.Attributes = 0 Then
            'cycle through each field
            For Each fld In tdf.Fields
                If fld.DefaultValue <> "" Then
                    strSQL = strSQL & "ALTER TABLE dbo.[" & tdf.Name & "]" & vbCrLf & _
                                      "    ADD CONSTRAINT DF_" & Replace(tdf.Name, " ", "_") & "__" & Replace(fld.Name, " ", "_") & " DEFAULT "
                    Select Case fld.Type
                        Case dbBinary, dbChar, dbGUID, dbMemo, dbText, dbVarBinary
                            strSQL = strSQL & "N'" & Replace(fld.DefaultValue, "'", "''") & "' "
                        Case dbDate, dbTime, dbTimeStamp
                            strSQL = strSQL & "CAST('" & fld.DefaultValue & "' AS DATETIME) "
                        Case dbBoolean
                            Select Case CStr(fld.DefaultValue)
                                Case "No", "False", "Off", "0"
                                    strSQL = strSQL & "0" & " "
                                Case "Yes", "True", "On", "1"
                                    strSQL = strSQL & "1" & " "
                            End Select
                        Case Else
                            strSQL = strSQL & fld.DefaultValue & " "
                    End Select
                    strSQL = strSQL & "FOR [" & fld.Name & "];" & vbCrLf
                    strSQL = strSQL & "GO" & vbCrLf
                End If
            Next fld
        End If
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    Next tdf
    
    GenerateDefaultValueSQL = strSQL
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set tdf = Nothing
    Set fld = Nothing
End Function


Private Function GenerateForeignKeySQL() As String
'PURPOSE: Scripts Foreign Keys from Access Tables for SQL Server
'RETURNS: SQL Server script as a string
    
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim strFields As String
    Dim strParentFields As String
    Dim strSQL As String
    Dim strIntegrity As String
    Dim intAttr As Long
    Dim strName As String
    Dim strNames As String
    Dim intCount As Long
    Dim i As Integer
    
    'Start the meter
    SysCmd acSysCmdInitMeter, "Scripting Foreign Keys", CurrentDb.Relations.Count
    
    'Start strSQL
    strSQL = vbCrLf & "/****** FOREIGN KEYS as of " & Now() & " ******/" & vbCrLf & vbCrLf
    
    For Each rel In CurrentDb.Relations
        If CurrentDb.TableDefs(rel.Table).Attributes = 0 Then
        
            strName = "FK_" & Replace(rel.ForeignTable, " ", "_") & "__" & Replace(rel.Table, " ", "_")
            strSQL = strSQL & "ALTER TABLE dbo.[" & rel.ForeignTable & "]" & vbCrLf & _
                              "    ADD CONSTRAINT " & strName
            'Count how many times this name has been used in earlier loops
            intCount = (Len(strNames) - Len(Replace(strNames, strName & ",", ""))) / Len(strName & ",")
            If intCount > 0 Then strSQL = strSQL & intCount + 1
            strNames = strNames & strName & ","
            strSQL = strSQL & vbCrLf
            
            strFields = ""
            strParentFields = ""
            strIntegrity = ""
            For Each fld In rel.Fields
                strFields = strFields & "[" & fld.ForeignName & "],"
                strParentFields = strParentFields & "[" & fld.Name & "],"
            Next
            strFields = Left(strFields, Len(strFields) - 1)
            strParentFields = Left(strParentFields, Len(strParentFields) - 1)
        
            intAttr = rel.Attributes
            If intAttr >= dbRelationDeleteCascade Then
                intAttr = intAttr - dbRelationDeleteCascade
                strIntegrity = strIntegrity & "ON DELETE CASCADE "
            End If
            If intAttr >= dbRelationUpdateCascade Then
                intAttr = intAttr - dbRelationUpdateCascade
                strIntegrity = strIntegrity & "ON UPDATE CASCADE "
            End If
            If strIntegrity <> "" Then strIntegrity = vbCrLf & "    " & Left(strIntegrity, Len(strIntegrity) - 1)
            
            strSQL = strSQL & "    FOREIGN KEY (" & strFields & ")" & vbCrLf & _
                              "    REFERENCES [" & rel.Table & "](" & strParentFields & ")" & _
                              strIntegrity & ";" & vbCrLf & "GO" & vbCrLf
        End If
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    Next
    
    GenerateForeignKeySQL = strSQL
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set rel = Nothing
    Set fld = Nothing
End Function

Private Function GenerateUniqueIndexSQL() As String
'PURPOSE: Scripts Unique Indexes from Access Tables for SQL Server
'RETURNS: SQL Server script as a string
    
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strKeys As String
    Dim strKeysForName As String
    Dim strSQL As String
    Dim i As Integer
    
    'Start the meter
    SysCmd acSysCmdInitMeter, "Scripting Unique Indexes", CurrentDb.TableDefs.Count
    
    'Start strSQL
    strSQL = vbCrLf & "/****** UNIQUE INDEXES as of " & Now() & " ******/" & vbCrLf & vbCrLf
    
    'loop through each table in the database
    For Each tdf In CurrentDb.TableDefs
        'Omits system tables
        If tdf.Attributes = 0 Then
            'Cycle through each index
            For Each idx In tdf.Indexes
                'If the index is a unique key
                If idx.Unique And Not idx.Primary Then
                    strKeys = ""
                    For Each fld In idx.Fields
                        strKeys = strKeys & "[" & fld.Name & "],"
                    Next fld
                    strKeys = Left(strKeys, Len(strKeys) - 1)
                    strKeysForName = Replace(strKeys, "[", "")
                    strKeysForName = Replace(strKeysForName, "]", "")
                    strKeysForName = Replace(strKeysForName, ",", "_")
                    strSQL = strSQL & "ALTER TABLE dbo.[" & tdf.Name & "]" & vbCrLf
                    strSQL = strSQL & "    ADD CONSTRAINT UQ_" & Replace(tdf.Name, " ", "_") & "__" & Replace(strKeysForName, " ", "_") & " " & _
                                      "UNIQUE (" & strKeys & ");" & vbCrLf
                    strSQL = strSQL & "GO" & vbCrLf
                End If
            Next idx
        End If
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    Next tdf
    
    GenerateUniqueIndexSQL = strSQL
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set tdf = Nothing
    Set idx = Nothing
    Set fld = Nothing
End Function


Private Function GenerateNonUniqueIndexSQL() As String
'PURPOSE: Scripts Non-Unique Indexes from Access Tables for SQL Server
'RETURNS: SQL Server script as a string
    
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim idxSkipCheck As DAO.Index
    Dim fldSkipCheck As DAO.Field
    Dim strKeys As String
    Dim strKeysSkipCheck As String
    Dim blnSkip As Boolean
    Dim strKeysForName As String
    Dim strSQL As String
    Dim i As Integer
    
    'Start the meter
    SysCmd acSysCmdInitMeter, "Scripting Non-Uniaue Indexes", CurrentDb.TableDefs.Count
    
    'Start strSQL
    strSQL = vbCrLf & "/****** NON-UNIQUE INDEXES as of " & Now() & " ******/" & vbCrLf & vbCrLf
    
    'loop through each table in the database
    For Each tdf In CurrentDb.TableDefs
        'Omits system tables
        If tdf.Attributes = 0 Then
            'Cycle through each index
            For Each idx In tdf.Indexes
                'If the index is not a unique or primary key
                If Not idx.Unique And Not idx.Primary Then
                    strKeys = ""
                    For Each fld In idx.Fields
                        strKeys = strKeys & "[" & fld.Name & "],"
                    Next fld
                    strKeys = Left(strKeys, Len(strKeys) - 1)
                    strKeysForName = Replace(strKeys, "[", "")
                    strKeysForName = Replace(strKeysForName, "]", "")
                    strKeysForName = Replace(strKeysForName, ",", "_")
                    If Not idx.Foreign Then
                        'In Access it looks like it created duplicate indexes for foreign keys sometimes.
                        'If this is a regular index, see if the same fields are already indexed as a foreign key
                        'If so, skip this index
                        blnSkip = False
                        For Each idxSkipCheck In tdf.Indexes
                            'If the index is not a unique or primary key and is a foreign key
                            If Not idxSkipCheck.Unique And Not idxSkipCheck.Primary And idxSkipCheck.Foreign Then
                                strKeysSkipCheck = ""
                                For Each fldSkipCheck In idxSkipCheck.Fields
                                    strKeysSkipCheck = strKeysSkipCheck & "[" & fldSkipCheck.Name & "],"
                                Next fldSkipCheck
                                strKeysSkipCheck = Left(strKeysSkipCheck, Len(strKeysSkipCheck) - 1)
                                'Now check if this is the same as the normal index trying to be created
                                If strKeys = strKeysSkipCheck Then
                                    blnSkip = True
                                    Exit For
                                End If
                            End If
                        Next idxSkipCheck
                    End If
                    
                    'If we're not skipping this
                    If Not blnSkip Then
                        strSQL = strSQL & "CREATE INDEX "
                        If idx.Foreign Then
                            strSQL = strSQL & "FK_"
                        Else
                            strSQL = strSQL & "IX_"
                        End If
                        strSQL = strSQL & Replace(tdf.Name, " ", "_") & "__" & Replace(strKeysForName, " ", "_") & " " & _
                                          "ON [" & tdf.Name & "](" & strKeys & ");" & vbCrLf
                        strSQL = strSQL & "GO" & vbCrLf
                    End If
                End If
            Next idx
        End If
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    Next tdf
    
    GenerateNonUniqueIndexSQL = strSQL
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set tdf = Nothing
    Set idx = Nothing
    Set fld = Nothing
End Function


Private Function GenerateInsertIntoSQL() As String
'PURPOSE: Scripts the data in Access Tables for SQL Server
'RETURNS: SQL Server script as a string
    
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rst As DAO.Recordset
    Dim strInsertInto As String
    Dim strValues As String
    Dim intCount As Long
    Dim strSQL() As String
    Dim intSQL As Long
    Dim intTotalCount As String
    Dim intAttr As Long
    Dim blnAutonum As Boolean
    Dim i As Integer
    
    'loop through each table in the database
    For Each tdf In CurrentDb.TableDefs
        
        blnAutonum = False
        'Omits system tables
        If tdf.Attributes = 0 Then
            
            'Get a recordset of the values
            Set rst = tdf.OpenRecordset
            If Not rst.EOF Then
                
                'Start the table
                ReDim Preserve strSQL(intSQL)
                strSQL(intSQL) = vbCrLf & "/****** [" & tdf.Name & "] DATA as of " & Now() & " ******/" & vbCrLf & vbCrLf
                intSQL = intSQL + 1
        
                'Start the meter
                i = 0
                SysCmd acSysCmdInitMeter, "Scripting " & tdf.Name & " Table Data", rst.RecordCount
                
                strInsertInto = "INSERT INTO dbo.[" & tdf.Name & "] ("
                'cycle through each field
                For Each fld In tdf.Fields
                    'Build the list of fields
                    strInsertInto = strInsertInto & "[" & fld.Name & "],"
                    'Check if this is an autonum field and add the IDENTITY_INSERT line
                    'to allow inserting of values for the autonum field
                    intAttr = fld.Attributes
                    If intAttr >= dbHyperlinkField Then intAttr = intAttr - dbHyperlinkField
                    If intAttr >= dbSystemField Then intAttr = intAttr - dbSystemField
                    If intAttr >= dbUpdatableField Then intAttr = intAttr - dbUpdatableField
                    If intAttr >= dbAutoIncrField Then
                        ReDim Preserve strSQL(intSQL)
                        strSQL(intSQL) = "SET IDENTITY_INSERT dbo.[" & tdf.Name & "] ON;" & vbCrLf & "GO" & vbCrLf
                        intSQL = intSQL + 1
                        blnAutonum = True
                    End If
                Next fld
                strInsertInto = Left(strInsertInto, Len(strInsertInto) - 1) & ")" & vbCrLf
                
                rst.MoveFirst
                intTotalCount = 0
                Do
                    strValues = "VALUES "
                    intCount = 0
                    Do
                        'Add a set of values for a record
                        strValues = strValues & "("
                        For Each fld In tdf.Fields
                            If IsNull(rst(fld.Name)) Then
                                strValues = strValues & "Null,"
                            Else
                                Select Case fld.Type
                                    Case dbByte, dbChar, dbGUID, dbLongBinary, dbMemo, dbText, dbVarBinary
                                        strValues = strValues & "'" & Replace(rst(fld.Name), "'", "''") & "',"
                                    Case dbDate, dbTime, dbTimeStamp
                                        strValues = strValues & "CAST('" & rst(fld.Name) & "' AS DATETIME),"
                                    Case dbBoolean
                                        Select Case CStr(rst(fld.Name))
                                            Case "No", "False", "Off", "0"
                                                strValues = strValues & "0,"
                                            Case "Yes", "True", "On", "1"
                                                strValues = strValues & "1,"
                                        End Select
                                    Case Else
                                        strValues = strValues & rst(fld.Name) & ","
                                End Select
                            End If
                        Next fld
                        strValues = Left(strValues, Len(strValues) - 1) & ")"
                        intCount = intCount + 1
                        rst.MoveNext
                        If intCount = 1000 Or rst.EOF Then
                            strValues = strValues & ";" & vbCrLf
                            intTotalCount = intTotalCount + intCount
                            Exit Do
                        Else
                            strValues = strValues & "," & vbCrLf & "       "
                        End If
                    Loop Until rst.EOF
                    'Add this segment ot the SQL
                    ReDim Preserve strSQL(intSQL)
                    strSQL(intSQL) = strInsertInto & strValues & "GO" & vbCrLf
                    intSQL = intSQL + 1
                    DoEvents
                    i = i + 1
                    SysCmd acSysCmdUpdateMeter, i
                Loop Until rst.EOF
            End If
            If blnAutonum Then
                ReDim Preserve strSQL(intSQL)
                strSQL(intSQL) = "SET IDENTITY_INSERT dbo.[" & tdf.Name & "] OFF;" & vbCrLf & "GO" & vbCrLf
                intSQL = intSQL + 1
            End If
        End If
    Next tdf
    
    GenerateInsertIntoSQL = Join(strSQL, "")
    
    'Remove the Meter
    SysCmd acSysCmdRemoveMeter
    
    Set tdf = Nothing
    Set fld = Nothing
    Set rst = Nothing
End Function


