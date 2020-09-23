Attribute VB_Name = "DBModule"

Public Sub CreateNewDB()

    'Step 1 : Dimensioning main database elements and strings
Dim NewDB As Database
Dim NewTable As TableDef
Dim DBName As String

    'Step 2 : New database file name
    'Name is optional string and the extention is too.
    'mdb extention tells it is an Access file
    'but you can change it what you want
    'for example i use .swt extention sometimes
        
    '(Put a line break on the following line before running)
        DBName = App.Path + "\data.mdb"
    'Kill if the file already exist
        If Dir(DBName) <> "" Then
            Kill DBName
        End If
    
    'Step 3 : Create the new database file
    'Lets say, the database file is a big house includes rooms

        Set NewDB = CreateDatabase(DBName, dbLangGeneral)
    'Second parameter shows locale for international settings
    
    'Now you created the database file
    'You can add the table which will store your data
    'Lets say the tables are the rooms in your big house
    
    'Step 4 : Adding a table
        Set NewTable = NewDB.CreateTableDef("Table1")
                
    'Step 5 : Adding the fields
    'Lets say the fields is the shelf on your room's walls
    With NewTable
        .Fields.Append .CreateField("Field1", dbInteger)
        .Fields.Append .CreateField("Field2", dbText)
        .Fields.Append .CreateField("Field3", dbDate)
    End With
    
    'Step 6 : Append the table into database file
        NewDB.TableDefs.Append NewTable
        
    'Now you have a database with one table and three fields
    'If you want to add new tables you have to go step 4
    'change the table name and follow to the steps 5-6
    
    'This is the procedure to create new database file
    'This procedure is very basic for learning
    'Actually this is very complex procedure
    'For example you can create a password for your DB
    'But you have to get basicly if you want to learn.
    
    'If you want to store data in this file go to
    'StoreData procedure which i created for you.
End Sub

Public Sub Main()
    Stop
    'Please ReadThis

'!!!WARNING!!!
'First you have to add Microsoft DAO Object Library
'using by References Dialog Box
'If not VB doesnot know the database elements
    
    
    'This is a learning program for beginners
    'Please dont response if you have advanced skill in DB programming
    
    'CreateNewDB creates a new database file-Ready
    'StoreData stores data into your database-Ready
    'QueryData queries your database-Not Ready
    'DeleteData deletes a specified record in a table-Not Ready
    'EditData edits a specified record in a table-Not Ready
    
    'If you decide to follow my codes i am going to send
    'them regularly and you will learn everything about
    'database programming.
    'Do not think that i am thinking about that i am an expert
    'I just try to be a good teacher.
    
    'You can test the codes first running project
    'putting the breakpoints to the header of the procedures.
    'To run the procedure write the name of the Sub into the
    'immediate window then press Enter.
    'It is the best way to understand the whole code.
End Sub


Public Sub StoreData()
'Here you learn to storing your data into the new database
'which you created in CreateNewDB procedure
    
    'Step 1 : Dimensioning main database elements and strings
Dim MainDB As Database
Dim table As Recordset
Dim DBName As String
    'Step 2 : Database file name
    'This name must be the same name with which is the
    'database created on CreateNewDB procedure
        
    '(Put a line break on the following line before running)
        DBName = App.Path + "\data.mdb"
    'Error if the file already exist
        If Dir(DBName) = "" Then
            response = MsgBox("Please run CreateNewDB first", vbOKOnly + vbExclamation, "Error")
            Exit Sub
        End If
    'Step 3 : Open existing database
        Set MainDB = OpenDatabase(DBName)
    'Step 4 : Set the TableDef
        Set table = MainDB.OpenRecordset("Table1")
    'Step 5 : Put the data you want to store
        table.AddNew
        
        table.Fields("Field1").Value = 100
        table.Fields("Field2").Value = "sample"
        table.Fields("Field3").Value = 40000
    
    'Step 6 : Update the database
        table.Update

    'table.fields("Fields1")=table.fields(0)
    'table.fields("Fields2")=table.fields(1)
    'table.fields("Fields3")=table.fields(2)
    'Fields have indexes
    'You can reach the fields by index numbers
    'First index number is 0 and it is the
    'firstly created field.
    'That's all for now.
End Sub
