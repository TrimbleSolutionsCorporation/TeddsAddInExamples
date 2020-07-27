Imports System
Imports System.Data.SqlServerCe
Imports System.Windows.Forms

Imports CSCWorld.Tedds.ExprInterop

Namespace AddIns

    'AddIn functions must be declared in a public Module or class for Tedds to load it
    Public Module DBaseAddIn

        '==================================================================================================================================
        'Subs and Functions must be public to be turned into Tedds functions
        'Within a class Subs and Functions must also be declared with the Shared keyword
        'Valid parameter and return types are void, bool, int, double, string
        'Functions can also accept/return an object (or dynamic) type but the instance passed/returned must be one of the above types
        'If an unrecognised type is returned Tedds will call ToString on the instance and return that
        '==================================================================================================================================
        'Input  : None
        'Return :  True = Database opened 
        '         False = Failed to open database
        '==================================================================================================================================
        Public Function DBaseOpen() As Boolean
            Using dlgOpenFile As New OpenFileDialog()
                If dlgOpenFile.ShowDialog() <> DialogResult.OK Then
                    DBaseOpen = False
                Else
                    DBaseOpen = DBaseOpen(String.Concat("Data Source = ", dlgOpenFile.FileName))
                End If
            End Using
        End Function

        '==================================================================================================================================
        'If multiple functions with the same name exist (with different signatures)
        'Tedds will use the arguments provided to determine which version to call
        '==================================================================================================================================
        'Input  : connectionString = Database connection string
        'Return :             True = Database opened 
        '                    False = Failed to open database
        '==================================================================================================================================
        Public Function DBaseOpen(connectionString As String) As Boolean
            'Close existing connection
            If Not IsNothing(_connection) And Not DBaseClose() Then
                DBaseOpen = False
            Else
                _connection = New SqlCeConnection(connectionString)
                _connection.Open()
                DBaseOpen = True
            End If
        End Function

        '==================================================================================================================================
        'Tedds uses the same rules on function names as the .Net framework (no numeric characters at the start of functions etc)
        'The only additional restriction is that the length of a function name is limited to 32 characters
        'Ensure that any function names declared do not clash with existing Tedds functions
        'Avoid short function names that are likely to be used as variable names by users e.g. Length
        'It is also recommended that all functions are given a matching prefix to indicate that they are all part of the same API
        '==================================================================================================================================
        'Input  : None
        'Return :  True = Closed 
        '         False = Already closed
        '==================================================================================================================================
        Public Function DBaseClose() As Boolean
            If IsNothing(_connection) Then
                DBaseClose = False
            Else
                _connection.Close()
                _connection = Nothing
                DBaseClose = True
            End If
        End Function

        '==================================================================================================================================
        'Input  : command = SQL script
        'Return :  True = Database read successfully
        '         False = Failed to read Database
        '==================================================================================================================================
        Public Function DBaseExecuteReader(command As String) As Boolean
            If IsNothing(_connection) Then
                DBaseExecuteReader = False
            ElseIf Not IsNothing(_reader) And Not DBaseCloseReader() Then     'Close existing reader 
                DBaseExecuteReader = False
            Else
                _reader = New SqlCeCommand(command, _connection).ExecuteReader()
                DBaseExecuteReader = True
            End If
        End Function

        '==================================================================================================================================
        'Use the Requirement attribute to place constraints on input parameters
        'These requirements will be checked by Tedds before the function is called
        'Use the ValidTypes attribute to specify which basic types can be passed to an object parameter
        '==================================================================================================================================
        'Input  : id = Selected database
        'Return : Returns the contents of the selected id
        '==================================================================================================================================
        Public Function DBaseRead(<Requirement("Index must be positive", Requirement.Positive)> <ValidTypes(GetType(Int32), GetType(String))> id As Object) As Object
            If IsNothing(_reader) Or id >= _reader.FieldCount Then
                DBaseRead = Nothing
            Else

                DBaseRead = _reader(id)
            End If

        End Function

        'The Units attribute can be used to inform Tedds to apply the associated dimensions to return values
        '==================================================================================================================================
        'Input  : id = Selected database
        'Return : Returns the contents of the selected id if it's in Meters or length
        '==================================================================================================================================
        Public Function DBaseReadLength(<Requirement("Index must be positive", Requirement.Positive)> id As Int32) As <Units("m")> Double
            DBaseReadLength = DBaseRead(id)

        End Function
        '==================================================================================================================================
        'Input  : id = Selected database
        'Return : Returns the contents of the selected id if it's in Newton or force
        '==================================================================================================================================
        Public Function DBaseReadForce(<Requirement("Index must be positive", Requirement.Positive)> id As Int32) As <Units("N")> Double
            DBaseReadForce = DBaseRead(id)
        End Function

        '==================================================================================================================================
        'The Units attribute can also be used to enforce dimensions on input parameters
        'If a Units attribute is not provided for a numerical parameter Tedds will require that the argument passed be dimensionless
        'The Alias attribute informs Tedds to ADDITIONALLY register a function with Tedds with the given alias
        'e.g. the functions DBaseSet, DBaseSetLength and DBaseSetForce could all be called from within Tedds
        'Because Tedds checks dimensions, functions that would be ambiguous to the .Net compiler (e.g. DBaseSet) can be resolved by Tedds
        '==================================================================================================================================
        'Input  :       table = Table name
        '              column = Column name
        '              length = Value in units of length. default value set to meters
        '         whereColumn = Column Key name
        '          whereValue = Row Id
        'Return :       False = Unable to set length
        '                True = Update successful
        '==================================================================================================================================
        <CSCWorld.Tedds.ExprInterop.Alias("DBaseSet")>
        Public Function DBaseSetLength(table As String, column As String, <Units("m")> length As Double, whereColumn As String, whereValue As Object) As Boolean
            DBaseSetLength = Update(table, column, length, whereColumn, whereValue)
        End Function
        '==================================================================================================================================
        'Input  :       table = Table name
        '              column = Column name
        '              length = Value in units of force. default value set to newtons
        '         whereColumn = Column Key name
        '          whereValue = Row Id
        'Return :       False = Unable to set force
        '                True = Update successful
        '==================================================================================================================================
        <CSCWorld.Tedds.ExprInterop.Alias("DBaseSet")>
        Public Function DBaseSetForce(table As String, column As String, <Units("N")> force As Double, whereColumn As String, whereValue As Object) As Boolean
            DBaseSetForce = Update(table, column, force, whereColumn, whereValue)
        End Function
        '==================================================================================================================================
        'Private helper methods are not exported to Tedds
        '==================================================================================================================================
        'Input  :       table = Table name
        '              column = Column name
        '               value = new Value 
        '         whereColumn = Column Key name
        '          whereValue = Row Id
        'Return :       False = Unable to value
        '                True = Update successful
        '==================================================================================================================================
        Private Function Update(table As String, column As String, value As Double, whereColumn As String, whereValue As Object) As Boolean
            Update = DBaseExecuteNonQuery(String.Format("UPDATE {0} SET {1}='{2}' WHERE {3}='{4}'", table, column, value, whereColumn, whereValue)) > 0
        End Function

#Region "Additional function for completeness"

        '==================================================================================================================================
        'Input  : None
        'Return :  True = Data available
        '         False = Data not available
        '==================================================================================================================================
        Public Function DBaseNextRecord() As Boolean
            If IsNothing(_reader) Then
                DBaseNextRecord = False
            ElseIf (_reader.Read()) Then
                DBaseNextRecord = True
            Else
                'Reader is finished
                DBaseCloseReader()
                DBaseNextRecord = False
            End If

        End Function

        '==================================================================================================================================
        'Input  : None
        'Return :  True = Data reader closed
        '         False = Data reader already closed
        '==================================================================================================================================
        Public Function DBaseCloseReader() As Boolean
            If IsNothing(_reader) Then
                DBaseCloseReader = False
            Else
                _reader.Close()
                _reader = Nothing
                DBaseCloseReader = True
            End If
        End Function

        '==================================================================================================================================
        'Input  : command = SQL Script
        'Return : Integer value indicating if data has been updated or not
        '==================================================================================================================================
        Public Function DBaseExecuteNonQuery(command As String) As Int32
            If IsNothing(_connection) Then
                DBaseExecuteNonQuery = 0
            Else
                DBaseExecuteNonQuery = New SqlCeCommand(command, _connection).ExecuteNonQuery()
            End If
        End Function

        '==================================================================================================================================
        'Input  : command = SQL Script
        'Return : Object containing the result from the query
        '==================================================================================================================================
        Public Function DBaseExecuteScalar(command As String) As Object
            If IsNothing(_connection) Then
                DBaseExecuteScalar = False
            Else
                DBaseExecuteScalar = New SqlCeCommand(command, _connection).ExecuteScalar()
            End If
        End Function


        Private _connection As SqlCeConnection
        Private _reader As SqlCeDataReader

#End Region

    End Module
End Namespace