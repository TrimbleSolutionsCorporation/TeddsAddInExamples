using System;
using System.Data.SqlServerCe;
using System.Windows.Forms;

using Tekla.Structural.ExpressoAddIn;

namespace AddIns
{
    //AddIn class must be public for Tedds to load it
    public static class DBaseAddIn
    {
        //==================================================================================================================================
        //Methods must be public static to be turned into Tedds functions
        //Valid parameter and return types are void, bool, int, double, string
        //Functions can also accept/return an object (or dynamic) type but the instance passed/returned must be one of the above types
        //If an unrecognised type is returned Tedds will call ToString on the instance and return that
        //==================================================================================================================================
        //Input  : None
        //Return :  True = Database opened 
        //         False = Failed to open database
        //==================================================================================================================================
        public static bool DBaseOpen()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return false;

                return DBaseOpen(string.Concat("Data Source = ", dialog.FileName));
            }
        }
        
        //==================================================================================================================================
        //If multiple functions with the same name exist (with different signatures)
        //Tedds will use the arguments provided to determine which version to call
        //==================================================================================================================================
        //Input  : connectionString = Database connection string
        //Return :             True = Database opened 
        //                    False = Failed to open database
        //==================================================================================================================================
        public static bool DBaseOpen(string connectionString)
        {
            //Close existing connection
            if (_connection != null && !DBaseClose())
                return false;

            _connection = new SqlCeConnection(connectionString);
            _connection.Open();
            return true;
        }
        
        //==================================================================================================================================
        //Tedds uses the same rules on function names as the .Net framework (no numeric characters at the start of functions etc)
        //The only additional restriction is that the length of a function name is limited to 32 characters
        //Ensure that any function names declared do not clash with existing Tedds functions
        //Avoid short function names that are likely to be used as variable names by users e.g. Length
        //It is also recommended that all functions are given a matching prefix to indicate that they are all part of the same API
        //==================================================================================================================================
        //Input  : None
        //Return :  True = Closed 
        //         False = Already closed
        //==================================================================================================================================
        public static bool DBaseClose()
        {
            if (_connection == null)
                return false;
            
            _connection.Close();
            _connection = null;
            return true;
        }
        //==================================================================================================================================
        //Input  : command = SQL script
        //Return :  True = Database read successfully
        //         False = Failed to read Database
        //==================================================================================================================================
        public static bool DBaseExecuteReader(string command)
        {
            if (_connection == null)
                return false;

            //Close existing reader 
            if (_reader != null && !DBaseCloseReader())
                return false;

            _reader = new SqlCeCommand(command, _connection).ExecuteReader();
            return true;
        }

        //==================================================================================================================================
        //Use the Requirement attribute to place constraints on input parameters
        //These requirements will be checked by Tedds before the function is called
        //Use the ValidTypes attribute to specify which basic types can be passed to a dynamic parameter
        //==================================================================================================================================
        //Input  : id = Selected database
        //Return : Returns the contents of the selected id
        //==================================================================================================================================
        public static object DBaseRead([Requirement("Index must be positive", Requirement.Positive)][ValidTypes(typeof(int), typeof(string))] dynamic id)
        {
            if (_reader == null || id >= _reader.FieldCount)
                return null;
            
            return _reader[id];
        }
        
        //==================================================================================================================================
        //The Units attribute can be used to inform Tedds to apply the associated dimensions to return values
        //==================================================================================================================================
        //Input  : id = Selected database
        //Return : Returns the contents of the selected id if it's in Meters or length
        //==================================================================================================================================
        [return: Units("m")]
        public static double DBaseReadLength([Requirement("Index must be positive", Requirement.Positive)] int id)
        {
            return (double)DBaseRead(id);
        }
        //==================================================================================================================================
        //Input  : id = Selected database
        //Return : Returns the contents of the selected id if it's in Newton or force
        //==================================================================================================================================
        [return: Units("N")]
        public static double DBaseReadForce([Requirement("Index must be positive", Requirement.Positive)] int id)
        {
            return (double)DBaseRead(id);
        }

        //==================================================================================================================================
        //The Units attribute can also be used to enforce dimensions on input parameters
        //If a Units attribute is not provided for a numerical parameter Tedds will require that the argument passed be dimensionless
        //The Alias attribute informs Tedds to ADDITIONALLY register a function with Tedds with the given alias
        //e.g. the functions DBaseSet, DBaseSetLength and DBaseSetForce could all be called from within Tedds
        //Because Tedds checks dimensions, functions that would be ambiguous to the .Net compiler (e.g. DBaseSet) can be resolved by Tedds
        //==================================================================================================================================
        //Input  :       table = Table name
        //              column = Column name
        //              length = Value in units of length. default value set to meters
        //         whereColumn = Column Key name
        //          whereValue = Row Id
        //Return :       False = Unable to set length
        //                True = Update successful
        //==================================================================================================================================
        [Alias("DBaseSet")]
        public static bool DBaseSetLength(string table, string column, [Units("m")] double length, string whereColumn, dynamic whereValue)
        {
            return Update(table, column, length, whereColumn, whereValue);
        }
        //==================================================================================================================================
        //Input  :       table = Table name
        //              column = Column name
        //              length = Value in units of force. default value set to newtons
        //         whereColumn = Column Key name
        //          whereValue = Row Id
        //Return :       False = Unable to set force
        //                True = Update successful
        //==================================================================================================================================
        [Alias("DBaseSet")]
        public static bool DBaseSetForce(string table, string column, [Units("N")] double force, string whereColumn, dynamic whereValue)
        {
            return Update(table, column, force, whereColumn, whereValue);
        }
        
        //==================================================================================================================================
        //Private helper methods are not exported to Tedds
        //==================================================================================================================================
        //Input  :       table = Table name
        //              column = Column name
        //               value = new Value 
        //         whereColumn = Column Key name
        //          whereValue = Row Id
        //Return :       False = Unable to value
        //                True = Update successful
        //==================================================================================================================================
        private static bool Update(string table, string column, double value, string whereColumn, dynamic whereValue)
        {
            return DBaseExecuteNonQuery(string.Format("UPDATE {0} SET {1}='{2}' WHERE {3}='{4}'", table, column, value, whereColumn, whereValue)) > 0;
        }

        #region "Additional functions for completeness"

        //==================================================================================================================================
        //Input  : None
        //Return :  True = Data available
        //         False = Data not available
        //==================================================================================================================================
        public static bool DBaseNextRecord()
        {
            if (_reader == null)
                return false;

            if (_reader.Read())
                return true;

            //Reader is finished
            DBaseCloseReader();
            return false;
        }

        //==================================================================================================================================
        //Input  : None
        //Return :  True = Data reader closed
        //         False = Data reader already closed
        //==================================================================================================================================
        public static bool DBaseCloseReader()
        {
            if (_reader == null)
                return false;

            _reader.Close();
            _reader = null;
            return true;
        }

        //==================================================================================================================================
        //Input  : command = SQL Script
        //Return : Integer value indicating if data has been updated or not
        //==================================================================================================================================
        public static int DBaseExecuteNonQuery(string command)
        {
            if (_connection == null)
                return 0;

            return new SqlCeCommand(command, _connection).ExecuteNonQuery();
        }
                
        //==================================================================================================================================
        //Input  : command = SQL Script
        //Return : Object containing the result from the query
        //==================================================================================================================================
        public static object DBaseExecuteScalar(string command)
        {
            if (_connection == null)
                return false;

            return new SqlCeCommand(command, _connection).ExecuteScalar();
        }

        private static SqlCeConnection _connection;
        private static SqlCeDataReader _reader;

        #endregion
    }
}
