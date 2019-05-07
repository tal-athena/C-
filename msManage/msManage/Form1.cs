using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Controls;
using System.Windows.Forms;

namespace msManage
{
    public partial class MsForm : Form
    {
        private List<string> lsFileNames;
        private string strTableName;
        private string strIterColumnName;

        public MsForm()
        {
            InitializeComponent();      
        }
        
        // Open files and Get available Tables which are exist in each files
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.Multiselect = true;
            fileDialog.Filter = "access files (*.accdb)|*.accdb";
            
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                lsFileNames = new List<string> (fileDialog.FileNames);
            }
            else
            {
                return;
            }

            // initalize
            listTables.Items.Clear();
            gridColumns.Rows.Clear();
            outputPane.Clear();

            List<string> strTableNames = getAvailableTables();

                        
            // add table names to listview
            foreach (string table in strTableNames)
            {                 
                listTables.Items.Add(table);
            }
            
        }

        //Merge button clicked
        private void button2_Click(object sender, EventArgs e)
        {
            //Check if columns is selected
            if (gridColumns.SelectedCells.Count != 1)
            {
                MessageBox.Show("Please select a column", "Alert");
                return;
            }

            strIterColumnName = gridColumns.SelectedCells[0].Value.ToString();

            outputPane.Clear();

            //then Merge tables to one file
            mergeToOneFile();

        }
        private List<string> getAvailableTables()
        {
            Dictionary<string, int> map = new Dictionary<string, int>();    //pair of table name and table count

            //check if the table exist in each file
            //count of table apperance and check is it same as file count
            for (int i = 0; i < lsFileNames.Count; i ++)
            {
                // in every file find tables and incresize the count of table names
                MyAccessDBController fileController = new MyAccessDBController(lsFileNames[i]);
                List<string> tableList = fileController.getAllTables();
                
                foreach (string tableName in tableList)
                {
                    if (!map.ContainsKey(tableName))
                        map[tableName] = 1;
                    else 
                        map[tableName] ++;
                }

                fileController.closeConnection();
            }

            List<string> strAvailableTableNames = new List<string>();

            foreach (String key in map.Keys)
            {
                // if table count and file count is same it's the available table
                if (map[key] == lsFileNames.Count)
                    strAvailableTableNames.Add(key);
                    
            }

            return strAvailableTableNames;
        }

        // Find all columns of selected table in each files and show it with DataGridView
        private void getMultipleColumns()
        {
            Dictionary<string, List<string>> map = new Dictionary<string, List<string>>(); // pair of column name and file list
            
            //in each file
            foreach (string fileName in lsFileNames)
            {
                MyAccessDBController dbController = new MyAccessDBController(fileName);

                List<string> columnList = dbController.getAllColumns(strTableName); // get column list

                // then add to map to get sorted list
                foreach (string column in columnList)
                {
                    if (!map.ContainsKey(column))
                    {
                        map[column] = new List<string>();
                    }
                    map[column].Add(fileName);
                }

                dbController.closeConnection();
            }

            
            //put the list (column name , file name ) to DataGridView-girdColumns
            foreach (string column in map.Keys)
            {
                List<string> fileNames = map[column];

                if (fileNames.Count != lsFileNames.Count) continue;

                foreach (string fileName in fileNames)
                {
                    object[] row = { false, column, fileName };

                    gridColumns.Rows.Add(row);

                    break;
                }
            }
        }
        private void mergeToOneFile()           // function to merge to one file
        {
            string destFile = lsFileNames[0];   
            destFile.Remove(destFile.IndexOf('.'));

            destFile += "_merged.accdb";        //get destination file name

            System.IO.File.Copy(lsFileNames[0], destFile, true);

            //get db connection for destination file
            MyAccessDBController destDbController = new MyAccessDBController(destFile);

            //clear all tables except selected table
            destDbController.dropAllTablesExceptSelectedTable(strTableName);

            List<string> mergedColumns = new List<string>();

            Dictionary<string, OleDbType> columnDataTypes = new Dictionary<string, OleDbType>();

            // get necessary columns which will be merged
            for (int i = 0; i < gridColumns.Rows.Count; i++)
            {
                //get column type
                string column = gridColumns.Rows[i].Cells[1].Value.ToString();

                mergedColumns.Add(column);
                columnDataTypes[column] = destDbController.getColumnType(column, strTableName);
            }


            //delete all columns except multiple columns
            destDbController.dropNotNecessaryColumns(mergedColumns, strTableName);

            //clear all records
            destDbController.makeEmptyTable(strTableName);


            outputPane.AppendText("Destination File : " + destFile + Environment.NewLine + Environment.NewLine);

            int totalLines = 0;

            // merge records one by one
            foreach (string fileName in lsFileNames)
            {
                int copiedLines = 0, skippedLines = 0;                

                MyAccessDBController dbController = new MyAccessDBController(fileName);

                OleDbDataReader reader = dbController.getRecordReader(mergedColumns, strTableName);
                                
                while (reader.Read())
                {
                    List<KeyValuePair<string, object>> record = new List<KeyValuePair<string, object>>();

                    int iterIndex = 0;
                    foreach (string column in mergedColumns)
                    {
                        record.Add(new KeyValuePair<string, object>(column, reader[column]));    // add to value list
                    }
                    for (iterIndex = 0; iterIndex < record.Count; iterIndex++)
                        if (record[iterIndex].Key == strIterColumnName)
                            break;

                    string quato = "";                    

                    switch (columnDataTypes[record[iterIndex].Key])
                    {
                        case OleDbType.Char:
                        case OleDbType.Binary:
                        case OleDbType.WChar:
                        case OleDbType.VarChar:
                        case OleDbType.LongVarChar:
                        case OleDbType.VarWChar:
                        case OleDbType.LongVarWChar:
                        case OleDbType.VarBinary:
                        case OleDbType.LongVarBinary:
                            quato = "\'";
                            break;
                    }

                    if (record[iterIndex].Value.ToString() == "") continue;
                    
                    if (!destDbController.isExistField(record[iterIndex].Value.ToString(), strIterColumnName, strTableName, quato))
                    {
                        destDbController.insertRecord(record, strTableName);
                        copiedLines++;
                    }
                    else
                    {
                        skippedLines++;
                    }

                }
                dbController.closeConnection();

                /*
                List<List<KeyValuePair<string, object>>> allRecords = dbController.getAllRecords(mergedColumns, strTableName);

                for (int i = 0; i < allRecords.Count; i ++)
                {
                    List<KeyValuePair<string, object>> record = allRecords[i];

                    int j;

                    for (j = 0; j < record.Count; j ++)
                    {
                        if (record[j].Key == strIterColumnName)
                        {
                            break;
                        }
                    }

                    string quato = "";
                    OleDbType type = destDbController.getColumnType(record[j].Key, strTableName);
                    
                    switch (type)
                    {
                        case OleDbType.Char:
                        case OleDbType.Binary:
                        case OleDbType.WChar:
                        case OleDbType.VarChar:
                        case OleDbType.LongVarChar:
                        case OleDbType.VarWChar:
                        case OleDbType.LongVarWChar:
                        case OleDbType.VarBinary:
                        case OleDbType.LongVarBinary:
                            quato = "\'";
                            break;
                    }

                    if (record[j].Value.ToString() != "" && !destDbController.isExistField(record[j].Value.ToString(), strIterColumnName, strTableName, quato))
                    {
                        destDbController.insertRecord(record, strTableName);
                        copiedLines++;
                    } else
                    {
                        skippedLines ++;
                    }
                }
                */

                outputPane.AppendText(copiedLines.ToString() + " lines are copied from " + fileName + Environment.NewLine);
                outputPane.AppendText(skippedLines.ToString() + " lines are skipped from " + fileName + Environment.NewLine);

                outputPane.AppendText(Environment.NewLine);

                totalLines += copiedLines;
            }

            outputPane.AppendText(Environment.NewLine + "Total Lines: " + totalLines.ToString() + Environment.NewLine);

            destDbController.closeConnection();

            MessageBox.Show("Merge Completed");
        }

        // user select table name and then get the columns in each file
        private void listTables_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (listTables.SelectedItems.Count > 0)
            {
                strTableName = listTables.SelectedItem.ToString();
                
                gridColumns.Rows.Clear();

                getMultipleColumns();
            }
        }

    }
    public partial class MyAccessDBController           //Access database helper
    {
        private string strFileName;     
        private OleDbConnection dbConnection;
        public string StrFileName { get => strFileName; set => strFileName = value; }

        public MyAccessDBController(string fileName)
        {
            strFileName = fileName;
            dbConnection = ConnectToDB();

            try
            {
                dbConnection.Open();
            } catch (OleDbException e)
            {
                Console.WriteLine("DB Error");                
            }
        }
        public OleDbConnection ConnectToDB()
        {
            // get ole db connection with Microsoft.ACE.OLEDB.12.0 driver
            return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                + strFileName
                + ";Persist Security Info=True");
        }

        public List<string> getAllTables()
        {               
            // to get all tables, use schema
            DataTable schema = dbConnection.GetSchema("Tables");

            List<string> tableNames = new List<string>();

            foreach (DataRow row in schema.Rows)
            {
                string tableName = row.Field<string>("TABLE_NAME");

                if (tableName.Contains("MSys"))     //system table
                    continue;

                tableNames.Add(tableName);
            }

            // return all table names exist in file, including MSysAccessStorage, MSysACEs, ..., and user made tables
            return tableNames;           
        }
        public List<string> getAllColumns(string tableName)
        {
           //open db connection
            
            List<string> columnList = new List<string>();

            //get all columns in table

            DataTable dbSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tableName, null });
            foreach (DataRow row in dbSchema.Rows)
            {
                columnList.Add(row["COLUMN_NAME"].ToString());
            }

            return columnList;
        }
        public void dropAllTablesExceptSelectedTable(string strTableName)
        {
               // get all tables which will be deleted           
            List<string> tableNames = getAllTables();
                        
            // delete every table in this file
            foreach (string table in tableNames)
            {
                if (table == strTableName)
                    continue;
                try
                {
                    //string sql = "DELETE FROM " + table;
                    //OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
                    string sql = "DROP TABLE " + table;
                    OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
                    cmd.ExecuteNonQuery();
                }
                catch (OleDbException)
                {
                    continue;
                }
            }

        }

        public OleDbType getColumnType(string columnName, string strTableName)
        {
            //get column datatype

            DataTable dbSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, strTableName, null });
            foreach (DataRow row in dbSchema.Rows)
            {
                if (row["COLUMN_NAME"].ToString() == columnName)
                {
                    return (OleDbType)row["DATA_TYPE"];
                }
                    
                   
            }
            return 0;            
        }
        public OleDbDataReader getRecordReader(List<string> cols, string tableName)
        {
            //create sql query
            string sql = "select";

            bool flag = false;

            foreach (string column in cols)
            {
                if (flag == false)
                {
                    sql += " [" + column + "]";
                    flag = true;
                }
                else sql += ", [" + column + "]";
            }

            sql += " from " + tableName;

            //execute query and get column values

            OleDbDataReader reader = null;
            OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
            reader = cmd.ExecuteReader();

            return reader;
        }
        public List<List<KeyValuePair<string, object>>> getAllRecords(List<string> cols, string tableName)
        {
            //create sql query
            string sql = "select";

            bool flag = false;

            foreach (string column in cols)
            {
                if (flag == false)
                {
                    sql += " [" + column + "]";
                    flag = true;
                }
                else sql += ", [" + column + "]";                
            }

            sql += " from " + tableName;

            //execute query and get column values

            OleDbDataReader reader = null;
            OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
            reader = cmd.ExecuteReader();

            //All record data
            List<List<KeyValuePair<string, object>>> records = new List<List<KeyValuePair<string, object>>>();

            int i = 0;

            while (reader.Read())
            {
                records.Add(new List<KeyValuePair<string, object>>());
                foreach(string column in cols)
                {   
                    records[i].Add(new KeyValuePair<string, object>(column, reader[column]));    // add to value list
                }
                i++;                    
            }

            return records;           
        }

        public void insertRecord(List<KeyValuePair<string, object>> record, string tableName)
        {
            bool flag = false;

            //create sql query

            string sql = "INSERT INTO " + tableName + "(";
            string values = "VALUES (";


            foreach (KeyValuePair<string, object> field in record)
            {
                if (field.Value.ToString() == "")
                    continue;

                string value = field.Value.ToString().Replace("'", "''");

                if (flag == false)
                {
                    sql += "[" + field.Key + "]";
                    values += "\'" + value + "\'";
                    flag = true;
                }
                else
                {
                    sql += ", " + "[" + field.Key + "]";
                    values += ", \'" + value + "\'";
                }
            }
            sql += ")" + " " + values + ")";

            //execute sql query

            OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
            cmd.ExecuteNonQuery();
            
        }

        public bool isExistField(string v, string strIterColumnName, string strTableName, string quato = "")
        {
            //create sql query
            v = v.Replace("'", "''");
            string sql = "SELECT [" + strIterColumnName + "] FROM " + strTableName + " WHERE [" + strIterColumnName + "]=" + quato + v + quato;

            //execute sql query

            OleDbDataReader reader;
            OleDbCommand cmd = new OleDbCommand(sql, dbConnection);            
            reader = cmd.ExecuteReader();

            return reader.HasRows;
        }

        public void dropNotNecessaryColumns(List<string> mergedColumns, string tableName)
        {

            List<string> allColumns = getAllColumns(tableName);

            foreach (string column in mergedColumns)
            {
                allColumns.Remove(column);                
            }
            
            foreach (string column in allColumns)
            {
                string sql = "ALTER TABLE " + tableName + " DROP COLUMN " + "[" + column + "]";
                OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
                cmd.ExecuteNonQuery();
            }            
        }
        public void makeEmptyTable(string tableName)
        {
            
            string sql = "DELETE FROM " + tableName;

            OleDbCommand cmd = new OleDbCommand(sql, dbConnection);
            cmd.ExecuteNonQuery();
            
        }
        public void closeConnection()
        {
            dbConnection.Close();
        }
        public bool openConnection()
        {
            try
            {
                dbConnection.Open();
            }
            catch (OleDbException e)
            {
                Console.WriteLine("DB Error");
                return false;
            }
            return true;
        }
    }    
}
