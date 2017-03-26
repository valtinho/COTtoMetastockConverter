using System;
using System.Data;
using System.Data.OleDb;

namespace COTtoMetastockConverter
{
    public static class ExcelHelpers
    {
        private const string _oleDbConnectionString = "Provider=\"Microsoft.Jet.OLEDB.4.0\";Data Source=\"$FILEPATH$\";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";";
        private const string _alreadyOpenError = "ERROR. Could not open the worksheet.\nPlease make sure that the file is not already open with Excel, or select a different file.";

        private static string findFirstSheetName(string sourceFilePath)
        {
            string connString = _oleDbConnectionString.Replace("$FILEPATH$", sourceFilePath);
            using (var conn = new OleDbConnection(connString))
            {
                try
                {
                    conn.Open();
                    //Get the data table containg the schema guid.
                    using (DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                    {
                        if (dt == null) return null;
                        //GetOleDbSchemaTable returns sheets in reverse order
                        //so the first sheet from left to right appears last
                        //other objects are also returned, but only sheet names end with $
                        int i = 0;
                        //default to XLS
                        string lastSheetName = "XLS";
                        foreach (DataRow row in dt.Rows)
                        {
                            string name = row["TABLE_NAME"].ToString();
                            if (name.EndsWith("$"))
                            {
                                lastSheetName = name;
                            }
                            i++;
                        }
                        return lastSheetName;
                    }
                }
                catch
                {
                    ErrorHelpers.deferredEx(_alreadyOpenError);
                }
                return null;
            }
        }

        public static DataTable convertXLSToDataTable(string sourceFilePath)
        {
            string strConn = String.Concat(_oleDbConnectionString.Replace("$FILEPATH$", sourceFilePath));
            using (var conn = new OleDbConnection(strConn))
            {
                try
                {
                    conn.Open();

                    string sheetName = findFirstSheetName(sourceFilePath);
                    using (var cmd = new OleDbCommand(String.Concat("SELECT * FROM [", sheetName, "]"), conn))
                    {
                        //select the data from the spreadsheet and convert it to a DataTable
                        cmd.CommandType = CommandType.Text;
                        using (var da = new OleDbDataAdapter(cmd))
                        {
                            using (var dt = new DataTable())
                            {
                                try
                                {
                                    da.Fill(dt);
                                    return dt;
                                }
                                catch (Exception)
                                {
                                    ErrorHelpers.immediateEx(String.Format("ERROR. Could not find data in a worksheet named {0}.\nPlease select a different file.", sheetName));
                                }
                            }
                        }
                    }
                }
                catch {
                    ErrorHelpers.immediateEx(_alreadyOpenError);
                }
            }
            return null;
        }

    }
}
