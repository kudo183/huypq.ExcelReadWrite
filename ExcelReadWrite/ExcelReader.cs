using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadWrite
{
    public static class ExcelReader
    {
        public static List<List<object>> Read(string filePath)
        {
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            String strNewPath = filePath;
            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strNewPath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String query = "SELECT * FROM [Sheet1$]"; // You can use any different queries to get the data from the excel sheet
            OleDbConnection conn = new OleDbConnection(connString);
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                cmd = new OleDbCommand(query, conn);
                da = new OleDbDataAdapter(cmd);
                da.Fill(ds);

                var result = new List<List<object>>();
                var rowData = new List<object>();

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    rowData.Clear();
                    foreach (var cell in row.ItemArray)
                    {
                        rowData.Add(cell);
                    }
                    result.Add(rowData);
                }

                return result;
            }
            catch(Exception ex)
            {
                // Exception Msg 
                throw ex;
            }
            finally
            {
                da.Dispose();
                conn.Close();
            }
        }
    }
}
