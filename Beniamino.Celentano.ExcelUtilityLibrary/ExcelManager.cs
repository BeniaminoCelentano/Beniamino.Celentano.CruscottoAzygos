using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beniamino.Celentano.ExcelUtilityLibrary
{
    public class ExcelManager
    {
        private string _fileName;
        private string _connectionString;

        public ExcelManager(string fileName, bool containsHeader)
        {
            this._fileName = fileName;
            string headerConn = containsHeader ? "HDR=YES;" : "HDR=NO;";
            _connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileName + ";Extended Properties=Excel 12.0;" + headerConn;
        }

        public DataTable GetSheetDataAsDataTable(string sheetName)
        {
            DataTable sheetData = new DataTable();
            using (OleDbConnection conn = this.returnConnection())
            {
                conn.Open();
                // retrieve the data using data adapter
                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("select * from [" + sheetName + "]", conn);
                sheetAdapter.Fill(sheetData);
            }
            return sheetData;
        }

        private OleDbConnection returnConnection()
        {
            return new OleDbConnection(_connectionString);
        }
    }
}
