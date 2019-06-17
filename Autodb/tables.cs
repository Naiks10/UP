using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;
 
namespace Autodb
{
    public class _Tables
    {
        public static DataTable table(string command)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, db.connection);
            DataTable _dataTable = new DataTable();
            sqlDataAdapter.Fill(_dataTable);
            return _dataTable;
        }
    }
}
