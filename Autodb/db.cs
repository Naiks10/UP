using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Autodb
{
    class db
    {
        public static DataTable currentServers;
        public static SqlConnection connection;
        public static string userId;
        public static string serverName;
        public static DateTime connectionStartTime;

        public static DataTable getServers()
        {
            SqlDataSourceEnumerator instance = SqlDataSourceEnumerator.Instance;
            return instance.GetDataSources();
        }

        public static async void fillServers()
        {
            //Вызвать этот метод и в currentServers появятся доступные сервера
            currentServers = await Task.Run(() => getServers());
        }

        public static void startConnection(String connectionString)
        {
            connection = new SqlConnection(connectionString);
            connection.Open();
        }
    }
}
