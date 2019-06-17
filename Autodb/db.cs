using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        public static async void fillServers(ComboBox cb)
        {
            //Вызвать этот метод и в currentServers появятся доступные сервера
            currentServers = await Task.Run(() => getServers());
            try
            {
                foreach (DataRow row in currentServers.Rows)
                {
                    if (row["InstanceName"].ToString() == "")
                        cb.Items.Add(row["ServerName"]);
                    else
                        cb.Items.Add(row["ServerName"] + "\\" + row["InstanceName"]);
                }
            }
            catch (NullReferenceException)
            {
                Application.Exit();
            }
        }

        public static void startConnection(String connectionString)
        {
            connection = new SqlConnection(connectionString);
            connection.Open();
        }
    }
}
