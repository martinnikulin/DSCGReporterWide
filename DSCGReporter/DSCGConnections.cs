using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;

namespace DSCGReporter
{
    public static class DSCGConnections
    {
        public static bool IntegratedSecurity;
        public static string UserName;
        public static string Password;
        public static string Database;
        public static string ServerAddr;

        public static SqlConnection CatalogConnection;
        public static SqlConnection GDBConnection;

        public static void OpenCatalogConnection()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CatalogConnection"].ConnectionString;

            CatalogConnection = new SqlConnection(connectionString);
            CatalogConnection.Open();
        }

        public static void OpenGDBConnection()
        {
            GDBConnection = new SqlConnection(GetConnectionString());
            GDBConnection.Open();
        }
        private static string GetConnectionString()
        {
            string connString = "";

            if (IntegratedSecurity)
            {
                connString = "Integrated Security=SSPI;Persist Security Info=True;Data Source=" + ServerAddr;
            }
            else
            {
                connString = "Persist Security Info=True;User Id=" + UserName + ";Password=" + Password + ";Data Source=" + ServerAddr;
            }

            connString = connString + ";Initial Catalog=" + Database;

            return connString;
        }
    }
}
