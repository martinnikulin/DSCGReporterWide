using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSCGReporter
{
    public static class DataRepo
    {
        public static int GdbType;
        public static int ReportType;
        public static int InterbedId;
        public static int VersionId;

        public static DataTable Variants;
        public static List<string> Categories;
        public static List<DataTable> SubVariants;
        public static DataTable Reserves;

        public static int thCount;
        public static int ashCount;
        public static int SumCatIndex;

        internal static void OpenTables()
        {
            Variants = GetVariants();
            Categories = GetCategories();
            SubVariants = GetSubVariants();

            thCount = Variants.AsEnumerable().Select(row => new { Thnn = row.Field<int>("Thnn") }).Distinct().Count();
            ashCount = Variants.AsEnumerable().Select(row => new { Ashnn = row.Field<int>("Ashnn") }).Distinct().Count();

        }

        public static DataTable GetProjects()
        {
            return GetTable(DSCGConnections.CatalogConnection, "select p.Id, ProjectName, [Database], p.Actual, s.LocalServer, s.RemoteServer as [Server] from Projects p inner join Servers s on s.Id = p.ServerId");
        }

        private static DataTable GetVariants()
        {
            string sql;
            if (GdbType == 1 & ReportType == 1)
            {
                sql = "select NVariant, InterbedId, Ibnn, Thnn, Ashnn, IbCond, ThCond, AshCond, VariantName1 as VariantName from Variants where InterbedId = " + InterbedId.ToString();
            }
            else
            {
                sql = "select NVariant, InterbedId, Ibnn, Thnn, Ashnn, IbCond, ThCondSum as ThCond, AshCondSum as AshCond, VariantName2 as VariantName from Variants where InterbedId = " + InterbedId.ToString();
            }

            return GetTable(DSCGConnections.GDBConnection, sql);
        }

        private static List<DataTable> GetSubVariants()
        {
            string sql;

            sql = "select Variant, SubVariant from fVariants(" + ReportType.ToString() + ", " + InterbedId.ToString() + ") order by Variant, SubVariant";
            DataTable subVariants = GetTable(DSCGConnections.GDBConnection, sql);

            List<DataTable> subVariantList = new List<DataTable>();

            foreach (DataRow row in Variants.AsEnumerable())
            {
                int variant = row.Field<int>("NVariant");

                var results = from subVariant in subVariants.AsEnumerable()
                              where subVariant.Field<int>("Variant") == variant
                              select subVariant;

                DataTable vs = results.CopyToDataTable();

                subVariantList.Add(vs);
            }
            return subVariantList;
        }

        public static DataTable GetInterbeds()
        {
            return GetTable(DSCGConnections.GDBConnection, "select distinct InterbedId, Ibcond from Variants");
        }

        private static List<string> GetCategories()
        {
            List<string> categories = new List<string>();

            DataTable cats = GetTable(DSCGConnections.GDBConnection, "select distinct d.Description as Category from Blocks b inner join Dictionary d on d.DictionaryId = b.Category order by Category");
            
            string abc = "";
            int catnumber = 0;

            foreach (var cat in cats.AsEnumerable())
            {
                string catname = cat.Field<string>("category");
                if (catname.CompareTo("C2") < 0) {
                    categories.Add(catname);
                    abc = abc + " + " + catname;
                    catnumber++;
                }
            }
            SumCatIndex = catnumber;

            if (catnumber > 1)
                categories.Add(abc.Substring(3));

            foreach (var cat in cats.AsEnumerable())
            {
                string catname = cat.Field<string>("category");
                if (catname.CompareTo("C1") > 0)
                    categories.Add(catname);
            }

            return categories;
        }
        public static DataTable GetReserves()
        {
            SqlCommand spCommand = new SqlCommand("Reserves", DSCGConnections.GDBConnection);
            spCommand.CommandType = CommandType.StoredProcedure;
            spCommand.Parameters.AddWithValue("@GdbType", GdbType);
            spCommand.Parameters.AddWithValue("@InterbedId", InterbedId);
            spCommand.Parameters.AddWithValue("@VersionId", VersionId);

            SqlDataAdapter adapter = new SqlDataAdapter(spCommand);

            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet);

            return dataSet.Tables[0];
        }

        private static DataTable GetTable(SqlConnection connection, string sql)
        {
            SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet);
            
            return dataSet.Tables[0];
        }
    }
}
