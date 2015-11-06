using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace AutTest
{
    class Reader
    {


        private DataTable dt;

        public DataTable Dt 
        {
            get { return dt; }
        }

       public Reader(String connString)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connString;
            conn.Open();

            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandText = "SELECT * FROM tblResults";
            cmd.Connection = conn;

            DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,null);

            OleDbDataReader rdr = cmd.ExecuteReader();
        }



        


    }
}
