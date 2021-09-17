using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.OleDb;

namespace Camada_Dados_DAL_DAO
{
    public class DB_DAO
    {
        private static OleDbConnection objConn;

        public static OleDbConnection ObjConn
        {
            get
            {
                if (objConn == null)
                {
                    objConn = new OleDbConnection(ConfigurationSettings.AppSettings["stringDeConexao"].ToString());
                }
                return objConn;
            }
        }

        public static void AbrirConexao()
        {
            if (ObjConn.State == System.Data.ConnectionState.Closed)
            {
                objConn.Open();
            }
        }

        public static void FecharConexao()
        {
            if (ObjConn.State == System.Data.ConnectionState.Open)
            {
                objConn.Close();
                objConn.Dispose();
                objConn = null;
            }
        }
    }
}
