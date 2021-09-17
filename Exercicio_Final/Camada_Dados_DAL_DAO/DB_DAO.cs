using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Configuration;

//singleton

namespace Camada_Dados_DAL_DAO
{
    public class DB_DAO
    {
        private static OleDbConnection objConn;

        public static OleDbConnection getConn()
        {
            if (objConn == null)
            {
                setConn();
            }
            return objConn;
        }

        public static void setConn()
        {
            objConn = new OleDbConnection(ConfigurationSettings.AppSettings["stringDeConexao"].ToString());
        }

        public static OleDbConnection ObjConn
        {
            get
            {
                if (objConn == null)
                {
                    setConn();
                }
                return objConn;
            }
        }

        public static void AbrirConexao()
        {
            if (getConn().State == System.Data.ConnectionState.Closed)
            {
                objConn.Open();
            }
        }

        public static void FecharConexao()
        {
            if (getConn().State == System.Data.ConnectionState.Open)
            {
                objConn.Close();
                objConn.Dispose();
                objConn = null;
            }
        }
    }
}
