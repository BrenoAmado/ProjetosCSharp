using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Camada_Model_VO;

namespace Camada_Dados_DAL_DAO
{
    public class PI_DAO : DAL_DAO
    {
        OleDbCommand objCom;
        OleDbDataAdapter objAdaptador;
        DataTable objTabela;

        public override DataTable ConsultarBd(Object objVO_VO)
        {
            try
            {
                PI_VO objparPI_VO = (PI_VO)objVO_VO;

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" ID,");
                strSQL.Append(" Cliente_ID,");
                strSQL.Append(" Descricao,");
                strSQL.Append(" Estado");
                strSQL.Append(" FROM");
                strSQL.Append(" Pedidos_Interior");

                if (objparPI_VO.ID > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" ID = ?");

                    objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                    objCom.Parameters.Add("?ID", System.Data.OleDb.OleDbType.BigInt);
                    objCom.Parameters["?ID"].Value = objparPI_VO.ID;
                }

                else if (objparPI_VO.Clientes_VO.ID > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" Cliente_ID = ?");

                    objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                    objCom.Parameters.Add("?Cliente_ID", System.Data.OleDb.OleDbType.BigInt);
                    objCom.Parameters["?Cliente_ID"].Value = objparPI_VO.Clientes_VO.ID;
                }

                else if (!string.IsNullOrEmpty(objparPI_VO.Descricao))
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" Descricao = ?");

                    objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                    objCom.Parameters.Add("?Descricao", System.Data.OleDb.OleDbType.VarChar);
                    objCom.Parameters["?Descricao"].Value = objparPI_VO.Descricao;
                }

                else
                {
                    objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                }

                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Consultar BD ==>" + ex.Message);
            }
        }

        public override bool InserirBd(Object objVO_VO)
        {
            try
            {
                PI_VO objparPI_VO = (PI_VO)objVO_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("INSERT INTO");
                strSQL.Append(" Pedidos_Interior (");
                strSQL.Append(" Cliente_ID,");
                strSQL.Append(" Descricao,");
                strSQL.Append(" Estado");
                strSQL.Append(") VALUES (");
                strSQL.Append("?,");
                strSQL.Append(" ?,");
                strSQL.Append(" ?)");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objCom.Parameters.Add("?Cliente_ID", OleDbType.BigInt);
                objCom.Parameters["?Cliente_ID"].Value = objparPI_VO.Clientes_VO.ID;

                objCom.Parameters.Add("?Descricao", OleDbType.VarChar);
                objCom.Parameters["?Descricao"].Value = objparPI_VO.Descricao;

                objCom.Parameters.Add("?Estado", OleDbType.SmallInt);
                objCom.Parameters["?Estado"].Value = objparPI_VO.Estado;

                if (objCom.ExecuteNonQuery() > 0)
                {
                    Resultado = true;
                }

                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Inserir BD ==>" + ex.Message);
            }
            finally
            {
                FecharConexao();
            }
        }

        public override bool ExcluirBd(Object objVO_VO)
        {
            try
            {
                PI_VO objparPI_VO = (PI_VO)objVO_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("DELETE");
                strSQL.Append(" FROM");
                strSQL.Append(" Pedidos_Interior");
                strSQL.Append(" WHERE");
                strSQL.Append(" ID = ?");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objCom.Parameters.Add("?ID", OleDbType.BigInt);
                objCom.Parameters["?ID"].Value = objparPI_VO.ID;

                if (objCom.ExecuteNonQuery() > 0)
                {
                    Resultado = true;
                }

                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Excluir BD ==>" + ex.Message);
            }
            finally
            {
                FecharConexao();
            }
        }

        public override bool AlterarBd(Object objVO_VO)
        {
            try
            {
                PI_VO objparPI_VO = (PI_VO)objVO_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("UPDATE");
                strSQL.Append(" Pedidos_Interior");
                strSQL.Append(" SET");
                strSQL.Append(" Cliente_ID = ?,");
                strSQL.Append(" Descricao = ?,");
                strSQL.Append(" Estado = ?)");
                strSQL.Append(" WHERE");
                strSQL.Append(" ID = ?");


                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objCom.Parameters.Add("?Cliente_ID", OleDbType.BigInt);
                objCom.Parameters["?Cliente_ID"].Value = objparPI_VO.Clientes_VO.ID;

                objCom.Parameters.Add("?Descricao", OleDbType.VarChar);
                objCom.Parameters["?Descricao"].Value = objparPI_VO.Descricao;

                objCom.Parameters.Add("?Estado", OleDbType.SmallInt);
                objCom.Parameters["?Estado"].Value = objparPI_VO.Estado;

                objCom.Parameters.Add("?ID", OleDbType.BigInt);
                objCom.Parameters["?ID"].Value = objparPI_VO.ID;

                if (objCom.ExecuteNonQuery() > 0)
                {
                    Resultado = true;
                }

                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Alterar BD ==>" + ex.Message);
            }
            finally
            {
                FecharConexao();
            }
        }
    }
}
