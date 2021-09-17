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
    public class Clientes_DAO : DAL_DAO
    {
        OleDbCommand objCom;
        OleDbDataReader objLeitorBd;
        OleDbDataAdapter objAdaptador;
        DataTable objTabela;

        public List<string> ImportarBdConectado()
        {
            try
            {
                List<string> Resultado = new List<string>();
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" Nome");
                strSQL.Append(" FROM");
                strSQL.Append(" Clientes");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                objLeitorBd = objCom.ExecuteReader();

                while (objLeitorBd.Read())
                {
                    Resultado.Add(objLeitorBd["Nome"].ToString());
                }

                objLeitorBd.Close();
                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Importar BD Conectado ==>" +ex.Message);
            }
            finally
            {
                FecharConexao();
            }
        }

        public List<string> ImportarBdDesconectado()
        {
            try
            {
                List<string> Resultado = new List<string>();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" Nome");
                strSQL.Append(" FROM");
                strSQL.Append(" Clientes");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                foreach (DataRow itemTabela in objTabela.Rows)
                {
                    Resultado.Add(itemTabela["Nome"].ToString());
                }

                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Importar BD Desconectado ==>" + ex.Message);
            }
        }

        public override DataTable ConsultarBd(Object objVO_VO)
        {
            try
            {
                Clientes_VO objparClientes_VO = (Clientes_VO)objVO_VO;

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" ID,");
                strSQL.Append(" Nome,");
                strSQL.Append(" Descricao,");
                strSQL.Append(" Is_Active");
                strSQL.Append(" FROM");
                strSQL.Append(" Clientes");

                if (objparClientes_VO.ID > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" ID = ?");

                    objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                    objCom.Parameters.Add("?ID", System.Data.OleDb.OleDbType.BigInt);
                    objCom.Parameters["?ID"].Value = objparClientes_VO.ID;
                }

                else if (!string.IsNullOrEmpty(objparClientes_VO.Nome))
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" Nome = ?");

                    objCom = new OleDbCommand(strSQL.ToString(), ObjConn);
                    objCom.Parameters.Add("?Nome", System.Data.OleDb.OleDbType.VarChar);
                    objCom.Parameters["?Nome"].Value = objparClientes_VO.Nome;
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
                Clientes_VO objparClientes_VO = (Clientes_VO)objVO_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("INSERT INTO");
                strSQL.Append(" Clientes (");
                strSQL.Append(" Nome,");
                strSQL.Append(" Descricao,");
                strSQL.Append(" Is_Active");
                strSQL.Append(") VALUES (");
                strSQL.Append("?,");
                strSQL.Append(" ?,");
                strSQL.Append(" ?)");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objCom.Parameters.Add("?Nome", OleDbType.VarChar);
                objCom.Parameters["?Nome"].Value = objparClientes_VO.Nome;

                objCom.Parameters.Add("?Descricao", OleDbType.VarChar);
                objCom.Parameters["?Descricao"].Value = objparClientes_VO.Descricao;

                objCom.Parameters.Add("?Is_Active", OleDbType.SmallInt);
                objCom.Parameters["?Is_Active"].Value = objparClientes_VO.Is_Active;

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
                Clientes_VO objparClientes_VO = (Clientes_VO)objVO_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("DELETE");
                strSQL.Append(" FROM");
                strSQL.Append(" Clientes");
                strSQL.Append(" WHERE");
                strSQL.Append(" ID = ?");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objCom.Parameters.Add("?ID", OleDbType.BigInt);
                objCom.Parameters["?ID"].Value = objparClientes_VO.ID;

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
                Clientes_VO objparClientes_VO = (Clientes_VO)objVO_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("UPDATE");
                strSQL.Append(" Clientes");
                strSQL.Append(" SET");
                strSQL.Append(" Nome = ?,");
                strSQL.Append(" Descricao = ?,");
                strSQL.Append(" Is_Active = ?)");
                strSQL.Append(" WHERE");
                strSQL.Append(" ID = ?");


                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objCom.Parameters.Add("?Nome", OleDbType.VarChar);
                objCom.Parameters["?Nome"].Value = objparClientes_VO.Nome;

                objCom.Parameters.Add("?Descricao", OleDbType.VarChar);
                objCom.Parameters["?Descricao"].Value = objparClientes_VO.Descricao;

                objCom.Parameters.Add("?Is_Active", OleDbType.SmallInt);
                objCom.Parameters["?Is_Active"].Value = objparClientes_VO.Is_Active;

                objCom.Parameters.Add("?ID", OleDbType.BigInt);
                objCom.Parameters["?ID"].Value = objparClientes_VO.ID;

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

        public DataTable ConsultarPIePEClientes(string strIDClientesSelecionados)
        {
            try
            {
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT C.ID, C.Nome, C.Descricao, P.ID AS ID_Pedido, P.Descricao AS Descricao_Pedido, P.Estado");
                strSQL.Append(" FROM Clientes C,");
                strSQL.Append(" (SELECT I.Cliente_ID, I.ID, I.Descricao, I.Estado");
                strSQL.Append(" FROM Pedidos_Interior I");
                strSQL.Append(" UNION");
                strSQL.Append(" SELECT E.Cliente_ID, E.ID, E.Descricao, E.Estado");
                strSQL.Append(" FROM Pedidos_Exterior E) P");
                strSQL.Append(" WHERE C.ID IN ("+ strIDClientesSelecionados +") AND C.ID = P.Cliente_ID");
                strSQL.Append(" ORDER BY C.ID, P.Cliente_ID");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Consultar PI e PE de Clientes ==>" + ex.Message);
            }
        }

        public DataTable ConsultarQuantidadePIClientes(string strIDClientesSelecionados)
        {
            try
            {
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT COUNT (ID) AS Quantidade");
                strSQL.Append(" FROM Pedidos_Interior");
                strSQL.Append(" WHERE Cliente_ID IN (" + strIDClientesSelecionados + ")");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Consultar Quantidade de PI Clientes ==>" + ex.Message);
            }
        }

        public DataTable ConsultarClientesSemPE()
        {
            try
            {
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT ID, Nome, Descricao, Is_Active");
                strSQL.Append(" FROM Clientes");
                strSQL.Append(" EXCEPT");
                strSQL.Append(" SELECT C.ID, C.Nome, C.Descricao, C.Is_Active");
                strSQL.Append(" FROM Clientes C");
                strSQL.Append(" INNER JOIN Pedidos_Exterior PEX");
                strSQL.Append(" ON C.ID = PEX.Cliente_ID");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Consultar Cliente Sem PE ==>" + ex.Message);
            }
        }

        public DataTable ConsultarQuantidadePedidos()
        {
            try
            {
                StringBuilder strSQL = new StringBuilder();

                strSQL.Append("SELECT ID AS ID_Do_Cliente, Nome, Quantidade_Pedidos FROM");
                strSQL.Append(" (SELECT Convert(varchar, ID) AS ID, Nome, COUNT(*) AS Quantidade_Pedidos FROM");
                strSQL.Append(" (SELECT C.ID, C.Nome, PIN.ID AS ID_Pedido, PIN.Descricao FROM Clientes C");
                strSQL.Append(" INNER JOIN Pedidos_Interior PIN ON C.ID = PIN.Cliente_ID");
                strSQL.Append(" UNION ALL");
                strSQL.Append(" SELECT C.ID, C.Nome, PEX.ID AS ID_Pedido, PEX.Descricao FROM Clientes C");
                strSQL.Append(" INNER JOIN Pedidos_Exterior PEX ON C.ID = PEX.Cliente_ID) AS Tabelas_Juntas");
                strSQL.Append(" GROUP BY ID, Nome");
                strSQL.Append(" UNION");
                strSQL.Append(" SELECT 'XXXXXXXXX', 'Total_Geral ===>', SUM(Quantidade_Pedidos) AS Quantidade_Pedidos_Total FROM");
                strSQL.Append(" (SELECT ID, Nome, COUNT(*) AS Quantidade_Pedidos FROM");
                strSQL.Append(" (SELECT C.ID, C.Nome, PIN.ID AS ID_Pedido, PIN.Descricao FROM Clientes C");
                strSQL.Append(" INNER JOIN Pedidos_Interior PIN ON C.ID = PIN.Cliente_ID");
                strSQL.Append(" UNION ALL");
                strSQL.Append(" SELECT C.ID, C.Nome, PEX.ID AS ID_Pedido, PEX.Descricao FROM Clientes C");
                strSQL.Append(" INNER JOIN Pedidos_Exterior PEX ON C.ID = PEX.Cliente_ID) AS Tabelas_Juntas");
                strSQL.Append(" GROUP BY ID, Nome) Total) AS Geral");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Consultar Quantidade de Pedidos ==>" + ex.Message);
            }
        }
    }
}
