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
    public class Preferencias_DAO : DAL_DAO
    {
        Preferencias_VO objPreferencias_VO;
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
                strSQL.Append(" Descricao");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());

                objLeitorBd = objCom.ExecuteReader();

                while (objLeitorBd.Read())
                {
                    Resultado.Add(objLeitorBd["Descricao"].ToString());
                }

                objLeitorBd.Close();
                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Importar BD Conectado ==>" + ex.Message);
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
                strSQL.Append(" Descricao");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objAdaptador = new OleDbDataAdapter(objCom);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                foreach (DataRow itemTabela in objTabela.Rows)
                {
                    Resultado.Add(itemTabela["Descricao"].ToString());
                }

                return Resultado;

            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Importar BD Desconectado ==>" + ex.Message);
            }
        }

        public override DataTable ConsultarBd(Object objparPreferencias_VO)
        {
            try
            {
                objPreferencias_VO = (Preferencias_VO)objparPreferencias_VO; 

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" ID,");
                strSQL.Append(" Descricao");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                if (objPreferencias_VO.ID > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" ID = :parID");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parID", objPreferencias_VO.ID);
                }
                else if (!string.IsNullOrEmpty(objPreferencias_VO.Descricao))
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" Descricao = :parDescricao");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parDescricao", objPreferencias_VO.Descricao);
                }
                else
                {
                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
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

        public override bool InserirBd(Object objparPreferencias_VO)
        {
            try
            {
                objPreferencias_VO = (Preferencias_VO)objparPreferencias_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("INSERT");
                strSQL.Append(" INTO");
                strSQL.Append(" Preferencias_3 (");
                strSQL.Append("Descricao");
                strSQL.Append(") VALUES (");
                strSQL.Append(":parDescricao");
                strSQL.Append(")");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parDescricao", objPreferencias_VO.Descricao);

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

        public override bool ExcluirBd(Object objparPreferencias_VO)
        {
            try
            {
                objPreferencias_VO = (Preferencias_VO)objparPreferencias_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("DELETE");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");
                strSQL.Append(" WHERE");
                strSQL.Append(" ID = :parID");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parID", objPreferencias_VO.ID);

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

        public override bool AlterarBd(Object objparPreferencias_VO)
        {
            try
            {
                objPreferencias_VO = (Preferencias_VO)objparPreferencias_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("UPDATE");
                strSQL.Append(" Preferencias_3");
                strSQL.Append(" SET");
                strSQL.Append(" Descricao = :parDescricao");
                strSQL.Append(" WHERE");
                strSQL.Append(" ID = :parID");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parDescricao", objPreferencias_VO.Descricao);
                objCom.Parameters.AddWithValue("parID", objPreferencias_VO.ID);

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

        public void GerarExcelAccessPorInterOp(string strNomePlanilha)
        {
            try
            {
                AbrirConexao();
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" ID");
                strSQL.Append(" , Descricao");
                strSQL.Append(" INTO");
                strSQL.Append(" [EXCEL 8.0; DATABASE=" + strNomePlanilha + "].[EXPORT EXCEL]");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_3");

                objCom = new OleDbCommand(strSQL.ToString(), ObjConn);

                if (objCom.ExecuteNonQuery() < 1)
                {
                    throw new Exception("Erro ao Gerar Exportação do Access por interOp na planilha");
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Falha na Exportação do Access" + ex.Message);
            }
            finally
            {
                FecharConexao();
            }
        }

    }
}
