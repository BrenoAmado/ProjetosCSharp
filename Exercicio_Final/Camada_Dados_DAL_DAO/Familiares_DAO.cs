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
    public class Familiares_DAO : DAL_DAO
    {
        Familiares_VO objFamiliares_VO;
        OleDbCommand objCom;
        OleDbDataAdapter objAdaptador;
        DataTable objTabela;

        public override DataTable ConsultarBd(Object objparFamiliares_VO)
        {
            try
            {
                objFamiliares_VO = (Familiares_VO)objparFamiliares_VO;

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("SELECT");
                strSQL.Append(" COD");
                strSQL.Append(", Nome");
                strSQL.Append(", Sexo");
                strSQL.Append(", Idade");
                strSQL.Append(", Ganho_Mensal_Total");
                strSQL.Append(", Gasto_Mensal_Total");
                strSQL.Append(", Observacao");
                strSQL.Append(" FROM");
                strSQL.Append(" Familiares");

                if (objFamiliares_VO.COD > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" COD = :parCOD");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parCOD", objFamiliares_VO.COD);
                }
                else if (!string.IsNullOrEmpty(objFamiliares_VO.Nome))
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" Nome = :parNome");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parNome", objFamiliares_VO.Nome);
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

        public override bool InserirBd(Object objparFamiliares_VO)      
        {
            try
            {
                objFamiliares_VO = (Familiares_VO)objparFamiliares_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("INSERT");
                strSQL.Append(" INTO");
                strSQL.Append(" Familiares (");
                strSQL.Append("Nome,");
                strSQL.Append(" Sexo,");
                strSQL.Append(" Idade,");
                strSQL.Append(" Ganho_Mensal_Total,");
                strSQL.Append(" Gasto_Mensal_Total,");
                strSQL.Append(" Observacao");
                strSQL.Append(") VALUES (");
                strSQL.Append(":parNome,");
                strSQL.Append(" :parSexo,");
                strSQL.Append(" :parIdade,");
                strSQL.Append(" :parGanho_Mensal_Total,");
                strSQL.Append(" :parGasto_Mensal_Total,");
                strSQL.Append(" :parObservacao");
                strSQL.Append(")");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parNome", objFamiliares_VO.Nome);
                objCom.Parameters.AddWithValue("parSexo", objFamiliares_VO.Sexo);
                objCom.Parameters.AddWithValue("parIdade", objFamiliares_VO.Idade);
                objCom.Parameters.AddWithValue("parGanho_Mensal_Total", objFamiliares_VO.Ganho);
                objCom.Parameters.AddWithValue("parGasto_Mensal_Total", objFamiliares_VO.Gasto);
                objCom.Parameters.AddWithValue("parObservacao", objFamiliares_VO.Observacao);


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

        public override bool ExcluirBd(Object objparFamiliares_VO)
        {
            try
            {
                objFamiliares_VO = (Familiares_VO)objparFamiliares_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("DELETE");
                strSQL.Append(" FROM");
                strSQL.Append(" Familiares");
                strSQL.Append(" WHERE");
                strSQL.Append(" COD = :parCOD");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parCOD", objFamiliares_VO.COD);

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

        public override bool AlterarBd (Object objparFamiliares_VO)
        {
            try
            {
                objFamiliares_VO = (Familiares_VO)objparFamiliares_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("UPDATE");
                strSQL.Append(" Familiares");
                strSQL.Append(" SET");
                strSQL.Append(" Nome = :parNome,");
                strSQL.Append(" Sexo = :parSexo,");
                strSQL.Append(" Idade = :parIdade,");
                strSQL.Append(" Ganho_Mensal_Total = :parGanho,");
                strSQL.Append(" Gasto_Mensal_Total = :parGasto,");
                strSQL.Append(" Observacao = :parObservacao");
                strSQL.Append(" WHERE");
                strSQL.Append(" COD = :parCOD");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parNome", objFamiliares_VO.Nome);
                objCom.Parameters.AddWithValue("parSexo", objFamiliares_VO.Sexo);
                objCom.Parameters.AddWithValue("parIdade", objFamiliares_VO.Idade);
                objCom.Parameters.AddWithValue("parGanho", objFamiliares_VO.Ganho);
                objCom.Parameters.AddWithValue("parGasto", objFamiliares_VO.Gasto);
                objCom.Parameters.AddWithValue("parObservacao", objFamiliares_VO.Observacao);
                objCom.Parameters.AddWithValue("parCOD", objFamiliares_VO.COD);

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
