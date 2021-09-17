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
    public class PrefFam_DAO : DAL_DAO
    {
        PrefFam_VO objPrefFam_VO;
        OleDbCommand objCom;
        OleDbDataAdapter objAdaptador;
        DataTable objTabela;

        public override DataTable ConsultarBd(Object objparPrefFam_VO)
        {
            try
            {
                objPrefFam_VO = (PrefFam_VO)objparPrefFam_VO;

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append(" SELECT");
                strSQL.Append(" COD");
                strSQL.Append(", ID");
                strSQL.Append(", Intensidade");
                strSQL.Append(", Observacao");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_De_Familiares");

                if (objPrefFam_VO.Familiares_VO.COD > 0 && objPrefFam_VO.Preferencias_VO.ID > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" COD = :parCOD");
                    strSQL.Append(" AND");
                    strSQL.Append(" ID = :parID");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parCOD", objPrefFam_VO.Familiares_VO.COD);
                    objCom.Parameters.AddWithValue("parID", objPrefFam_VO.Preferencias_VO.ID);
                }
                else if (objPrefFam_VO.Familiares_VO.COD > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" COD = :parCOD");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parCOD", objPrefFam_VO.Familiares_VO.COD);
                }
                else if (objPrefFam_VO.Preferencias_VO.ID > 0)
                {
                    strSQL.Append(" WHERE");
                    strSQL.Append(" ID = :parID");

                    objCom = new OleDbCommand(strSQL.ToString(), getConn());
                    objCom.Parameters.AddWithValue("parID", objPrefFam_VO.Preferencias_VO.ID);
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

        public override bool InserirBd(Object objparPrefFam_VO)
        {
            try
            {
                objPrefFam_VO = (PrefFam_VO)objparPrefFam_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("INSERT");
                strSQL.Append(" INTO");
                strSQL.Append(" Preferencias_De_Familiares (");
                strSQL.Append("COD,");
                strSQL.Append(" ID,");
                strSQL.Append(" Intensidade,");
                strSQL.Append(" Observacao");
                strSQL.Append(") VALUES (");
                strSQL.Append(":parCOD,");
                strSQL.Append(" :parID,");
                strSQL.Append(" :parIntensidade,");
                strSQL.Append(" :parObservacao");
                strSQL.Append(")");

                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parCOD", objPrefFam_VO.Familiares_VO.COD);
                objCom.Parameters.AddWithValue("parID", objPrefFam_VO.Preferencias_VO.ID);
                objCom.Parameters.AddWithValue("parIntensidade", objPrefFam_VO.Intensidade);
                objCom.Parameters.AddWithValue("parObservacao", objPrefFam_VO.Observacao);


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

        public override bool ExcluirBd(Object objparPrefFam_VO)
        {
            try
            {
                objPrefFam_VO = (PrefFam_VO)objparPrefFam_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("DELETE");
                strSQL.Append(" FROM");
                strSQL.Append(" Preferencias_De_Familiares");
                strSQL.Append(" WHERE");
                strSQL.Append(" COD = :parCOD");
                strSQL.Append(" AND");
                strSQL.Append(" ID = :parID");


                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parCOD", objPrefFam_VO.Familiares_VO.COD);
                objCom.Parameters.AddWithValue("parID", objPrefFam_VO.Preferencias_VO.ID);


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

        public override bool AlterarBd(Object objparPrefFam_VO)
        {
            try
            {
                objPrefFam_VO = (PrefFam_VO)objparPrefFam_VO;

                bool Resultado = false;
                AbrirConexao();

                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("UPDATE");
                strSQL.Append(" Preferencias_De_Familiares");
                strSQL.Append(" SET");
                strSQL.Append(" Intensidade = :parIntensidade,");
                strSQL.Append(" Observacao = :parObservacao");
                strSQL.Append(" WHERE");
                strSQL.Append(" COD = :parCOD");
                strSQL.Append(" AND");
                strSQL.Append(" ID = :parID");


                objCom = new OleDbCommand(strSQL.ToString(), getConn());
                objCom.Parameters.AddWithValue("parIntensidade", objPrefFam_VO.Intensidade);
                objCom.Parameters.AddWithValue("parObservacao", objPrefFam_VO.Observacao);
                objCom.Parameters.AddWithValue("parCOD", objPrefFam_VO.Familiares_VO.COD);
                objCom.Parameters.AddWithValue("parID", objPrefFam_VO.Preferencias_VO.ID);

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
