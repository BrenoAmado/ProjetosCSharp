using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Camada_Fachada_Facade_FD;
using Camada_Model_VO;

namespace Camada_Negocios_Business_Bll
{
    public class Preferencias_Bll
    {
        StreamReader objLeitor;
        string strLinhaLida;

        Preferencias_FD objPreferencias_FD;

        public List<string> ImportarTxt()
        {
            try
            {
                List<string> Resultado = new List<string>();

                objLeitor = new StreamReader(@"C:\CursoProgramar\Preferencias.txt");
                strLinhaLida = objLeitor.ReadLine();

                while (strLinhaLida != null)
                {
                    Resultado.Add(strLinhaLida);
                    strLinhaLida = objLeitor.ReadLine();
                }

                return Resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Importar Texto ==>" + ex.Message);
            }
            finally
            {
                objLeitor.Close();
            }
        }

        public List<string> ImportarBdConectado()
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();

                return objPreferencias_FD.ImportarBdConectado();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public List<string> ImportarBdDesconectado()
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();

                return objPreferencias_FD.ImportarBdDesconectado();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarBd(Preferencias_VO objPreferencias_VO)
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();

                return objPreferencias_FD.ConsultarBd(objPreferencias_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool InserirBd(Preferencias_VO objPreferencias_VO)
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();

                return objPreferencias_FD.InserirBd(objPreferencias_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBd(Preferencias_VO objPreferencias_VO)
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();

                return objPreferencias_FD.ExcluirBd(objPreferencias_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool AlterarBd(Preferencias_VO objPreferencias_VO)
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();

                return objPreferencias_FD.AlterarBd(objPreferencias_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public void GerarExcelAccessPorInterOp(string strNomePlanilha)
        {
            try
            {
                objPreferencias_FD = new Preferencias_FD();
                objPreferencias_FD.GerarExcelAccessPorInterOp(strNomePlanilha);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
