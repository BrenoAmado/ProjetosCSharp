using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_Dados_DAL_DAO;
using Camada_Model_VO;

namespace Camada_Fachada_Facade_FD
{
    public class Preferencias_FD
    {
        Preferencias_DAO objPreferencias_DAO;

        public List<string> ImportarBdConectado()
        {
            try
            {
                objPreferencias_DAO = new Preferencias_DAO();

                return objPreferencias_DAO.ImportarBdConectado();

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
                objPreferencias_DAO = new Preferencias_DAO();

                return objPreferencias_DAO.ImportarBdDesconectado();

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
                objPreferencias_DAO = new Preferencias_DAO();

                return objPreferencias_DAO.ConsultarBd(objPreferencias_VO);

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
                objPreferencias_DAO = new Preferencias_DAO();

                return objPreferencias_DAO.InserirBd(objPreferencias_VO);

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
                objPreferencias_DAO = new Preferencias_DAO();

                return objPreferencias_DAO.ExcluirBd(objPreferencias_VO);

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
                objPreferencias_DAO = new Preferencias_DAO();

                return objPreferencias_DAO.AlterarBd(objPreferencias_VO);

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
                objPreferencias_DAO = new Preferencias_DAO();
                objPreferencias_DAO.GerarExcelAccessPorInterOp(strNomePlanilha);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
