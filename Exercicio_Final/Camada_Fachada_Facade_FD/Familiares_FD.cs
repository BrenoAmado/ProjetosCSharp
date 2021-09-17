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
    public class Familiares_FD
    {
        Familiares_DAO objFamiliares_DAO;

        public DataTable ConsultarBd(Familiares_VO objFamiliares_VO)
        {
            try
            {
                objFamiliares_DAO = new Familiares_DAO();

                return objFamiliares_DAO.ConsultarBd(objFamiliares_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool InserirBd(Familiares_VO objFamiliares_VO)
        {
            try
            {
                objFamiliares_DAO = new Familiares_DAO();

                return objFamiliares_DAO.InserirBd(objFamiliares_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBd(Familiares_VO objFamiliares_VO)
        {
            try
            {
                objFamiliares_DAO = new Familiares_DAO();

                return objFamiliares_DAO.ExcluirBd(objFamiliares_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool AlterarBd(Familiares_VO objFamiliares_VO)
        {
            try
            {
                objFamiliares_DAO = new Familiares_DAO();

                return objFamiliares_DAO.AlterarBd(objFamiliares_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
