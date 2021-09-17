using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_Dados_DAL_DAO;
using Camada_Model_VO;

namespace Camada_Fachada_Facade_FD   //para objetos obsoletos. /e extensos/.
{
    public class PrefFam_FD
    {
        PrefFam_DAO objPrefFam_DAO;

        public DataTable ConsultarBd(PrefFam_VO objPrefFam_VO)
        {
            try
            {
                objPrefFam_DAO = new PrefFam_DAO();

                return objPrefFam_DAO.ConsultarBd(objPrefFam_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool InserirBd(PrefFam_VO objPrefFam_VO)
        {
            try
            {
                objPrefFam_DAO = new PrefFam_DAO();

                return objPrefFam_DAO.InserirBd(objPrefFam_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBd(PrefFam_VO objPrefFam_VO)
        {
            try
            {
                objPrefFam_DAO = new PrefFam_DAO();

                return objPrefFam_DAO.ExcluirBd(objPrefFam_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool AlterarBd(PrefFam_VO objPrefFam_VO)
        {
            try
            {
                objPrefFam_DAO = new PrefFam_DAO();

                return objPrefFam_DAO.AlterarBd(objPrefFam_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
