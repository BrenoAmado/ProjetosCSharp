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
    public class PE_FD
    {
        PE_DAO objPE_DAO;

        public DataTable ConsultarBd(PE_VO objPE_VO)
        {
            try
            {
                objPE_DAO = new PE_DAO();
                return objPE_DAO.ConsultarBd(objPE_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool InserirBd(PE_VO objPE_VO)
        {
            try
            {
                objPE_DAO = new PE_DAO();
                return objPE_DAO.InserirBd(objPE_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBd(PE_VO objPE_VO)
        {
            try
            {
                objPE_DAO = new PE_DAO();
                return objPE_DAO.ExcluirBd(objPE_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool AlterarBd(PE_VO objPE_VO)
        {
            try
            {
                objPE_DAO = new PE_DAO();
                return objPE_DAO.AlterarBd(objPE_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
