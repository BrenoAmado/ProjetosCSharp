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
    public class PI_FD
    {
        PI_DAO objPI_DAO;

        public DataTable ConsultarBd(PI_VO objPI_VO)
        {
            try
            {
                objPI_DAO = new PI_DAO();
                return objPI_DAO.ConsultarBd(objPI_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool InserirBd(PI_VO objPI_VO)
        {
            try
            {
                objPI_DAO = new PI_DAO();
                return objPI_DAO.InserirBd(objPI_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBd(PI_VO objPI_VO)
        {
            try
            {
                objPI_DAO = new PI_DAO();
                return objPI_DAO.ExcluirBd(objPI_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool AlterarBd(PI_VO objPI_VO)
        {
            try
            {
                objPI_DAO = new PI_DAO();
                return objPI_DAO.AlterarBd(objPI_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
