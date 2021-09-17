using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_Fachada_Facade_FD;
using Camada_Model_VO;

namespace Camada_Negocios_Business_Bll
{
    public class PI_Bll
    {
        PI_FD objPI_FD;

        public DataTable ConsultarBd(PI_VO objPI_VO)
        {
            try
            {
                objPI_FD = new PI_FD();
                return objPI_FD.ConsultarBd(objPI_VO);

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
                objPI_FD = new PI_FD();
                return objPI_FD.InserirBd(objPI_VO);

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
                objPI_FD = new PI_FD();
                return objPI_FD.ExcluirBd(objPI_VO);

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
                objPI_FD = new PI_FD();
                return objPI_FD.AlterarBd(objPI_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
