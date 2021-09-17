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
    public class PE_Bll
    {
        PE_FD objPE_FD;

        public DataTable ConsultarBd(PE_VO objPE_VO)
        {
            try
            {
                objPE_FD = new PE_FD();
                return objPE_FD.ConsultarBd(objPE_VO);

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
                objPE_FD = new PE_FD();
                return objPE_FD.InserirBd(objPE_VO);

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
                objPE_FD = new PE_FD();
                return objPE_FD.ExcluirBd(objPE_VO);

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
                objPE_FD = new PE_FD();
                return objPE_FD.AlterarBd(objPE_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}

