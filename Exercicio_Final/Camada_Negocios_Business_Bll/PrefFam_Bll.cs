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
    public class PrefFam_Bll
    {
        PrefFam_FD objPrefFam_FD;

        public DataTable ConsultarBd(PrefFam_VO objPrefFam_VO)
        {
            try
            {
                objPrefFam_FD = new PrefFam_FD();

                return objPrefFam_FD.ConsultarBd(objPrefFam_VO);

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
                objPrefFam_FD = new PrefFam_FD();

                return objPrefFam_FD.InserirBd(objPrefFam_VO);

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
                objPrefFam_FD = new PrefFam_FD();

                return objPrefFam_FD.ExcluirBd(objPrefFam_VO);

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
                objPrefFam_FD = new PrefFam_FD();

                return objPrefFam_FD.AlterarBd(objPrefFam_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
