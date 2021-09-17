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
    public class Familiares_Bll
    {
        Familiares_FD objFamiliares_FD;

        public DataTable ConsultarBd(Familiares_VO objFamiliares_VO)
        {
            try
            {
                objFamiliares_FD = new Familiares_FD();

                return objFamiliares_FD.ConsultarBd(objFamiliares_VO);

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
                objFamiliares_FD = new Familiares_FD();

                return objFamiliares_FD.InserirBd(objFamiliares_VO);

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
                objFamiliares_FD = new Familiares_FD();

                return objFamiliares_FD.ExcluirBd(objFamiliares_VO);

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
                objFamiliares_FD = new Familiares_FD();

                return objFamiliares_FD.AlterarBd(objFamiliares_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
