using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

//strategy (as classes filhas sempre devem conter o que tem na DAL_DAO (implementação = override))

namespace Camada_Dados_DAL_DAO
{
    public abstract class DAL_DAO : DB_DAO
    {
        public abstract DataTable ConsultarBd(Object objvo_VO);

        public abstract bool InserirBd(Object objvo_VO);

        public abstract bool ExcluirBd(Object objvo_VO);

        public abstract bool AlterarBd(Object objvo_VO);
    }
}
